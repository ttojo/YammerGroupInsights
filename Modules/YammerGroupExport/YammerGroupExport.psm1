#-----------------------------------------------------------[Functions]------------------------------------------------------------

<#
.SYNOPSIS
  Yammer REST API にアクセスするための開発者トークンを取得します。

.DESCRIPTION
  Yammer REST API にアクセスするための開発者トークンを取得します。
  開発者トークンは Azure Automation の資格情報として予め登録しておきます。

.INPUTS
  なし

.OUTPUTS
  開発者トークン文字列を返します。

.LINK
  https://developer.yammer.com/docs/test-token

.NOTES
  Version:        1.0
  Author:         Toshio Tojo
  Creation Date:  2019/7/20
#>
function Get-DeveloperToken {
    $credential = Get-AutomationPSCredential -Name "YammerDeveloper"
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($credential.Password)
    [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
}

<#
.SYNOPSIS
  REST API を呼び出します。（リトライあり）

.DESCRIPTION
  REST API を呼び出します。
  HTTP 応答が 304 または 400 ~ 599 の場合には、リトライを実施します。

.PARAMETER Uri
  Web リクエストの発行先 URI を指定します。

.PARAMETER Headers
  Web リクエストのヘッダーをハッシュテーブル形式で指定します。

.PARAMETER Method
  Web リクエストのメソッドを指定します。省略時は Get です。

.PARAMETER RetryCount
  リトライ回数を指定します。省略時は 0 (リトライなし) です。

.PARAMETER RetryInterval
  リトライの間隔 (秒) を指定します。省略時は 10 秒です。

.INPUTS
  なし

.OUTPUTS
  Web 応答を返します。(形式は呼び出した REST API により異なります。)

.NOTES
  Version:        1.0
  Author:         Toshio Tojo
  Creation Date:  2019/7/20
#>
function Invoke-RestAPI {
    param (
        [Parameter(Mandatory = $true)] [Uri] $Uri,
        [Parameter(Mandatory = $false)] [System.Collections.IDictionary] $Headers,
        [Parameter(Mandatory = $false)] [Microsoft.PowerShell.Commands.WebRequestMethod] $Method = [Microsoft.PowerShell.Commands.WebRequestMethod]::Get,
        [Parameter(Mandatory = $false)] [Int32] $RetryCount = 0,
        [Parameter(Mandatory = $false)] [Int32] $RetryInterval = 10
    )

    $completed = $false
    $retries = 0

    while (-not $completed) {
        try {
            $response = Invoke-RestMethod -Uri $uri -Headers $Headers -Method $Method
            $completed = $true
        } catch {
            $ex = $_.Exception

            if ($ex.Response -ne $null) {
                
				# 指定された回数まではリトライしてみる
                if ($retries -lt $RetryCount) {
                    $sc = [int]$ex.Response.StatusCode.Value__
                    if (($sc -eq 304) -or ((400 -le $sc) -and ($sc -le 599))) {
                        $retries++
                        Write-Verbose "Retry...($retries/$RetryCount)"
                        Start-Sleep -Seconds $RetryInterval
                        continue
                    }
                }
                $errorResponse = $ex.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($errorResponse)
                $reader.BaseStream.Position = 0
                $reader.DiscardBufferedData()
                $responseBody = $reader.ReadToEnd();

                Write-Host "Response content:`n$responseBody" -f Red
                Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
            } else {
                Write-Host "Unhanded exception:`n$ex"
            }

            #re-throw
            throw
        }
    }

    return $response
}

<#
.SYNOPSIS
  Yammer グループ情報を取得します。

.DESCRIPTION
  Yammer グループ情報を取得します。

.PARAMETER GroupId
  対象となる Yammer グループの ID を指定します。

.INPUTS
  なし

.OUTPUTS
  Yammer グループ情報を JSON 形式で返します。

.NOTES
  Version:        1.0
  Author:         Toshio Tojo
  Creation Date:  2019/7/20
#>
function Get-GroupInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] $GroupId
    )

    $apiVersion = "v1"
    $Resource = "groups/$($GroupId).json"
    $uri = "https://www.yammer.com/api/$apiVersion/$Resource"
    $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method 'Get' -RetryCount 10 -RetryInterval 30
    $response
}

<#
.SYNOPSIS
  Yammer グループのメンバーを列挙します。

.DESCRIPTION
  Yammer グループのメンバーを列挙します。

.PARAMETER GroupId
  対象となる Yammer グループの ID を指定します。

.INPUTS
  なし

.OUTPUTS
  Yammer ユーザーのリストを JSON 形式で返します。

.LINK
  https://developer.yammer.com/docs/usersin_groupidjson

.NOTES
  Version:        1.0
  Author:         Toshio Tojo
  Creation Date:  2019/7/20
#>
function Get-GroupMembers {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] $GroupId
    )

    $page = 1

    $apiVersion = "v1"
    $Resource = "users/in_group/$($GroupId).json"
    $uri = "https://www.yammer.com/api/$apiVersion/$Resource"
    $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30
    $hasNext = $response.more_available
    $members = $response.users

    while ($hasNext) {
        $page++
        $queryParams = "page=$($page)"
        $uri = "https://www.yammer.com/api/$apiVersion/$($Resource)?$queryParams"
        $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30
	    $hasNext = $response.more_available
        $members += $response.users
    }

    $members
}

<#
.SYNOPSIS
  Yammer グループのスレッドを列挙します。

.DESCRIPTION
  Yammer グループのスレッドを列挙します。

.PARAMETER GroupId
  対象となる Yammer グループの ID を指定します。

.INPUTS
  なし

.OUTPUTS
  Yammer メッセージのリストを JSON 形式で返します。

.LINK
  https://developer.yammer.com/docs/messagesin_groupgroup_id

.NOTES
  Version:        1.0
  Author:         Toshio Tojo
  Creation Date:  2019/7/20
#>
function Get-GroupThreads {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] $GroupId
    )

    $progress = 0
    $apiVersion = "v1"
    $Resource = "messages/in_group/$($GroupId).json"
    $uri = "https://www.yammer.com/api/$apiVersion/$($Resource)"
    $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30
    $hasNext = $response.meta.older_available
    $threads = $response.messages

    while ($hasNext) {
        $progress++
        if (($progress % 20) -eq 0) {
            Write-Verbose "$($progress) 件のメッセージを処理しました。"
        }
        $queryParams = "older_than=" + $response.messages[$response.messages.length - 1].id
        $uri = "https://www.yammer.com/api/$apiVersion/$($Resource)?$queryParams"
        $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30
	    $hasNext = $response.meta.older_available
        $threads += $response.messages
    }

    $threads
}

<#
.SYNOPSIS
  Yammer スレッドのメッセージを列挙します。

.DESCRIPTION
  Yammer スレッドのメッセージを列挙します。

.PARAMETER ThreadId
  対象となる Yammer スレッドの ID を指定します。

.INPUTS
  なし

.OUTPUTS
  Yammer メッセージのリストを JSON 形式で返します。

.LINK
  https://developer.yammer.com/docs/messagesin_threadthreadidjson

.NOTES
  Version:        1.0
  Author:         Toshio Tojo
  Creation Date:  2019/7/20
#>
function Get-ThreadMessages {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] $ThreadId
    )

    $apiVersion = "v1"
    $Resource = "messages/in_thread/$($ThreadId).json"
    $queryParams = "threaded=true"
    $uri = "https://www.yammer.com/api/$apiVersion/$Resource"
    $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30
    $hasNext = $response.meta.older_available
    $messages = $response.messages

    while ($hasNext) {
        $queryParams = "older_than=" + $response.messages[$response.messages.length - 1].id
        $uri = "https://www.yammer.com/api/$apiVersion/$($Resource)?$queryParams"
        $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30
        $hasNext = $response.meta.older_available
        $messages += $response.messages
    }

    $messages
}

<#
.SYNOPSIS
  Yammer メッセージにいいねしたユーザーを列挙します。

.DESCRIPTION
  Yammer メッセージにいいねしたユーザーを列挙します。

.PARAMETER MessageId
  対象となる Yammer メッセージの ID を指定します。

.INPUTS
  なし

.OUTPUTS
  Yammer ユーザーのリストを JSON 形式で返します。

.LINK
  https://developer.yammer.com/docs/usersliked_messagemessage_idjson

.NOTES
  Version:        1.0
  Author:         Toshio Tojo
  Creation Date:  2019/7/20
#>
function Get-LikedUsers {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] $MessageId
    )

    $page = 1

    $apiVersion = "v1"
    $Resource = "users/liked_message/$($MessageId).json"
    $uri = "https://www.yammer.com/api/$apiVersion/$Resource"
    $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30
    $hasNext = $response.more_available
    $members = $response.users

    while ($hasNext) {
        $page++
        $queryParams = "page=$($page)"
        $uri = "https://www.yammer.com/api/$apiVersion/$($Resource)?$queryParams"
        $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30
        $hasNext = $response.more_available
        $members += $response.users
    }

    $members
}


#----------------------------------------------------------[Declarations]----------------------------------------------------------

Write-Verbose "開発者トークンを取得します。"
$developerToken = Get-DeveloperToken
$authHeader = @{
    'Content-Type' = 'application/json'
    'Authorization' = "Bearer $developerToken"
}
