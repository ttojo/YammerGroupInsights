param (
    [Parameter(Mandatory = $true)] $GroupId
)

$VerbosePreference = 'Continue'

function Get-DeveloperToken {
    $credential = Get-AutomationPSCredential -Name "YammerDeveloper"
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($credential.Password)
    [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
}

# REST API の呼び出し（リトライあり）
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

# グループ情報を取得する
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

# グループのメンバーを列挙
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

# グループのスレッドを列挙
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

# スレッドのメッセージを列挙
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

# いいねしたユーザーを列挙
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


#-------- ここからがメイン処理 --------

Write-Verbose "開発者トークンを取得します。"
$developerToken = Get-DeveloperToken
$authHeader = @{
    'Content-Type' = 'application/json'
    'Authorization' = "Bearer $developerToken"
}

$messageList = @()
$likedList = @()

Write-Verbose "グループ (ID=$groupId) の処理を始めます。"
$groupInfo = Get-GroupInfo -GroupId $groupId
Write-Verbose "グループ名は [$($groupInfo.full_name)] です。"

Write-Host "メンバー リストを作成します。"
$groupMembers = Get-GroupMembers -GroupId $groupId

Write-Verbose "スレッド一覧を作成します。"
$threadMessages = Get-GroupThreads -GroupId $groupId
Write-Verbose "トータル $($threadMessages.length) 件のメッセージを取得しました。"

$threadMessages2 = $threadMessages | Sort-Object thread_id -Unique
Write-Verbose "$($threadMessages2.length) 件のスレッドが見つかりました。"

foreach ($thread in $threadMessages2) {
	$messages = Get-ThreadMessages -ThreadId $thread.thread_id
	foreach ($message in $messages) {
		if ($message.liked_by.count -gt 0) {
			$likedUsers = Get-LikedUsers -MessageId $message.id
			#foreach ($likedUser in $likedUsers) {
			#	$likedList += [PSCustomObject]@{
			#		message_id = $message.id
			#		liked_user = $likedUser
			#	}
			#}
		}
	}
	$messageList += $messages
}

$LocalTargetDirectory = "C:\"
$Date = Get-Date -Format "yyyyMMdd-HHmmss"

$conn = Get-AutomationConnection -Name "AzureRunAsConnection"
Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID -ApplicationID $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint | Out-Null
Set-AzureRmCurrentStorageAccount -ResourceGroupName 'TOTOJO-STU-RG' -StorageAccountName 'yammergroupinsight' | Out-Null

Write-Verbose "ファイルに保存します。"
$DateBlobName = "YammerMembers_" + $GroupId + "_" + $Date + ".json"
$LocalFile = $LocalTargetDirectory + $DateBlobName
$groupMembers | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Write-Verbose "$($LocalFile) に保存しました。"

Write-Verbose "Azure ストレージに保存します。"
Set-AzureStorageBlobContent -File $LocalFile -Container "json" -Blob $DateBlobName | Out-Null
Write-Verbose "Azure ストレージに保存しました。"

Write-Verbose "ファイルに保存します。"
$DateBlobName = "YammerMessages_" + $GroupId + "_" + $Date + ".json"
$LocalFile = $LocalTargetDirectory + $DateBlobName
$messageList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Write-Verbose "$($LocalFile) に保存しました。"

Write-Verbose "Azure ストレージに保存します。"
Set-AzureStorageBlobContent -File $LocalFile -Container "json" -Blob $DateBlobName | Out-Null
Write-Verbose "Azure ストレージに保存しました。"
