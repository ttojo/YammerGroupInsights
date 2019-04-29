param (
    [Parameter(Mandatory = $true)] $GroupId
)

$VerbosePreference = 'Continue'

function Get-DeveloperToken {
    $credential = Get-AutomationPSCredential -Name "YammerDeveloper"
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($credential.Password)
    [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
}

# Rest API の呼び出し
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
                        Write-Verbose "リトライします...($retries/$RetryCount)"
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

# グループ情報を取得
function Get-GroupInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] $GroupId
    )

    $apiVersion = "v1"
    $Resource = "groups/$($GroupId).json"
    $uri = "https://www.yammer.com/api/$apiVersion/$Resource"
    #try {
        $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method 'Get' -RetryCount 10 -RetryInterval 30

        $response

    #} catch {

    #    $ex = $_.Exception

    #    if ($ex.Response -ne $null)
    #    {
    #        $errorResponse = $ex.Response.GetResponseStream()
    #        $reader = New-Object System.IO.StreamReader($errorResponse)
    #        $reader.BaseStream.Position = 0
    #        $reader.DiscardBufferedData()
    #        $responseBody = $reader.ReadToEnd();

    #        Write-Host "Response content:`n$responseBody" -f Red
    #        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    #    }
    #    else
    #    {
    #        Write-Host "Unhanded exception:`n$ex"
    #    }

    #    break
    #}
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
    #try {
        $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30

        $members = $response.users

        while ($response.users.length -gt 0) {
            $page++
            $queryParams = "page=$($page)"
            $uri = "https://www.yammer.com/api/$apiVersion/$($Resource)?$queryParams"
            $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30
            $members += $response.users
        }

        $members

    #} catch {

    #    $ex = $_.Exception

    #    if ($ex.Response -ne $null)
    #    {
    #        $errorResponse = $ex.Response.GetResponseStream()
    #        $reader = New-Object System.IO.StreamReader($errorResponse)
    #        $reader.BaseStream.Position = 0
    #        $reader.DiscardBufferedData()
    #        $responseBody = $reader.ReadToEnd();

    #        Write-Host "Response content:`n$responseBody" -f Red
    #        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    #    }
    #    else
    #    {
    #        Write-Host "Unhanded exception:`n$ex"
    #    }

    #    break
    #}
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
    #try {
        $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30

        $threads = $response.messages

        while ($response.messages.length -gt 0) {
            $progress++
            if (($progress % 20) -eq 0) {
                Write-Verbose "$($progress) 件のメッセージを処理しました。"
                #Write-Verbose "いったん休憩します。"
                #Start-Sleep -Seconds 60
                #Write-Verbose "処理再開します。"
            }
            $queryParams = "older_than=" + $response.messages[$response.messages.length - 1].id
            $uri = "https://www.yammer.com/api/$apiVersion/$($Resource)?$queryParams"
            $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30
            $threads += $response.messages
        }

        $threads

    #} catch {

    #    $ex = $_.Exception

    #    if ($ex.Response -ne $null)
    #    {
    #        $errorResponse = $ex.Response.GetResponseStream()
    #        $reader = New-Object System.IO.StreamReader($errorResponse)
    #        $reader.BaseStream.Position = 0
    #        $reader.DiscardBufferedData()
    #        $responseBody = $reader.ReadToEnd();

    #        Write-Host "Response content:`n$responseBody" -f Red
    #        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    #    }
    #    else
    #    {
    #        Write-Host "Unhanded exception:`n$ex"
    #    }

    #    break
    #}
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
    #try {
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

    #} catch {

    #    $ex = $_.Exception

    #    if ($ex.Response -ne $null)
    #    {
    #        $errorResponse = $ex.Response.GetResponseStream()
    #        $reader = New-Object System.IO.StreamReader($errorResponse)
    #        $reader.BaseStream.Position = 0
    #        $reader.DiscardBufferedData()
    #        $responseBody = $reader.ReadToEnd();

    #        Write-Host "Response content:`n$responseBody" -f Red
    #        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    #    }
    #    else
    #    {
    #        Write-Host "Unhanded exception:`n$ex"
    #    }

    #    break
    #}
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
    #try {
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

    #} catch {

    #    $ex = $_.Exception

    #    if ($ex.Response -ne $null)
    #    {
    #        $errorResponse = $ex.Response.GetResponseStream()
    #        $reader = New-Object System.IO.StreamReader($errorResponse)
    #        $reader.BaseStream.Position = 0
    #        $reader.DiscardBufferedData()
    #        $responseBody = $reader.ReadToEnd();

    #        Write-Host "Response content:`n$responseBody" -f Red
    #        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    #    }
    #    else
    #    {
    #        Write-Host "Unhanded exception:`n$ex"
    #    }

    #    break
    #}
}

function Get-GroupStatus {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] $GroupId
    )

    # 結果格納用のテーブル作成
    Write-Verbose "グループ メンバーを取得します。"
    $membersHash = @{}
    $members = Get-GroupMembers -GroupId $groupId
    Write-Verbose "$($members.length) 人のメンバーがいます。"
    foreach($member in $members) {
        if (($member.type -eq 'user') -and ($member.state -eq 'active')) {
            $membersHash.Add($member.id, [PSCustomObject]@{
                    fullName = $member.full_name
                    jobTitle = $member.job_title
                    email = $member.email
                    messageCount = 0
                    threadCount = 0
                    responseCount = 0
                    likeCount = 0
                    likedCount = 0
                }
            )
        }
    }

    # スレッドの一覧
    Write-Verbose "スレッド一覧を作成します。"
    $threadMessages = Get-GroupThreads -GroupId $groupId
    Write-Verbose "トータル $($threadMessages.length) 件のメッセージを取得しました。"

    $threadMessages2 = $threadMessages | Sort-Object thread_id -Unique
    Write-Verbose "$($threadMessages2.length) 件のスレッドが見つかりました。"

    $progress = 0
    #$lastDate = [DateTime]::MinValue

    $threadMessages2 | ForEach-Object {
        $thread = $_
        #$now = (Get-Date)
        #$duration = New-TimeSpan -Start $lastDate -End $now
        #if ($duration.TotalMilliseconds -lt 3000) {
        #    Start-Sleep -Milliseconds (3000 - [int]$duration.TotalMilliseconds)
        #}
        #$lastDate = $now
        Get-ThreadMessages -ThreadId $thread.thread_id | ForEach-Object {
            $message = $_

            if ($membersHash.ContainsKey($message.sender_id) -eq $true) {
                $memberStatus = $membersHash[$message.sender_id]
                $memberStatus.messageCount++
                if ($message.replied_to_id -eq $null) {
                    $memberStatus.threadCount++
                } else {
                    $memberStatus.responseCount++
                }
                $memberStatus.likedCount += $message.liked_by.count
            }

            if ($message.liked_by.count -gt 0) {
                if ($message.liked_by.count -eq $message.liked_by.names.length) {
                    $message.liked_by.names | ForEach-Object {
                        $likedUser = $_
                        if ($membersHash.ContainsKey($likedUser.user_id) -eq $true) {
                            $memberStatus = $membersHash[$likedUser.user_id]
                            $memberStatus.likeCount++
                        }
                    }
                } else {
                    Get-LikedUsers -MessageId $message.id | ForEach-Object {
                        $likedUser = $_
                        if ($membersHash.ContainsKey($likedUser.id) -eq $true) {
                            $memberStatus = $membersHash[$likedUser.id]
                            $memberStatus.likeCount++
                        }
                    }
                }
            }
        }

        $progress++
        if ($VerbosePreference -eq 'Continue')
        {
            if (($progress % 10) -eq 0)
            {
                Write-Verbose "$($progress)/$($threadMessages2.length) 件のスレッドを処理しました。"
            }
        }
        #if (($progress % 500) -eq 0) {
        #    Write-Verbose "いったん休憩します。"
        #    Start-Sleep -Seconds 300
        #    Write-Verbose "処理再開します。"
        #} elseif (($progress % 200) -eq 0) {
        #    Write-Verbose "いったん休憩します。"
        #    Start-Sleep -Seconds 120
        #    Write-Verbose "処理再開します。"
        #} elseif (($progress % 100) -eq 0) {
        #    Write-Verbose "いったん休憩します。"
        #    Start-Sleep -Seconds 60
        #    Write-Verbose "処理再開します。"
        #}
    }
    if ($VerbosePreference -eq 'Continue')
    {
        Write-Verbose "$($progress)/$($threadMessages2.length) 件のスレッドを処理しました。"
    }

    $membersHash.Values | Sort-Object fullName
}


$developerToken = Get-DeveloperToken
$authHeader = @{
    'Content-Type' = 'application/json'
    'Authorization' = "Bearer $developerToken"
}

Write-Verbose "グループの情報を取得します。"
$groupInfo = Get-GroupInfo -GroupId $GroupId
Write-Verbose "===== $($groupInfo.full_name) の処理開始 ====="

$groupStatus = Get-GroupStatus -GroupId $GroupId
Write-Output $groupStatus

Write-Verbose "ファイルに保存します。"
$LocalTargetDirectory = "C:\"
$Date = Get-Date -Format "yyyyMMdd-HHmmss"
$DateBlobName = "Yammer" + $GroupId + "_" + $Date + ".csv"
$LocalFile = $LocalTargetDirectory + $DateBlobName
$groupStatus | Export-CSV -Path $LocalFile -Delimiter "," -NoTypeInformation -Force
Write-Verbose "$($LocalFile) に保存しました。"

Write-Verbose "Azure ストレージに保存します。"
$conn = Get-AutomationConnection -Name "AzureRunAsConnection"
Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID -ApplicationID $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint | Out-Null
Set-AzureRmCurrentStorageAccount -ResourceGroupName 'TOTOJO-STU-RG' -StorageAccountName 'yammergroupinsight' | Out-Null
Set-AzureStorageBlobContent -File $LocalFile -Container "csvfiles" -Blob $DateBlobName | Out-Null
Write-Verbose "Azure ストレージに保存しました。"
