﻿param (
    [Parameter(Mandatory = $true)] $GroupIdsString
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


#-------- ここからがメイン処理 --------

Write-Verbose "開発者トークンを取得します。"
$developerToken = Get-DeveloperToken
$authHeader = @{
    'Content-Type' = 'application/json'
    'Authorization' = "Bearer $developerToken"
}

Write-Verbose "Yammer グループ リスト $GroupIdsString"
$groupIds = $GroupIdsString -split ","

$LocalTargetDirectory = "C:\"
$Date = Get-Date -Format "yyyyMMdd-HHmmss"

$conn = Get-AutomationConnection -Name "AzureRunAsConnection"
Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID -ApplicationID $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint | Out-Null
Set-AzureRmCurrentStorageAccount -ResourceGroupName 'TOTOJO-STU-RG' -StorageAccountName 'yammergroupinsight' | Out-Null

$groupList = @()
$messageList = @()
$likeList = @()
$userList = @()
$memberList = @()

foreach ($groupId in $groupIds) {
    Write-Verbose "=== グループ ID [$groupId] の処理 ==="

    Write-Verbose "グループ情報を取得"
    $group = Get-GroupInfo -GroupId $groupId
	$groupList += $group


	write-Verbose "グループ メンバーを読み込み"
    $prefix = "YammerMembers_" + $groupId + "_"
    Write-Verbose "Prefix: $prefix"
    $blob = Get-AzureStorageBlob -Container "json" -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -First 1
    Write-Verbose "対象ファイル: $($blob.Name)"

    Write-Verbose "ファイルを一旦ダウンロード"
    $blob | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null

    Write-Verbose "JSON ファイルをロード"
    $jsonPath = $LocalTargetDirectory + $blob.Name
	$groupMembers = Get-Content $jsonPath | ConvertFrom-Json

    Write-Verbose "最近のファイルだけ残して後は削除"
    Get-AzureStorageBlob -Container "json" -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -Skip 5 | Remove-AzureStorageBlob


    Write-Verbose "グループ メンバーのリストを作成"
	$groupMembers | ForEach-Object { $memberList += [PSCustomObject]@{ group_id = $groupId; user_id = $_.id } }

	$userList += $groupMembers
	$userList = $userList | Sort-Object id -Unique


	write-Verbose "グループ メッセージを読み込み"
    $prefix = "YammerMessages_" + $groupId + "_"
    Write-Verbose "Prefix: $prefix"
    $blob = Get-AzureStorageBlob -Container "json" -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -First 1
    Write-Verbose "対象ファイル: $($blob.Name)"

    Write-Verbose "ファイルを一旦ダウンロード"
    $blob | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null

    Write-Verbose "JSON ファイルをロード"
    $jsonPath = $LocalTargetDirectory + $blob.Name
	$groupMessages = Get-Content $jsonPath | ConvertFrom-Json
	$messageList += $groupMessages

    Write-Verbose "最近のファイルだけ残して後は削除"
    Get-AzureStorageBlob -Container "json" -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -Skip 5 | Remove-AzureStorageBlob


	write-Verbose "グループ メッセージを読み込み"
    $prefix = "YammerLiked_" + $groupId + "_"
    Write-Verbose "Prefix: $prefix"
    $blob = Get-AzureStorageBlob -Container "json" -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -First 1
    Write-Verbose "対象ファイル: $($blob.Name)"

    Write-Verbose "ファイルを一旦ダウンロード"
    $blob | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null

    Write-Verbose "JSON ファイルをロード"
    $jsonPath = $LocalTargetDirectory + $blob.Name
	$groupLikes = Get-Content $jsonPath | ConvertFrom-Json
	$likeList += $groupLikes

    Write-Verbose "最近のファイルだけ残して後は削除"
    Get-AzureStorageBlob -Container "json" -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -Skip 5 | Remove-AzureStorageBlob
}

Write-Verbose "グループ一覧をファイル化する"
$blobName = "YammerGroups-Current.json"
$LocalFile = $LocalTargetDirectory + $blobName
$groupList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Get-AzureStorageBlob -Container "json" -Prefix $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container "json" -Blob $blobName | Out-Null

#Write-Verbose "ユーザーの GUID を取得する"
#$requestTable = @()
#$userList | ForEach-Object {
#	$requestTable += @{"id" = $_.id; "email" = $_.email}
#}
#$messagePayLoad = ConvertTo-Json $requestTable
#$messagePayload = [System.Text.Encoding]::UTF8.GetBytes($messagePayload)
#$uri = "https://prod-26.westcentralus.logic.azure.com:443/workflows/d7db68b8f2564b51a5df0f53ebc49015/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=HKQoJ4rX9E2z2Dx_5DI7F9pi-SUSLdRnULmla9KWoKE"
#$response = Invoke-RestMethod -Method Post -Uri $uri -Body $messagePayload -ContentType "application/json"
#$response.users | ForEach-Object {
#	$u = $_
#	$userList | Where-Object { $_.id -eq $u.id } | ForEach-Object {
#		$_ | Add-Member -MemberType NoteProperty -Name "guid" -Value $u.guid
#	}
#}

Write-Verbose "ユーザー一覧をファイル化する"
$blobName = "YammerUsers-Current.json"
$LocalFile = $LocalTargetDirectory + $blobName
$userList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Get-AzureStorageBlob -Container "json" -Prefix $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container "json" -Blob $blobName | Out-Null

Write-Verbose "メンバー一覧をファイル化する"
$blobName = "YammerMembers-Current.json"
$LocalFile = $LocalTargetDirectory + $blobName
$memberList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Get-AzureStorageBlob -Container "json" -Prefix $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container "json" -Blob $blobName | Out-Null

Write-Verbose "メッセージ一覧から未使用絡むの値を削除（高速化）"
foreach ($msg in $messageList) {
	$msg.body.parsed = ""
	$msg.body.rich = ""
	$msg.attachments = $null
	$msg.content_excerpt = ""
}

Write-Verbose "メッセージ一覧をファイル化する"
$blobName = "YammerMessages-Current.json"
$LocalFile = $LocalTargetDirectory + $blobName
$messageList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Get-AzureStorageBlob -Container "json" -Prefix $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container "json" -Blob $blobName | Out-Null

Write-Verbose "いいね一覧をファイル化する"
$blobName = "YammerLikes-Current.json"
$LocalFile = $LocalTargetDirectory + $blobName
$likeList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Get-AzureStorageBlob -Container "json" -Prefix $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container "json" -Blob $blobName | Out-Null
