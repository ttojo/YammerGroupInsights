param (
    [Parameter(Mandatory = $true)] $GroupId
)

$VerbosePreference = 'Continue'

Import-Module YammerGroupExport

Write-Verbose "開発者トークン = $($developerToken)"

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
			if ($message.liked_by.count -eq $message.liked_by.names.length) {
                $message.liked_by.names | ForEach-Object {
                    $likedUser = $_
					$likedList += [PSCustomObject]@{
						message_id = $message.id
						user_id = $likedUser.user_id
					}
                }
			} else {
				$likedUsers = Get-LikedUsers -MessageId $message.id
				foreach ($likedUser in $likedUsers) {
					$likedList += [PSCustomObject]@{
						message_id = $message.id
						user_id = $likedUser.id
					}
				}
			}
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

Write-Verbose "ファイルに保存します。"
$DateBlobName = "YammerLiked_" + $GroupId + "_" + $Date + ".json"
$LocalFile = $LocalTargetDirectory + $DateBlobName
$likedList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Write-Verbose "$($LocalFile) に保存しました。"

Write-Verbose "Azure ストレージに保存します。"
Set-AzureStorageBlobContent -File $LocalFile -Container "json" -Blob $DateBlobName | Out-Null
Write-Verbose "Azure ストレージに保存しました。"
