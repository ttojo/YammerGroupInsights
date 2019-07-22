#-------------------------------------------------------------------------------------------------------
#
# Yammer グループの活動状況をエクスポート
#
# Version:        1.0
# Author:         Toshio Tojo
# Company Name:   Microsoft Japan
# Copyright:      (c) 2019 Toshio Tojo, Microsoft Japan. All rights reserved.
# Creation Date:  2019/7/20
#
#-------------------------------------------------------------------------------------------------------

param (
    [Parameter(Mandatory = $true)] $GroupId
)

# デバッグ時に詳細メッセージを出力する場合は有効にする
# $VerbosePreference = 'Continue'

# 対象グループの名前を取得 (デバッグ用)
Write-Verbose "グループ (ID=$groupId) の処理を始めます。"
$groupInfo = Get-GroupInfo -GroupId $groupId
Write-Verbose "グループ名は [$($groupInfo.full_name)] です。"

# グループに所属するメンバーのリストを作成
Write-Host "メンバー リストを作成します。"
$groupMembers = Get-GroupMembers -GroupId $groupId

Write-Verbose "スレッド一覧を作成します。"
$threadMessages = Get-GroupThreads -GroupId $groupId
Write-Verbose "トータル $($threadMessages.length) 件のメッセージを取得しました。"

$threadMessages2 = $threadMessages | Sort-Object thread_id -Unique
Write-Verbose "$($threadMessages2.length) 件のスレッドが見つかりました。"

Write-Output "$($threadMessages.length) -> $($threadMessages2.length)"

$messageList = @()
$likedList = @()

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

$ResourceGroupName = Get-AutomationVariable -Name "ResourceGroupName"
$StorageAccountName = Get-AutomationVariable -Name "StorageAccountName"
$JsonContainerName = Get-AutomationVariable -Name "JsonContainerName"

$conn = Get-AutomationConnection -Name "AzureRunAsConnection"
Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID -ApplicationID $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint | Out-Null
Set-AzureRmCurrentStorageAccount -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName | Out-Null

Write-Verbose "ファイルに保存します。"
$DateBlobName = "YammerMembers_" + $GroupId + "_" + $Date + ".json"
$LocalFile = $LocalTargetDirectory + $DateBlobName
$groupMembers | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Write-Verbose "$($LocalFile) に保存しました。"

Write-Verbose "Azure ストレージに保存します。"
Set-AzureStorageBlobContent -File $LocalFile -Container $JsonContainerName -Blob $DateBlobName | Out-Null
Write-Verbose "Azure ストレージに保存しました。"

Write-Verbose "ファイルに保存します。"
$DateBlobName = "YammerMessages_" + $GroupId + "_" + $Date + ".json"
$LocalFile = $LocalTargetDirectory + $DateBlobName
$messageList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Write-Verbose "$($LocalFile) に保存しました。"

Write-Verbose "Azure ストレージに保存します。"
Set-AzureStorageBlobContent -File $LocalFile -Container $JsonContainerName -Blob $DateBlobName | Out-Null
Write-Verbose "Azure ストレージに保存しました。"

Write-Verbose "ファイルに保存します。"
$DateBlobName = "YammerLiked_" + $GroupId + "_" + $Date + ".json"
$LocalFile = $LocalTargetDirectory + $DateBlobName
$likedList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Write-Verbose "$($LocalFile) に保存しました。"

Write-Verbose "Azure ストレージに保存します。"
Set-AzureStorageBlobContent -File $LocalFile -Container $JsonContainerName -Blob $DateBlobName | Out-Null
Write-Verbose "Azure ストレージに保存しました。"
