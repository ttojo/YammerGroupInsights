#-------------------------------------------------------------------------------------------------------
#
# Yammer グループの活動状況をマージ
#
# Version:        1.0
# Author:         Toshio Tojo
# Company Name:   Microsoft Japan
# Copyright:      (c) 2019 Toshio Tojo, Microsoft Japan. All rights reserved.
# Creation Date:  2019/7/20
#
#-------------------------------------------------------------------------------------------------------

param (
	# マージ対象となる Yammer グループ ID の配列
    [Parameter(Mandatory = $true)] [string[]]$GroupIds
)

# デバッグ時に詳細メッセージを出力する場合は有効にする
# $VerbosePreference = 'Continue'

# ローカル ファイルの出力先
$LocalTargetDirectory = "C:\"

# ファイル名をユニークにするために埋め込む文字列
$Date = Get-Date -Format "yyyyMMdd-HHmmss"

# ファイル格納先の情報は Automation 変数から取得する
$ResourceGroupName = Get-AutomationVariable -Name "ResourceGroupName"
$StorageAccountName = Get-AutomationVariable -Name "StorageAccountName"
$JsonContainerName = Get-AutomationVariable -Name "JsonContainerName"
$GenerationsToKeep = Get-AutomationVariable -Name "GenerationsToKeep"

# PowerShell から Azure に接続し、出力先のストレージ アカウントをセットする
$conn = Get-AutomationConnection -Name "AzureRunAsConnection"
Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID -ApplicationID $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint | Out-Null
Set-AzureRmCurrentStorageAccount -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName | Out-Null

# マージしたデータの入れ物
$groupList = @()
$messageList = @()
$likeList = @()
$userList = @()
$memberList = @()

foreach ($groupId in $GroupIds) {
    Write-Verbose "=== グループ ID [$groupId] の処理 ==="

	# Yammer グループ情報を取得してグループ リストに追加する
    Write-Verbose "グループ情報を取得"
    $group = Get-GroupInfo -GroupId $groupId
	$groupList += $group

	# グループ メンバーをマージする
	#
	# 基本的な流れは次の通り
	# 1) 対象グループのグループ メンバー ファイルのうち、最新のファイルをローカルに一旦ダウンロード
	# 2) ダウンロードしたファイル (JSON 形式) をオブジェクトにロード
	# 3) 対象グループのグループ メンバー ファイルのうち、保存対象となる世代より古いファイルを削除
	# 上記の一連の処理はメッセージやいいねユーザーについても同様
	Write-Verbose "グループ メンバーを読み込み"
    $prefix = "YammerMembers_" + $groupId + "_"
    Write-Verbose "Prefix: $prefix"
    $blob = Get-AzureStorageBlob -Container $JsonContainerName -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -First 1
    Write-Verbose "対象ファイル: $($blob.Name)"

    Write-Verbose "ファイルを一旦ダウンロード"
    $blob | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null

    Write-Verbose "JSON ファイルをロード"
    $jsonPath = $LocalTargetDirectory + $blob.Name
	$groupMembers = Get-Content $jsonPath | ConvertFrom-Json

    Write-Verbose "最近のファイルだけ残して後は削除"
    Get-AzureStorageBlob -Container $JsonContainerName -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -Skip $GenerationsToKeep | Remove-AzureStorageBlob

	# グループ メンバーは ID 情報だけを抜粋して、それ以外は捨てる
    Write-Verbose "グループ メンバーのリストを作成"
	$groupMembers | ForEach-Object { $memberList += [PSCustomObject]@{ group_id = $groupId; user_id = $_.id } }

	# ユーザー リストにグループ メンバーを追加する (ただし、ユーザーの重複は排除しておく)
	$userList += $groupMembers
	$userList = $userList | Sort-Object id -Unique

	# グループ メッセージをマージする
	Write-Verbose "グループ メッセージを読み込み"
    $prefix = "YammerMessages_" + $groupId + "_"
    Write-Verbose "Prefix: $prefix"
    $blob = Get-AzureStorageBlob -Container $JsonContainerName -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -First 1
    Write-Verbose "対象ファイル: $($blob.Name)"

    Write-Verbose "ファイルを一旦ダウンロード"
    $blob | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null

    Write-Verbose "JSON ファイルをロード"
    $jsonPath = $LocalTargetDirectory + $blob.Name
	$groupMessages = Get-Content $jsonPath | ConvertFrom-Json
	$messageList += $groupMessages

    Write-Verbose "最近のファイルだけ残して後は削除"
    Get-AzureStorageBlob -Container $JsonContainerName -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -Skip $GenerationsToKeep | Remove-AzureStorageBlob

	# グループのいいねユーザーをマージする
	Write-Verbose "グループ いいねユーザーを読み込み"
    $prefix = "YammerLiked_" + $groupId + "_"
    Write-Verbose "Prefix: $prefix"
    $blob = Get-AzureStorageBlob -Container $JsonContainerName -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -First 1
    Write-Verbose "対象ファイル: $($blob.Name)"

    Write-Verbose "ファイルを一旦ダウンロード"
    $blob | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null

    Write-Verbose "JSON ファイルをロード"
    $jsonPath = $LocalTargetDirectory + $blob.Name
	$groupLikes = Get-Content $jsonPath | ConvertFrom-Json
	$likeList += $groupLikes

    Write-Verbose "最近のファイルだけ残して後は削除"
    Get-AzureStorageBlob -Container $JsonContainerName -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -Skip $GenerationsToKeep | Remove-AzureStorageBlob
}

# グループ一覧をマージ ファイルに保存する (ファイル名は固定)
Write-Verbose "グループ一覧をファイル化する"
$blobName = "YammerGroups-Current.json"
$LocalFile = $LocalTargetDirectory + $blobName
$groupList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Get-AzureStorageBlob -Container $JsonContainerName -Prefix $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container $JsonContainerName -Blob $blobName | Out-Null

# ユーザー一覧をマージ ファイルに保存する (ファイル名は固定)
Write-Verbose "ユーザー一覧をファイル化する"
$blobName = "YammerUsers-Current.json"
$LocalFile = $LocalTargetDirectory + $blobName
$userList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Get-AzureStorageBlob -Container $JsonContainerName -Prefix $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container $JsonContainerName -Blob $blobName | Out-Null

# グループ メンバー一覧をマージ ファイルに保存する (ファイル名は固定)
Write-Verbose "メンバー一覧をファイル化する"
$blobName = "YammerMembers-Current.json"
$LocalFile = $LocalTargetDirectory + $blobName
$memberList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Get-AzureStorageBlob -Container $JsonContainerName -Prefix $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container $JsonContainerName -Blob $blobName | Out-Null

# メッセージ一覧から未使用カラムの値を削除（高速化のため）
foreach ($msg in $messageList) {
	$msg.body.parsed = ""
	$msg.body.rich = ""
	$msg.attachments = $null
	$msg.content_excerpt = ""
}

# メッセージ一覧をマージ ファイルに保存する (ファイル名は固定)
Write-Verbose "メッセージ一覧をファイル化する"
$blobName = "YammerMessages-Current.json"
$LocalFile = $LocalTargetDirectory + $blobName
$messageList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Get-AzureStorageBlob -Container $JsonContainerName -Prefix $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container $JsonContainerName -Blob $blobName | Out-Null

# いいねユーザー一覧をマージ ファイルに保存する (ファイル名は固定)
Write-Verbose "いいね一覧をファイル化する"
$blobName = "YammerLikes-Current.json"
$LocalFile = $LocalTargetDirectory + $blobName
$likeList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Get-AzureStorageBlob -Container $JsonContainerName -Prefix $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container $JsonContainerName -Blob $blobName | Out-Null
