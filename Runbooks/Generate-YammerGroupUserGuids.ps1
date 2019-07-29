#-------------------------------------------------------------------------------------------------------
#
# Yammer ユーザーの GUID 一覧を作成
#
# Version:        1.0
# Author:         Toshio Tojo
# Company Name:   Microsoft Japan
# Copyright:      (c) 2019 Toshio Tojo, Microsoft Japan. All rights reserved.
# Creation Date:  2019/7/20
#
#-------------------------------------------------------------------------------------------------------

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
$UserGuidQueryUri = Get-AutomationVariable -Name "UserGuidQueryUri"

# ユーザーの GUID を問い合わせる (戻り値はコレクション形式となることに注意)
function Get-UserGuids ($Id, $Email) {
	# Web サービスに GUID を問い合わせ
	$requestHashtable = @{"id" = $Id; "email" = $Email}
	$messagePayLoad = ConvertTo-Json $requestHashtable
	$messagePayload = [System.Text.Encoding]::UTF8.GetBytes($messagePayload)

	$response = Invoke-RestAPI -Method Post -Uri $UserGuidQueryUri -Body $messagePayload -ContentType "application/json" -RetryCount 10 -RetryInterval 30

	# 戻り値はコレクション
	$response.users
}

# PowerShell から Azure に接続し、出力先のストレージ アカウントをセットする
$conn = Get-AutomationConnection -Name "AzureRunAsConnection"
Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID -ApplicationID $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint | Out-Null
Set-AzureRmCurrentStorageAccount -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName | Out-Null

# Yammer ユーザー一覧ファイルを読み込む
Write-Verbose "ユーザー一覧を一旦ダウンロードする"
$blobName = "YammerUsers-Current.json"
Get-AzureStorageBlob -Container $JsonContainerName -Blob $blobName | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null
$LocalFile = $LocalTargetDirectory + $blobName
$userList = Get-Content $LocalFile | ConvertFrom-Json

# Yammer ユーザー GUID 一覧ファイルを読み込む
Write-Verbose "ユーザー GUID を一旦ダウンロードする"
$blobName = "YammerUserGuids-Current.json"
Get-AzureStorageBlob -Container $JsonContainerName -Blob $blobName | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null
$LocalFile = $LocalTargetDirectory + $blobName
$guidList = Get-Content $LocalFile | ConvertFrom-Json

# Yammer ユーザー GUID 一覧はハッシュ テーブルで持つ (高速化のため)
Write-Verbose "処理高速化のために GUID リストをハッシュで持つ"
$guidHash = @{}
$guidList | ForEach-Object {
	$guidHash.Add($_.id, [PSCustomObject]@{id = $_.id; email = $_.email; guid = $_.guid})
}

# Yammer ユーザー GUID 一覧の入れ物
$results = @()

# Yammer ユーザー GUID 一覧に含まれない Yammer ユーザーの GUID を追加する
$userList | ForEach-Object {
	$user = $_

	if ($guidHash.ContainsKey($user.id)) {
		$results += $guidHash[$user.id]
	} else {
		# 未知の Yammer ユーザーの GUID を問い合わせる
		Get-UserGuids -Id $user.id -Email $user.email | ForEach-Object {
			$newGuid = [PSCustomObject]@{id = $_.id; email = $_.email; guid = $_.guid}
			$guidHash.Add($_.id, $newGuid)
			$results += $newGuid
		}
	}
}

# Yammer ユーザー GUID 一覧ファイルを作り直す
Write-Verbose "ユーザー一覧をファイル化する"
$blobName = "YammerUserGuids-Current.json"
$LocalFile = $LocalTargetDirectory + "YammerUserGuids_" + $Date + ".json"
$results | Sort-Object id -Unique | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Get-AzureStorageBlob -Container $JsonContainerName -Blob $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container $JsonContainerName -Blob $blobName | Out-Null
