#-------------------------------------------------------------------------------------------------------
#
# Yammer メッセージのキーフレーズを抽出
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
$KeyPhraseExtractionUri = Get-AutomationVariable -Name "KeyPhraseExtractionUri"
$TextAnalyticsApiKey = Get-AutomationVariable -Name "TextAnalyticsApiKey"
$GenerationsToKeep = Get-AutomationVariable -Name "GenerationsToKeep"

# PowerShell から Azure に接続し、出力先のストレージ アカウントをセットする
$conn = Get-AutomationConnection -Name "AzureRunAsConnection"
Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID -ApplicationID $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint | Out-Null
Set-AzureRmCurrentStorageAccount -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName | Out-Null

# メッセージからキーフレーズを抽出する (戻り値はキーフレーズのコレクション)
function Get-KeyPhrases ($messageToEvaluate) {
    #create the JSON request
    $documents = @()
    $requestHashtable = @{"language" = "ja"; "id" = "1"; "text" = "$messageToEvaluate"};
    $documents += $requestHashtable
    $final = @{documents = $documents}
    $messagePayload = ConvertTo-Json $final
    $messagePayload = [System.Text.Encoding]::UTF8.GetBytes($messagePayload)

    #invoke the Text Analytics Keyword API
    $keywordResult = Invoke-RestAPI -Method Post -Uri $KeyPhraseExtractionUri -Header @{ "Ocp-Apim-Subscription-Key" = $TextAnalyticsApiKey } -Body $messagePayload -ContentType "application/json"  -RetryCount 10 -RetryInterval 30

	if ($keywordResult.documents.length -eq 0) {
		return ""
	}

    #return the keywords
    return $keywordResult.documents.keyPhrases
}

$keyPhraseHash = @{}
$messageList = @()

# 前回までに処理したキーフレーズ一覧を読み込んでおく
Write-Verbose "前回までのキーフレーズ ファイルを読み込み"
$prefix = "YammerKeyPhrases_"
Write-Verbose "Prefix: $prefix"
$blob = Get-AzureStorageBlob -Container $JsonContainerName -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -First 1
if ($blob -ne $null) {
	Write-Verbose "対象ファイル: $($blob.Name)"

    Write-Verbose "ファイルを一旦ダウンロード"
    $blob | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null

    Write-Verbose "JSON ファイルをロード"
    $jsonPath = $LocalTargetDirectory + $blob.Name
	$keyPhrases = Get-Content $jsonPath | ConvertFrom-Json
	$keyPhrases | ForEach-Object {
		$keyPhraseHash.Add($_.id, [PSCustomObject]@{id = $_.id; key_phrases = $_.key_phrases})
	}

    Write-Verbose "最近のファイルだけ残して後は削除"
    Get-AzureStorageBlob -Container $JsonContainerName -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -Skip $GenerationsToKeep | Remove-AzureStorageBlob
}

# 処理対象のメッセージ一覧を読み込む
Write-Verbose "グループ メッセージを読み込み"
$prefix = "YammerMessages-Current.json"
Write-Verbose "Prefix: $prefix"
$blob = Get-AzureStorageBlob -Container $JsonContainerName -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -First 1
Write-Verbose "対象ファイル: $($blob.Name)"

Write-Verbose "ファイルを一旦ダウンロード"
$blob | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null

Write-Verbose "JSON ファイルをロード"
$jsonPath = $LocalTargetDirectory + $blob.Name
$messageList = Get-Content $jsonPath | ConvertFrom-Json

$count = 0

# 既存のキーフレーズ一覧に含まれないメッセージ (= 新規に追加されたメッセージ) のみキーフレーズ抽出を実施
foreach ($msg in $messageList) {
	if ($keyPhraseHash.ContainsKey($msg.id) -ne $true) {
		$phrases = @()

		Write-Verbose "テキスト分析"
		$keywords = Get-KeyPhrases -messageToEvaluate $msg.body.plain
		$keywords | ForEach-Object { $phrases += $_ }

		$keyPhraseHash.Add($msg.id, [PSCustomObject]@{id = $msg.id; key_phrases = $phrases})
		$count++
	}
}

# 新しく作成したキーフレーズ一覧をファイルに保存 (作業履歴用)
Write-Verbose "できあがったキーフレーズ ファイルを保存"
$DateBlobName = "YammerKeyPhrases_" + $Date + ".json"
$LocalFile = $LocalTargetDirectory + $DateBlobName
$keyPhraseHash.Values | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Write-Verbose "$($LocalFile) に保存しました。"

Write-Verbose "作成履歴を保存"
Set-AzureStorageBlobContent -File $LocalFile -Container $JsonContainerName -Blob $DateBlobName | Out-Null

# 新しく作成したキーフレーズ一覧をファイルに保存 (最新)
Write-Verbose "最新のキーフレーズ ファイルを保存"
$blobName = "YammerKeyPhrases-Current.json"
Get-AzureStorageBlob -Container $JsonContainerName -Prefix $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container $JsonContainerName -Blob $blobName | Out-Null

Write-Verbose "$($count) 件のレコードを処理しました。"
