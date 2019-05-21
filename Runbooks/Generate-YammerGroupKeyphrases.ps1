param (
    [Parameter(Mandatory = $true)] $PrimaryApiKey
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
		[Parameter(Mandatory = $false)] [Object] $Body,
		[Parameter(Mandatory = $false)] [string] $ContentType,
        [Parameter(Mandatory = $false)] [Int32] $RetryCount = 0,
        [Parameter(Mandatory = $false)] [Int32] $RetryInterval = 10
    )

    $completed = $false
    $retries = 0

    while (-not $completed) {
        try {
            $response = Invoke-RestMethod -Uri $uri -Headers $Headers -Method $Method -Body $Body -ContentType $ContentType
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

function Get-AzureKeywords ($messageToEvaluate)
{
	$azureRegion = "japaneast"

    #define cognitive services URLs
    $keyPhraseURI = "https://$azureRegion.api.cognitive.microsoft.com/text/analytics/v2.0/keyPhrases"

    #create the JSON request
    $documents = @()
    $requestHashtable = @{"language" = "ja"; "id" = "1"; "text" = "$messageToEvaluate" };
    $documents += $requestHashtable
    $final = @{documents = $documents}
    $messagePayload = ConvertTo-Json $final
    $messagePayload = [System.Text.Encoding]::UTF8.GetBytes($messagePayload)

    #invoke the Text Analytics Keyword API
    $keywordResult = Invoke-RestAPI -Method Post -Uri $keyPhraseURI -Header @{ "Ocp-Apim-Subscription-Key" = $PrimaryApiKey } -Body $messagePayload -ContentType "application/json"  -RetryCount 10 -RetryInterval 30

	if ($keywordResult.documents.length -eq 0) {
		return ""
	}

    #return the keywords
    return $keywordResult.documents.keyPhrases
}


#-------- ここからがメイン処理 --------

$LocalTargetDirectory = "C:\"
$Date = Get-Date -Format "yyyyMMdd-HHmmss"

$conn = Get-AutomationConnection -Name "AzureRunAsConnection"
Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID -ApplicationID $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint | Out-Null
Set-AzureRmCurrentStorageAccount -ResourceGroupName 'TOTOJO-STU-RG' -StorageAccountName 'yammergroupinsight' | Out-Null

$keyPhraseHash = @{}
$messageList = @()

Write-Verbose "前回までのキーフレーズ ファイルを読み込み"
$prefix = "YammerKeyPhrases_"
Write-Verbose "Prefix: $prefix"
$blob = Get-AzureStorageBlob -Container "json" -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -First 1
if ($blob -ne $null) {
	Write-Verbose "対象ファイル: $($blob.Name)"

    Write-Verbose "ファイルを一旦ダウンロード"
    $blob | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null

    Write-Verbose "JSON ファイルをロード"
    $jsonPath = $LocalTargetDirectory + $blob.Name
	$keyPhrases = Get-Content $jsonPath | ConvertFrom-Json
	$keyPhrases | ForEach-Object {
		$keyPhraseHash.Add($_.id, [PSCustomObject]@{
			id = $_.id
			key_phrase = $_.key_phrase
		})
	}

    Write-Verbose "最近のファイルだけ残して後は削除"
    Get-AzureStorageBlob -Container "json" -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -Skip 5 | Remove-AzureStorageBlob
}

Write-Verbose "グループ メッセージを読み込み"
$prefix = "YammerMessages-Current.json"
Write-Verbose "Prefix: $prefix"
$blob = Get-AzureStorageBlob -Container "json" -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -First 1
Write-Verbose "対象ファイル: $($blob.Name)"

Write-Verbose "ファイルを一旦ダウンロード"
$blob | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null

Write-Verbose "JSON ファイルをロード"
$jsonPath = $LocalTargetDirectory + $blob.Name
$messageList = Get-Content $jsonPath | ConvertFrom-Json

$count = 0

foreach ($msg in $messageList) {
	if ($count -gt 1000) {
		break
	}

	$dt = [datetime]$msg.created_at
	if ($dt.Year -lt 2018) {
		continue
	}

	if ($keyPhraseHash.ContainsKey($msg.id) -eq $true) {
		continue
	}

	$phrase = $null

	Write-Verbose "テキスト分析"
	$keywords = Get-AzureKeywords -messageToEvaluate $msg.body.plain
	$keywords | ForEach-Object {
		if ($phrase -eq $null) {
			$phrase = $_
		} else {
			$phrase = $phrase + "," + $_
		}
	}

	$keyPhraseHash.Add($msg.id, [PSCustomObject]@{
		id = $msg.id
		key_phrase = $phrase
	})
	$count++
}


Write-Verbose "できあがったキーフレーズ ファイルを保存"
$DateBlobName = "YammerKeyPhrases_" + $Date + ".json"
$LocalFile = $LocalTargetDirectory + $DateBlobName
$keyPhraseHash.Values | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Write-Verbose "$($LocalFile) に保存しました。"

Write-Verbose "作成履歴を保存"
Set-AzureStorageBlobContent -File $LocalFile -Container "json" -Blob $DateBlobName | Out-Null

Write-Verbose "最新のキーフレーズ ファイルを保存"
$blobName = "YammerKeyPhrases-Current.json"
Get-AzureStorageBlob -Container "json" -Prefix $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container "json" -Blob $blobName | Out-Null
