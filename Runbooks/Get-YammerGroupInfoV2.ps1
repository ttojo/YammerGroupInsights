
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

# 全グループ情報を取得する
function Get-AllGroups {
    $apiVersion = "v1"
    $Resource = "groups.json"
	$page = 1
	$result = @()
	do {
		$uri = "https://www.yammer.com/api/$apiVersion/$($Resource)?page=$page"
	    $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30
		$result += $response
		$page++
	} while ($response.length -gt 0)
    $result
}


#-------- ここからがメイン処理 --------

Write-Verbose "開発者トークンを取得します。"
$developerToken = Get-DeveloperToken
$authHeader = @{
    'Content-Type' = 'application/json'
    'Authorization' = "Bearer $developerToken"
}

Write-Verbose "全グループの情報を取得します。"
$groupList = Get-AllGroups

$LocalTargetDirectory = "C:\"
$Date = Get-Date -Format "yyyyMMdd-HHmmss"

$conn = Get-AutomationConnection -Name "AzureRunAsConnection"
Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID -ApplicationID $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint | Out-Null
Set-AzureRmCurrentStorageAccount -ResourceGroupName 'TOTOJO-STU-RG' -StorageAccountName 'yammergroupinsight' | Out-Null

Write-Verbose "ファイルに保存します。"
$DateBlobName = "YammerAllGroups_" + $Date + ".json"
$LocalFile = $LocalTargetDirectory + $DateBlobName
$groupList | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Write-Verbose "$($LocalFile) に保存しました。"

Write-Verbose "Azure ストレージに保存します。"
Set-AzureStorageBlobContent -File $LocalFile -Container "json" -Blob $DateBlobName | Out-Null
Write-Verbose "Azure ストレージに保存しました。"

Get-AzureStorageBlob -Container "json" -Prefix "YammerAllGroups-Current.json" | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $ExcelPath -Container "json" -Blob "YammerAllGroups-Current.json" | Out-Null

Get-AzureStorageBlob -Container "json" -Prefix "YammerAllGroups_" | Sort-Object LastModified -Desc | Select-Object -Skip 3 | Remove-AzureStorageBlob
