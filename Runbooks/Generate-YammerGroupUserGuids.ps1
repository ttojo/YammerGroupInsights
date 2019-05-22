param (
    [Parameter(Mandatory = $true)] $WebUri
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


#-------- ここからがメイン処理 --------

$LocalTargetDirectory = "C:\"
$Date = Get-Date -Format "yyyyMMdd-HHmmss"

$conn = Get-AutomationConnection -Name "AzureRunAsConnection"
Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID -ApplicationID $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint | Out-Null
Set-AzureRmCurrentStorageAccount -ResourceGroupName 'TOTOJO-STU-RG' -StorageAccountName 'yammergroupinsight' | Out-Null

Write-Verbose "ユーザー一覧を一旦ダウンロードする"
$blobName = "YammerUsers-Current.json"
Get-AzureStorageBlob -Container "json" -Blob $blobName | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null
$LocalFile = $LocalTargetDirectory + $blobName
$userList = Get-Content $LocalFile | ConvertFrom-Json

$results = @()

$userList | ForEach-Object {
	$user = $_

	$requestHashtable = @{"id" = $user.id; "email" = $user.email}
	$messagePayLoad = ConvertTo-Json $requestHashtable
	$messagePayload = [System.Text.Encoding]::UTF8.GetBytes($messagePayload)

	$response = Invoke-RestAPI -Method Post -Uri $uri -Body $messagePayload -ContentType "application/json" -RetryCount 10 -RetryInterval 30

	$results += $response.users
}

Write-Verbose "ユーザー一覧をファイル化する"
$blobName = "YammerUserGuids-Current.json"
$LocalFile = $LocalTargetDirectory + $blobName
$results | Sort-Object id -Unique | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding utf8 -FilePath $LocalFile -Force
Get-AzureStorageBlob -Container "json" -Prefix $blobName | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $LocalFile -Container "json" -Blob $blobName | Out-Null
