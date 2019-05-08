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
    $Resource = "/groups/$($GroupId).json"
    $uri = "https://www.yammer.com/api/$apiVersion/$Resource"
    $response = Invoke-RestAPI -Uri $uri -Headers $authHeader -Method Get -RetryCount 10 -RetryInterval 30
    $response
}

# Developer Token
$developerToken = Get-DeveloperToken
$authHeader = @{
    'Content-Type' = 'application/json'
    'Authorization' = "Bearer $developerToken"
}

Write-Verbose "対象の Yammer グループ IDを取得"
$groupIdsString = Get-AutomationVariable -Name 'YammerGroupIds'
Write-Verbose "Yammer グループ リスト $groupIdsString"
$groupIds = $groupIdsString -split ","

$LocalTargetDirectory = "C:\"
$Date = Get-Date -Format "yyyyMMdd-HHmmss"
$ExcelName = "YammerGroup" + "_" + $Date + ".xlsx"
$ExcelPath = $LocalTargetDirectory + $ExcelName

$excel = Export-Excel -Path $ExcelPath -ClearSheet -PassThru -FreezeTopRow -WorksheetName "全グループ統合"
$worksheet0 = $excel.Workbook.Worksheets[1]
#$worksheet0.SheetName = "全グループ統合"
#$worksheet.View.FreezePanes(2, 1)

$worksheet0.Cells[1, 1].Value = "グループ"
$worksheet0.Cells[1, 2].Value = "フルネーム"
$worksheet0.Cells[1, 3].Value = "ジョブ タイトル"
$worksheet0.Cells[1, 4].Value = "メール アドレス"
$worksheet0.Cells[1, 5].Value = "年月"
$worksheet0.Cells[1, 6].Value = "投稿数"
$worksheet0.Cells[1, 7].Value = "作成スレッド数"
$worksheet0.Cells[1, 8].Value = "返信数"
$worksheet0.Cells[1, 9].Value = "いいねした数"
$worksheet0.Cells[1, 10].Value = "いいねされた数"

$worksheet0.Column(1).Width = 25
$worksheet0.Column(2).Width = 25
$worksheet0.Column(3).Width = 25
$worksheet0.Column(4).Width = 25
$worksheet0.Column(5).Width = 13
$worksheet0.Column(6).Width = 13
$worksheet0.Column(7).Width = 13
$worksheet0.Column(8).Width = 13
$worksheet0.Column(9).Width = 13
$worksheet0.Column(10).Width = 13

$row0 = 2

Write-Verbose "ストレージに接続"
$conn = Get-AutomationConnection -Name "AzureRunAsConnection"
Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID -ApplicationID $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint | Out-Null
Set-AzureRmCurrentStorageAccount -ResourceGroupName 'TOTOJO-STU-RG' -StorageAccountName 'yammergroupinsight' | Out-Null

foreach ($groupId in $groupIds) {
    Write-Verbose "=== グループ ID [$groupId] の処理 ==="

    Write-Verbose "グループ情報を取得"
    $group = Get-GroupInfo -GroupId $groupId

    # シート名の最大が 31 文字なので、グループ名が長いやつは切ります
    $sheetName = if ($group.full_name.length -le 31) { $group.full_name } else { $group.full_name.Substring(0, 31) }
    $worksheet = Add-WorkSheet -ExcelPackage $excel -ClearSheet -WorksheetName $sheetName
    $worksheet.View.FreezePanes(3, 1)

    # シートの左上にグループ名を入れてみる
    $worksheet.Cells[1, 1].Value = $group.full_name
    $worksheet.Cells[1, 1].Style.Font.Size = 18
    $worksheet.Cells[1, 1].Style.Font.Bold = $true

    $worksheet.Cells[2, 1].Value = "フルネーム"
    $worksheet.Cells[2, 2].Value = "ジョブ タイトル"
    $worksheet.Cells[2, 3].Value = "メール アドレス"
    $worksheet.Cells[2, 4].Value = "年月"
    $worksheet.Cells[2, 5].Value = "投稿数"
    $worksheet.Cells[2, 6].Value = "作成スレッド数"
    $worksheet.Cells[2, 7].Value = "返信数"
    $worksheet.Cells[2, 8].Value = "いいねした数"
    $worksheet.Cells[2, 9].Value = "いいねされた数"

    $worksheet.Column(1).Width = 25
    $worksheet.Column(2).Width = 25
    $worksheet.Column(3).Width = 25
    $worksheet.Column(4).Width = 13
    $worksheet.Column(5).Width = 13
    $worksheet.Column(6).Width = 13
    $worksheet.Column(7).Width = 13
    $worksheet.Column(8).Width = 13
    $worksheet.Column(9).Width = 13

    $prefix = "YammerGroup" + $groupId
    Write-Verbose "Prefix: $prefix"
    $blob = Get-AzureStorageBlob -Container "csvfiles" -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -First 1
    Write-Verbose "対象ファイル: $($blob.Name)"

    Write-Verbose "ファイルを一旦ダウンロード"
    $blob | Get-AzureStorageBlobContent -Destination $LocalTargetDirectory | Out-Null

    Write-Verbose "CSV ファイルをロード"
    $csvPath = $LocalTargetDirectory + $blob.Name
    $groupStatus = Import-Csv $csvPath -Encoding UTF8
    #$groupStatus | ft

    Write-Verbose "最近のファイルだけ残して後は削除"
    Get-AzureStorageBlob -Container "csvfiles" -Prefix $prefix | Sort-Object LastModified -Desc | Select-Object -Skip 5 | Remove-AzureStorageBlob

    for ($row = 0; $row -lt $groupStatus.length; $row++) {
        $worksheet.Cells[(3 + $row), 1].Value = $groupStatus[$row].fullName
        $worksheet.Cells[(3 + $row), 2].Value = $groupStatus[$row].jobTitle
        $worksheet.Cells[(3 + $row), 3].Value = $groupStatus[$row].email
        $worksheet.Cells[(3 + $row), 4].Value = $groupStatus[$row].yyyyMM
        $worksheet.Cells[(3 + $row), 5].Value = [int]$groupStatus[$row].messageCount
        $worksheet.Cells[(3 + $row), 6].Value = [int]$groupStatus[$row].threadCount
        $worksheet.Cells[(3 + $row), 7].Value = [int]$groupStatus[$row].responseCount
        $worksheet.Cells[(3 + $row), 8].Value = [int]$groupStatus[$row].likeCount
        $worksheet.Cells[(3 + $row), 9].Value = [int]$groupStatus[$row].likedCount

        $worksheet0.Cells[($row0 + $row), 1].Value = $group.full_name
        for ($col = 1; $col -le 9; $col++) {
            $worksheet0.Cells[($row0 + $row), ($col + 1)].Value = $worksheet.Cells[(3 + $row), $col].Value
        }
    }
    $row0 += $groupStatus.length

    $table = Add-ExcelTable -PassThru -Range $worksheet.Cells[2, 1, (2 + $groupStatus.length), 9] -TableName "Table$groupId" -TableStyle Medium1 `
        -ShowHeader -ShowFilter -ShowRowStripes:$true -ShowTotal:$true
    $table.Columns[4].TotalsRowFunction = "sum"
    $table.Columns[5].TotalsRowFunction = "sum"
    $table.Columns[6].TotalsRowFunction = "sum"
    $table.Columns[7].TotalsRowFunction = "sum"
    $table.Columns[8].TotalsRowFunction = "sum"
}

$table = Add-ExcelTable -PassThru -Range $worksheet0.Cells[1, 1, $row0, 10] -TableName "Table0" -TableStyle Medium1 `
    -ShowHeader -ShowFilter -ShowRowStripes:$true -ShowTotal:$false

Close-ExcelPackage -ExcelPackage $excel -Show:$false

Set-AzureStorageBlobContent -File $ExcelPath -Container "reports" -Blob $ExcelName | Out-Null

Get-AzureStorageBlob -Container "reports" -Prefix "Yammer-Current.xlsx" | Remove-AzureStorageBlob
Set-AzureStorageBlobContent -File $ExcelPath -Container "reports" -Blob "Yammer-Current.xlsx" | Out-Null

Get-AzureStorageBlob -Container "reports" -Prefix "YammerGroup" | Sort-Object LastModified -Desc | Select-Object -Skip 3 | Remove-AzureStorageBlob
