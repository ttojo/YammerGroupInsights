param (
    [Parameter(Mandatory = $true)] $GroupId
)

$VerbosePreference = 'Continue'


#-------- ここからがメイン処理 --------

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

Write-Verbose "=== グループ ID [$GroupId] の処理 ==="

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


Write-Verbose "グループの開始月～終了月だけを結果に落とす"
$startDate = [DateTime]::Parse(($groupMessages | Sort-Object created_at | Select -First 1).created_at)
$lastDate = ([DateTime]::Parse(($groupMessages | Sort-Object created_at | Select -Last 1).created_at)).AddMonths(1)
$lastDate = New-Object -TypeName DateTime -ArgumentList $lastDate.Year,$lastDate.Month,1,0,0,0

# 処理高速化のためにハッシュ テーブルのネストを作成
$membersHash = @{}
$groupMembers | Where-Object { ($_.type -eq 'user') -and ($_.state -eq 'active') } | ForEach-Object {
    $member = $_
    $countsHash = @{}
    for ($dt = $startDate; $dt -lt $lastDate; $dt = $dt.AddMonths(1)) {
        $ym = $dt.ToString("yyyy 年 MM 月")
        $countsHash.Add($ym, [PSCustomObject]@{
            yyyyMM = $ym
            messageCount = 0
            threadCount = 0
            responseCount = 0
            likeCount = 0
            likedCount = 0
        })
    }
    $membersHash.Add($member.id, [PSCustomObject]@{
        fullName = $member.full_name
        jobTitle = $member.job_title
        email = $member.email
        counts = $countsHash
    })
}

# 高速化のためにいいねハッシュを作成
$likesHash = @{}
$groupLikes | ForEach-Object {
    if ($likesHash.ContainsKey($_.message_id) -eq $false) {
        $users = @()
        $likesHash.Add($_.message_id, $users)
    }
    $likesHash[$_.message_id] += $_.user_id
}

# メッセージの解析
$groupMessages | ForEach-Object {
    $message = $_

    if ($membersHash.ContainsKey($message.sender_id) -eq $true) {
        $memberStatus = $membersHash[$message.sender_id]
        $created = ([DateTime]::Parse($message.created_at)).ToString("yyyy 年 MM 月")
        $countStatus = $memberStatus.counts[$created]
        $countStatus.messageCount++
        if ($message.replied_to_id -eq $null) {
            $countStatus.threadCount++
        } else {
            $countStatus.responseCount++
        }
        $countStatus.likedCount += $message.liked_by.count
    }

    if ($likesHash.ContainsKey($message.id) -eq $true) {
        $likesHash[$message.id] | ForEach-Object {
            if ($membersHash.ContainsKey($_) -eq $true) {
                $memberStatus = $membersHash[$_]
                $created = ([DateTime]::Parse($message.created_at)).ToString("yyyy 年 MM 月")
                $countStatus = $memberStatus.counts[$created]
                $countStatus.likeCount++
            }
        }
    }
}

# 結果テーブルに吐き出す
$result = @()
$membersHash.Values | Sort-Object fullName | ForEach-Object {
    $user = $_
    $user.counts.Values | Sort-Object yyyyMM | ForEach-Object {
        $result += [PSCustomObject]@{
            fullName = $user.fullName
            jobTitle = $user.jobTitle
            email = $user.email
            yyyyMM = $_.yyyyMM
            messageCount = $_.messageCount
            threadCount = $_.threadCount
            responseCount = $_.responseCount
            likeCount = $_.likeCount
            likedCount = $_.likedCount
        }
    }
}


Write-Verbose "ファイルに保存します。"
$DateBlobName = "YammerGroup" + $GroupId + "_" + $Date + ".csv"
$LocalFile = $LocalTargetDirectory + $DateBlobName
$result | Export-CSV -Path $LocalFile -Delimiter "," -NoTypeInformation -Force -Encoding UTF8
Write-Verbose "$($LocalFile) に保存しました。"

Write-Verbose "Azure ストレージに保存します。"
$conn = Get-AutomationConnection -Name "AzureRunAsConnection"
Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID -ApplicationID $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint | Out-Null
Set-AzureRmCurrentStorageAccount -ResourceGroupName 'TOTOJO-STU-RG' -StorageAccountName 'yammergroupinsight' | Out-Null
Set-AzureStorageBlobContent -File $LocalFile -Container "csvfiles" -Blob $DateBlobName | Out-Null
Write-Verbose "Azure ストレージに保存しました。"
