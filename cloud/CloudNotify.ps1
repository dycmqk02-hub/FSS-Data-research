# GitHub Actions에서 PC 전원 상태와 무관하게 매일 실행되는 새 글 확인 스크립트.
# 로컬 exe(src/Notify.ps1)와 확인 로직은 같지만, 상태/설정을 %AppData% 대신
# 이 저장소 안의 cloud/ 폴더 파일로 읽고 쓴다 - Actions 러너는 매번 새로 뜨는 임시 환경이라
# 이전 실행과 비교할 상태를 어딘가 영구 저장해야 하는데, 별도 DB 없이 이 저장소에 커밋하는 방식을 씀
# (워크플로우가 실행 끝에 cloud/notify-state.json 변경분을 커밋/푸시함).
# Gmail 앱 비밀번호만 GitHub Secret(GMAIL_APP_PASSWORD)에서 읽고, 그 외 설정은 notify-config.json(비밀 아님)에 둠.

$ErrorActionPreference = "Stop"
$cloudDir = $PSScriptRoot
$configPath = Join-Path $cloudDir "notify-config.json"
$statePath = Join-Path $cloudDir "notify-state.json"

. (Join-Path $cloudDir "Config.ps1")
. (Join-Path $cloudDir "Scraper.ps1")

function Get-FssMenuKey {
    param([Parameter(Mandatory)] $Entry)
    if ($Entry.Type -eq "bbs") { return "bbs_$($Entry.BbsId)" }
    if ($Entry.Type -eq "job") { return "job_$($Entry.JobPath)" }
    return "other_$($Entry.MenuNo)"
}

function Get-FssNotifiableMenus {
    return $script:FssMenuMap | Where-Object {
        $_.Type -eq "bbs" -or ($_.Type -eq "job" -and $_.SubType -eq "list")
    }
}

function Get-FssMenuCurrentItems {
    param([Parameter(Mandatory)] $Entry)

    if ($Entry.Type -eq "bbs") {
        $extra = if ($Entry.ExtraParams) { $Entry.ExtraParams } else { "" }
        $items = Get-FssBoardList -BbsId $Entry.BbsId -MenuNo $Entry.MenuNo -PageIndex 1 -ExtraParams $extra
        return $items | ForEach-Object {
            [PSCustomObject]@{ Id = $_.NttId; Title = $_.Title; Url = $_.ViewUrl }
        }
    }

    if ($Entry.Type -eq "job" -and $Entry.SubType -eq "list") {
        $extra = if ($Entry.ExtraParams) { $Entry.ExtraParams } else { "" }
        $rows = Get-FssJobListRows -JobPath $Entry.JobPath -MenuNo $Entry.MenuNo -PageIndex 1 -ExtraParams $extra
        return $rows | ForEach-Object {
            $vurl = "$script:FssBaseUrl/fss/job/$($Entry.JobPath)/view.do?menuNo=$($Entry.MenuNo)&acntnWrkSlno=$($_.Slno)"
            if ($Entry.ExtraParams) { $vurl += "&$($Entry.ExtraParams)" }
            [PSCustomObject]@{ Id = $_.Slno; Title = $_.Title; Url = $vurl }
        }
    }

    return @()
}

function Send-FssCloudNotificationEmail {
    param(
        [Parameter(Mandatory)] $Config,
        [Parameter(Mandatory)] [string]$AppPassword,
        [Parameter(Mandatory)] [System.Collections.Generic.List[object]]$NewItemsByMenu
    )

    $totalCount = ($NewItemsByMenu | ForEach-Object { $_.Items.Count } | Measure-Object -Sum).Sum
    $subject = "[FSS-DataTool] 금융감독원 새 게시글 ${totalCount}건"
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("금융감독원 업무자료(공시/회계)에 새로 등록된 게시글입니다.")
    [void]$sb.AppendLine("확인 시각: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') (GitHub Actions, PC 꺼져있어도 동작)")
    [void]$sb.AppendLine("")
    foreach ($group in $NewItemsByMenu) {
        [void]$sb.AppendLine("■ $($group.Entry.Category) > $($group.Entry.Group) > $($group.Entry.Name)")
        foreach ($item in $group.Items) {
            [void]$sb.AppendLine("  - $($item.Title)")
            [void]$sb.AppendLine("    $($item.Url)")
        }
        [void]$sb.AppendLine("")
    }
    $body = $sb.ToString()

    $mail = New-Object System.Net.Mail.MailMessage
    $mail.From = New-Object System.Net.Mail.MailAddress($Config.SenderEmail)
    foreach ($r in $Config.Recipients) { if ($r) { $mail.To.Add($r) } }
    $mail.Subject = $subject
    $mail.SubjectEncoding = [System.Text.Encoding]::UTF8
    $mail.Body = $body
    $mail.BodyEncoding = [System.Text.Encoding]::UTF8

    $smtp = New-Object System.Net.Mail.SmtpClient($Config.SmtpHost, [int]$Config.SmtpPort)
    $smtp.EnableSsl = [bool]$Config.UseSsl
    $smtp.Credentials = New-Object System.Net.NetworkCredential($Config.SenderEmail, $AppPassword)
    try { $smtp.Send($mail) } finally { $mail.Dispose(); $smtp.Dispose() }
}

# ---- 메인 ----
if (-not (Test-Path $configPath)) { throw "설정 파일을 찾을 수 없습니다: $configPath" }
$config = Get-Content -Path $configPath -Raw -Encoding UTF8 | ConvertFrom-Json

if (-not $config.Enabled) {
    Write-Output "알림이 비활성화되어 있어 확인을 건너뜁니다."
    exit 0
}

$appPassword = $env:GMAIL_APP_PASSWORD
if (-not $appPassword) { throw "GMAIL_APP_PASSWORD 환경변수(GitHub Secret)가 설정되지 않았습니다." }

$state = @{}
if (Test-Path $statePath) {
    $json = Get-Content -Path $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($p in $json.PSObject.Properties) { $state[$p.Name] = @($p.Value) }
}

$monitored = Get-FssNotifiableMenus | Where-Object { $config.MonitoredKeys -contains (Get-FssMenuKey $_) }
Write-Output "모니터링 대상: $($monitored.Count)개 메뉴"

$newItemsByMenu = New-Object System.Collections.Generic.List[object]

foreach ($entry in $monitored) {
    $key = Get-FssMenuKey $entry
    try {
        $current = @(Get-FssMenuCurrentItems -Entry $entry)
    } catch {
        Write-Output "확인 실패 ($($entry.Name)): $($_.Exception.Message)"
        continue
    }
    $currentIds = @($current | ForEach-Object { $_.Id })

    if (-not $state.ContainsKey($key)) {
        Write-Output "최초 등록: $($entry.Name) - 기준선 저장(이번 회차는 알림 없음)"
        $state[$key] = $currentIds
        continue
    }

    $prevIds = @($state[$key])
    $newOnes = @($current | Where-Object { $prevIds -notcontains $_.Id })
    if ($newOnes.Count -gt 0) {
        Write-Output "$($entry.Name): 새 글 $($newOnes.Count)건 발견"
        $newItemsByMenu.Add([PSCustomObject]@{ Entry = $entry; Items = $newOnes })
    }
    $state[$key] = $currentIds
}

($state | ConvertTo-Json -Depth 5) | Out-File -FilePath $statePath -Encoding utf8

if ($newItemsByMenu.Count -gt 0) {
    Send-FssCloudNotificationEmail -Config $config -AppPassword $appPassword -NewItemsByMenu $newItemsByMenu
    Write-Output "알림 이메일 발송 완료 ($($newItemsByMenu.Count)개 메뉴에서 새 글 발견)"
} else {
    Write-Output "새 글 없음"
}
