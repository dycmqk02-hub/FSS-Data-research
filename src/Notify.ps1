# 새 글 알림 기능
# - 모니터링 대상: Type="bbs" 전체 + Type="job" and SubType="list"(accnutAdtorInfo형).
#   Type="job" and SubType="table"(cprCoreInfo형)은 게시글 단위가 아니라 표 데이터라 "새 글" 개념이 없어 제외.
# - 설정/상태 파일은 %AppData%\FSS-DataTool\ 에 저장(exe는 단일 파일이라 설정을 exe 밖에 보관해야 함).
# - 주의: AppPassword는 현재 평문 JSON으로 저장됨. 사용자 개인 PC 계정 폴더 안에만 저장되지만,
#   보안이 중요하면 Gmail 앱 비밀번호를 전용으로 새로 발급해서 쓰고 다른 용도로 재사용하지 않을 것.

function Get-FssNotifyDir {
    $dir = Join-Path ([Environment]::GetFolderPath("ApplicationData")) "FSS-DataTool"
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    return $dir
}

function Get-FssNotifyConfigPath { Join-Path (Get-FssNotifyDir) "notify-config.json" }
function Get-FssNotifyStatePath { Join-Path (Get-FssNotifyDir) "notify-state.json" }
function Get-FssNotifyLogPath { Join-Path (Get-FssNotifyDir) "notify-log.txt" }

function Get-FssNotifyTaskName { "FSS-DataTool-NotifyCheck" }

function Get-FssMergedScriptPath {
    try {
        $exeDir = Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    } catch {
        $exeDir = $PSScriptRoot
    }
    return Join-Path $exeDir "FSS-DataTool-merged.ps1"
}

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

function Get-FssNotifyConfig {
    $path = Get-FssNotifyConfigPath
    $default = [PSCustomObject]@{
        Enabled       = $false
        SenderEmail   = ""
        AppPassword   = ""
        SmtpHost      = "smtp.gmail.com"
        SmtpPort      = 587
        UseSsl        = $true
        Recipients    = @()
        MonitoredKeys = @()
        IntervalHours = 24
    }
    if (-not (Test-Path $path)) { return $default }

    try {
        $loaded = Get-Content -Path $path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        return $default
    }

    foreach ($prop in $default.PSObject.Properties.Name) {
        if (-not ($loaded.PSObject.Properties.Name -contains $prop)) {
            $loaded | Add-Member -MemberType NoteProperty -Name $prop -Value $default.$prop
        }
    }
    return $loaded
}

function Save-FssNotifyConfig {
    param([Parameter(Mandatory)] $Config)
    $path = Get-FssNotifyConfigPath
    ($Config | ConvertTo-Json -Depth 5) | Out-File -FilePath $path -Encoding utf8
}

function Get-FssNotifyState {
    $path = Get-FssNotifyStatePath
    $ht = @{}
    if (-not (Test-Path $path)) { return $ht }
    try {
        $json = Get-Content -Path $path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        return $ht
    }
    foreach ($p in $json.PSObject.Properties) { $ht[$p.Name] = @($p.Value) }
    return $ht
}

function Save-FssNotifyState {
    param([Parameter(Mandatory)] [hashtable]$State)
    $path = Get-FssNotifyStatePath
    ($State | ConvertTo-Json -Depth 5) | Out-File -FilePath $path -Encoding utf8
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

function Send-FssNotificationEmail {
    param(
        [Parameter(Mandatory)] $Config,
        [System.Collections.Generic.List[object]]$NewItemsByMenu = (New-Object System.Collections.Generic.List[object]),
        [switch]$TestOnly
    )

    if ($NewItemsByMenu.Count -eq 0) {
        $subject = "[FSS-DataTool] 테스트 메일"
        $body = "알림 설정이 정상적으로 동작합니다. 현재는 새로 등록된 게시글이 없습니다.`r`n확인 시각: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    } else {
        $totalCount = ($NewItemsByMenu | ForEach-Object { $_.Items.Count } | Measure-Object -Sum).Sum
        $subject = "[FSS-DataTool] 금융감독원 새 게시글 ${totalCount}건"
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("금융감독원 업무자료(공시/회계)에 새로 등록된 게시글입니다.")
        [void]$sb.AppendLine("확인 시각: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
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
    }

    $mail = New-Object System.Net.Mail.MailMessage
    $mail.From = New-Object System.Net.Mail.MailAddress($Config.SenderEmail)
    foreach ($r in $Config.Recipients) {
        if ($r) { $mail.To.Add($r) }
    }
    $mail.Subject = $subject
    $mail.SubjectEncoding = [System.Text.Encoding]::UTF8
    $mail.Body = $body
    $mail.BodyEncoding = [System.Text.Encoding]::UTF8

    $smtp = New-Object System.Net.Mail.SmtpClient($Config.SmtpHost, [int]$Config.SmtpPort)
    $smtp.EnableSsl = [bool]$Config.UseSsl
    $smtp.Credentials = New-Object System.Net.NetworkCredential($Config.SenderEmail, $Config.AppPassword)
    try {
        $smtp.Send($mail)
    } finally {
        $mail.Dispose()
        $smtp.Dispose()
    }
}

function Invoke-FssNotificationCheck {
    param(
        [switch]$SendTestMail,
        [scriptblock]$OnProgress = $null
    )

    function Write-NotifyLog {
        param([string]$Msg)
        $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Msg"
        if ($OnProgress) { & $OnProgress $line }
        try { Add-Content -Path (Get-FssNotifyLogPath) -Value $line -Encoding utf8 } catch {}
    }

    $config = Get-FssNotifyConfig
    if (-not $config.Enabled) {
        Write-NotifyLog "알림이 비활성화되어 있어 확인을 건너뜁니다."
        return
    }
    if (-not $config.SenderEmail -or -not $config.Recipients -or $config.Recipients.Count -eq 0) {
        Write-NotifyLog "발신/수신 이메일이 설정되지 않아 확인을 건너뜁니다."
        return
    }

    $state = Get-FssNotifyState
    $monitored = Get-FssNotifiableMenus | Where-Object { $config.MonitoredKeys -contains (Get-FssMenuKey $_) }

    if ($monitored.Count -eq 0) {
        Write-NotifyLog "모니터링 대상으로 선택된 메뉴가 없습니다."
        return
    }

    $newItemsByMenu = New-Object System.Collections.Generic.List[object]

    foreach ($entry in $monitored) {
        $key = Get-FssMenuKey $entry
        try {
            $current = @(Get-FssMenuCurrentItems -Entry $entry)
        } catch {
            Write-NotifyLog "확인 실패 ($($entry.Name)): $($_.Exception.Message)"
            continue
        }
        $currentIds = @($current | ForEach-Object { $_.Id })

        if (-not $state.ContainsKey($key)) {
            Write-NotifyLog "최초 등록: $($entry.Name) - 기준선 저장(이번 회차는 알림 없음)"
            $state[$key] = $currentIds
            continue
        }

        $prevIds = @($state[$key])
        $newOnes = @($current | Where-Object { $prevIds -notcontains $_.Id })
        if ($newOnes.Count -gt 0) {
            Write-NotifyLog "$($entry.Name): 새 글 $($newOnes.Count)건 발견"
            $newItemsByMenu.Add([PSCustomObject]@{ Entry = $entry; Items = $newOnes })
        }
        $state[$key] = $currentIds
    }

    Save-FssNotifyState -State $state

    if ($newItemsByMenu.Count -gt 0) {
        Send-FssNotificationEmail -Config $config -NewItemsByMenu $newItemsByMenu
        Write-NotifyLog "알림 이메일 발송 완료 (모니터링 대상 메뉴 중 $($newItemsByMenu.Count)곳에서 새 글 발견)"
    } elseif ($SendTestMail) {
        Send-FssNotificationEmail -Config $config -NewItemsByMenu (New-Object System.Collections.Generic.List[object]) -TestOnly
        Write-NotifyLog "테스트 메일 발송 완료(현재 새 글 없음)"
    } else {
        Write-NotifyLog "새 글 없음"
    }
}

function Register-FssNotifyTask {
    param(
        [Parameter(Mandatory)] [string]$ScriptPath,
        [int]$IntervalHours = 24
    )

    $taskName = Get-FssNotifyTaskName
    $psExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    $action = "`"$psExe`" -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`" -CheckOnly"

    if ($IntervalHours -ge 24) {
        & schtasks.exe /Create /TN $taskName /TR $action /SC DAILY /ST 09:00 /F | Out-Null
    } else {
        & schtasks.exe /Create /TN $taskName /TR $action /SC HOURLY /MO $IntervalHours /F | Out-Null
    }

    if ($LASTEXITCODE -ne 0) {
        throw "작업 스케줄러 등록 실패(schtasks.exe 종료코드 $LASTEXITCODE)"
    }
    return $taskName
}

function Unregister-FssNotifyTask {
    $taskName = Get-FssNotifyTaskName
    & schtasks.exe /Delete /TN $taskName /F 2>$null | Out-Null
}

function Test-FssNotifyTaskRegistered {
    $taskName = Get-FssNotifyTaskName
    & schtasks.exe /Query /TN $taskName 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0)
}
