Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-FssGui {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

    $defaultDownloadRoot = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "FSS-Downloads"

    $form = New-Object System.Windows.Forms.Form
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
    $form.Text = "금융감독원 업무자료(공시/회계) 조회·다운로드"
    $form.ClientSize = New-Object System.Drawing.Size(1080, 760)
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = New-Object System.Drawing.Size(760, 480)

    # ---- 레이아웃 상수 ----
    $TREE_WIDTH = 270
    $TOP_HEIGHT = 122
    $BTNROW_HEIGHT = 48
    $LOG_HEIGHT = 110
    $BOTTOM_HEIGHT = $BTNROW_HEIGHT + $LOG_HEIGHT

    # ---- 좌측 트리 ----
    $tree = New-Object System.Windows.Forms.TreeView

    $catNodes = @{}
    $grpNodes = @{}
    foreach ($entry in $script:FssMenuMap) {
        if (-not $catNodes.ContainsKey($entry.Category)) {
            $n = New-Object System.Windows.Forms.TreeNode($entry.Category)
            $tree.Nodes.Add($n) | Out-Null
            $catNodes[$entry.Category] = $n
        }
        $catNode = $catNodes[$entry.Category]
        $grpKey = "$($entry.Category)/$($entry.Group)"
        if (-not $grpNodes.ContainsKey($grpKey)) {
            $g = New-Object System.Windows.Forms.TreeNode($entry.Group)
            $catNode.Nodes.Add($g) | Out-Null
            $grpNodes[$grpKey] = $g
        }
        $grpNode = $grpNodes[$grpKey]
        $itemNode = New-Object System.Windows.Forms.TreeNode($entry.Name)
        $itemNode.Tag = $entry
        $grpNode.Nodes.Add($itemNode) | Out-Null
    }
    $tree.ExpandAll()

    # ---- 상단 툴바 패널 ----
    $topPanel = New-Object System.Windows.Forms.Panel

    $lblCurrent = New-Object System.Windows.Forms.Label
    $lblCurrent.Left = 10; $lblCurrent.Top = 8; $lblCurrent.Width = 800; $lblCurrent.Height = 26
    $lblCurrent.Font = New-Object System.Drawing.Font("맑은 고딕", 11, [System.Drawing.FontStyle]::Bold)
    $lblCurrent.Text = "왼쪽 트리에서 항목을 선택하세요"

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Left = 10; $txtSearch.Top = 42; $txtSearch.Width = 260

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Left = 280; $btnSearch.Top = 40; $btnSearch.Width = 70
    $btnSearch.Text = "조회"

    $btnPrev = New-Object System.Windows.Forms.Button
    $btnPrev.Left = 360; $btnPrev.Top = 40; $btnPrev.Width = 60
    $btnPrev.Text = "◀ 이전"

    $btnNext = New-Object System.Windows.Forms.Button
    $btnNext.Left = 425; $btnNext.Top = 40; $btnNext.Width = 60
    $btnNext.Text = "다음 ▶"

    $lblPage = New-Object System.Windows.Forms.Label
    $lblPage.Left = 495; $lblPage.Top = 46; $lblPage.Width = 150
    $lblPage.Text = "1페이지"

    $lblRange = New-Object System.Windows.Forms.Label
    $lblRange.Left = 10; $lblRange.Top = 78; $lblRange.Width = 220
    $lblRange.Text = "결산연월(회계법인 통합조회 전용, 비우면 전체):"

    $txtSdate = New-Object System.Windows.Forms.TextBox
    $txtSdate.Left = 230; $txtSdate.Top = 76; $txtSdate.Width = 70
    $txtSdate.Text = ""

    $lblRangeDash = New-Object System.Windows.Forms.Label
    $lblRangeDash.Left = 304; $lblRangeDash.Top = 78; $lblRangeDash.Width = 12
    $lblRangeDash.Text = "~"

    $txtEdate = New-Object System.Windows.Forms.TextBox
    $txtEdate.Left = 318; $txtEdate.Top = 76; $txtEdate.Width = 70
    $txtEdate.Text = ""

    $topPanel.Controls.AddRange(@($lblCurrent, $txtSearch, $btnSearch, $btnPrev, $btnNext, $lblPage, $lblRange, $txtSdate, $lblRangeDash, $txtEdate))

    # ---- 다운로드 버튼 줄 ----
    $btnPanel = New-Object System.Windows.Forms.Panel

    $btnDownloadSelected = New-Object System.Windows.Forms.Button
    $btnDownloadSelected.Left = 10; $btnDownloadSelected.Top = 8; $btnDownloadSelected.Width = 150; $btnDownloadSelected.Height = 30
    $btnDownloadSelected.Text = "선택 항목 다운로드"

    $btnDownloadSeries = New-Object System.Windows.Forms.Button
    $btnDownloadSeries.Left = 170; $btnDownloadSeries.Top = 8; $btnDownloadSeries.Width = 200; $btnDownloadSeries.Height = 30
    $btnDownloadSeries.Text = "이 게시판 전체(시리즈) 다운로드"

    $chkPdf = New-Object System.Windows.Forms.CheckBox
    $chkPdf.Left = 380; $chkPdf.Top = 12; $chkPdf.Width = 260
    $chkPdf.Text = "HWP→PDF 자동 변환 (한글 설치 시)"
    $chkPdf.Checked = (Test-HwpInstalled)
    if (-not (Test-HwpInstalled)) { $chkPdf.Enabled = $false; $chkPdf.Text = "HWP→PDF 변환 (한글 미설치로 비활성)" }

    $btnNotifySettings = New-Object System.Windows.Forms.Button
    $btnNotifySettings.Left = 650; $btnNotifySettings.Top = 8; $btnNotifySettings.Width = 110; $btnNotifySettings.Height = 30
    $btnNotifySettings.Text = "알림 설정"

    $btnPanel.Controls.AddRange(@($btnDownloadSelected, $btnDownloadSeries, $chkPdf, $btnNotifySettings))

    # ---- 로그 ----
    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Multiline = $true
    $txtLog.ScrollBars = "Vertical"
    $txtLog.ReadOnly = $true

    # ---- 중앙 그리드 ----
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.ReadOnly = $true
    $grid.SelectionMode = "FullRowSelect"
    $grid.MultiSelect = $true
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.Columns.Add("Title", "제목") | Out-Null
    $grid.Columns.Add("NttId", "관리번호") | Out-Null
    $grid.Columns["NttId"].Visible = $false
    $grid.Columns["Title"].FillWeight = 100

    $form.Controls.AddRange(@($tree, $topPanel, $btnPanel, $txtLog, $grid))

    # ---- 수동 레이아웃 (Dock 대신 Resize 이벤트로 직접 좌표 계산) ----
    $doLayout = {
        $cw = $form.ClientSize.Width
        $ch = $form.ClientSize.Height

        $tree.SetBounds(0, 0, $TREE_WIDTH, $ch)
        $topPanel.SetBounds($TREE_WIDTH, 0, $cw - $TREE_WIDTH, $TOP_HEIGHT)

        $gridHeight = $ch - $TOP_HEIGHT - $BOTTOM_HEIGHT
        if ($gridHeight -lt 50) { $gridHeight = 50 }
        $grid.SetBounds($TREE_WIDTH, $TOP_HEIGHT, $cw - $TREE_WIDTH, $gridHeight)

        $btnPanel.SetBounds($TREE_WIDTH, $TOP_HEIGHT + $gridHeight, $cw - $TREE_WIDTH, $BTNROW_HEIGHT)
        $txtLog.SetBounds($TREE_WIDTH, $TOP_HEIGHT + $gridHeight + $BTNROW_HEIGHT, $cw - $TREE_WIDTH, $LOG_HEIGHT)

        foreach ($c in @($tree, $topPanel, $grid, $btnPanel, $txtLog)) {
            $c.Invalidate()
            $c.Update()
        }
        $form.Refresh()
    }

    $form.Add_Resize($doLayout)
    $form.Add_Shown($doLayout)
    & $doLayout

    function Write-Log {
        param([string]$msg)
        $txtLog.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $msg`r`n")
        [System.Windows.Forms.Application]::DoEvents()
    }

    $script:CurrentEntry = $null
    $script:CurrentPage = 1
    $script:JobRowFiles = @{}

    function Load-CurrentBoard {
        if (-not $script:CurrentEntry) { return }
        $entry = $script:CurrentEntry
        $grid.Rows.Clear()
        $script:JobRowFiles = @{}

        if ($entry.Type -eq "info") {
            Write-Log "정보성 페이지입니다 (목록/첨부파일 없음). 브라우저에서 직접 확인: https://www.fss.or.kr/fss/main/contents.do?menuNo=$($entry.MenuNo)"
            $lblPage.Text = "-"
            return
        }

        if ($entry.Type -eq "job" -and $entry.SubType -eq "table") {
            Write-Log "'$($entry.Name)'은 결과가 가로로 긴 데이터 표(연번~결산년월)로 제공되어 화면 목록 대신 엑셀로만 내려받습니다. 회계법인명(검색창)/결산연월 범위(선택)를 입력한 뒤 [엑셀 다운로드]를 누르세요."
            $lblPage.Text = "-"
            return
        }

        if ($entry.Type -eq "job" -and $entry.SubType -eq "list") {
            Write-Log "조회 중: $($entry.Name) ($($script:CurrentPage)페이지)"
            try {
                $extra = if ($entry.ExtraParams) { $entry.ExtraParams } else { "" }
                $rows = Get-FssJobListRows -JobPath $entry.JobPath -MenuNo $entry.MenuNo -PageIndex $script:CurrentPage -SearchWrd $txtSearch.Text -Sdate $txtSdate.Text -Edate $txtEdate.Text -ExtraParams $extra
                foreach ($r in $rows) {
                    $grid.Rows.Add($r.Title, $r.Slno) | Out-Null
                    $script:JobRowFiles[$r.Slno] = $r.Files
                }
                $lblPage.Text = "$($script:CurrentPage)페이지 ($($rows.Count)건)"
                Write-Log "조회 완료: $($rows.Count)건"
            } catch {
                Write-Log "조회 실패: $($_.Exception.Message)"
            }
            return
        }

        Write-Log "조회 중: $($entry.Name) ($($script:CurrentPage)페이지)"
        try {
            $extra = if ($entry.ExtraParams) { $entry.ExtraParams } else { "" }
            $items = Get-FssBoardList -BbsId $entry.BbsId -MenuNo $entry.MenuNo -PageIndex $script:CurrentPage -SearchWrd $txtSearch.Text -ExtraParams $extra
            foreach ($it in $items) {
                $grid.Rows.Add($it.Title, $it.NttId) | Out-Null
            }
            $lblPage.Text = "$($script:CurrentPage)페이지 ($($items.Count)건)"
            Write-Log "조회 완료: $($items.Count)건"
        } catch {
            Write-Log "조회 실패: $($_.Exception.Message)"
        }
    }

    $tree.Add_AfterSelect({
        param($s, $e)
        if ($e.Node.Tag) {
            $script:CurrentEntry = $e.Node.Tag
            $script:CurrentPage = 1
            $lblCurrent.Text = "$($script:CurrentEntry.Category) > $($script:CurrentEntry.Group) > $($script:CurrentEntry.Name)"
            if ($script:CurrentEntry.Type -eq "job" -and $script:CurrentEntry.SubType -eq "table") {
                $btnDownloadSeries.Text = "엑셀 다운로드"
                $btnDownloadSelected.Enabled = $false
            } else {
                $btnDownloadSeries.Text = "이 게시판 전체(시리즈) 다운로드"
                $btnDownloadSelected.Enabled = $true
            }
            Load-CurrentBoard
        }
    })

    $btnSearch.Add_Click({ $script:CurrentPage = 1; Load-CurrentBoard })
    $btnNext.Add_Click({ $script:CurrentPage++; Load-CurrentBoard })
    $btnPrev.Add_Click({ if ($script:CurrentPage -gt 1) { $script:CurrentPage--; Load-CurrentBoard } })

    function Process-DownloadedFile {
        param([string]$SavedPath)
        if ($chkPdf.Checked -and $SavedPath -match '\.hwp$') {
            $pdfPath = [System.IO.Path]::ChangeExtension($SavedPath, ".pdf")
            $result = Convert-HwpToPdf -HwpPath $SavedPath -PdfPath $pdfPath
            Write-Log "  $([System.IO.Path]::GetFileName($SavedPath)) -> $($result.Message)"
        }
    }

    $btnDownloadSelected.Add_Click({
        $entry = $script:CurrentEntry
        $isJobList = ($entry -and $entry.Type -eq "job" -and $entry.SubType -eq "list")
        if (-not $entry -or ($entry.Type -ne "bbs" -and -not $isJobList)) {
            Write-Log "게시판형 또는 목록형 검색결과 항목을 선택한 뒤 다운로드할 행을 선택하세요."
            return
        }
        if ($grid.SelectedRows.Count -eq 0) {
            Write-Log "다운로드할 항목을 목록에서 선택하세요."
            return
        }

        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = "다운로드 받을 폴더를 선택하세요"
        if (-not (Test-Path $defaultDownloadRoot)) { New-Item -ItemType Directory -Path $defaultDownloadRoot -Force | Out-Null }
        $fbd.SelectedPath = $defaultDownloadRoot
        if ($fbd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $outDir = $fbd.SelectedPath

        foreach ($row in $grid.SelectedRows) {
            $id = $row.Cells["NttId"].Value
            $title = $row.Cells["Title"].Value
            Write-Log "처리 중: $title"
            try {
                $files = if ($isJobList) { $script:JobRowFiles[$id] } else { (Get-FssBoardDetail -BbsId $entry.BbsId -MenuNo $entry.MenuNo -NttId $id).Files }
                if (-not $files -or $files.Count -eq 0) {
                    Write-Log "  첨부파일 없음"
                    continue
                }
                foreach ($f in $files) {
                    $saved = Save-FssAttachment -DownloadUrl $f.DownloadUrl -OutDir $outDir -PreferredFileName $f.FileName
                    Write-Log "  저장됨: $saved"
                    Process-DownloadedFile -SavedPath $saved
                }
            } catch {
                Write-Log "  오류: $($_.Exception.Message)"
            }
        }
        Write-Log "다운로드 완료: $outDir"
    })

    $btnDownloadSeries.Add_Click({
        $entry = $script:CurrentEntry
        if (-not $entry) {
            Write-Log "항목을 먼저 선택하세요."
            return
        }

        if ($entry.Type -eq "job" -and $entry.SubType -eq "table") {
            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Filter = "Excel 파일 (*.xls)|*.xls"
            $sfd.FileName = ($entry.Name -replace '[\\/:*?"<>|]', '_') + "_" + (Get-Date -Format 'yyyyMMdd') + ".xls"
            if (-not (Test-Path $defaultDownloadRoot)) { New-Item -ItemType Directory -Path $defaultDownloadRoot -Force | Out-Null }
            $sfd.InitialDirectory = $defaultDownloadRoot
            if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

            Write-Log "'$($entry.Name)' 엑셀 다운로드 시작..."
            try {
                $extra = if ($entry.ExtraParams) { $entry.ExtraParams } else { "" }
                $saved = Save-FssJobExcel -JobPath $entry.JobPath -MenuNo $entry.MenuNo -OutPath $sfd.FileName -SearchStr $txtSearch.Text -Sdate $txtSdate.Text -Edate $txtEdate.Text -ExtraParams $extra
                Write-Log "엑셀 다운로드 완료: $saved"
            } catch {
                Write-Log "오류: $($_.Exception.Message)"
            }
            return
        }

        $isJobList = ($entry.Type -eq "job" -and $entry.SubType -eq "list")
        if ($entry.Type -ne "bbs" -and -not $isJobList) {
            Write-Log "게시판형 또는 목록형 검색결과 항목을 먼저 선택하세요."
            return
        }

        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = "'$($entry.Name)' 전체를 저장할 폴더를 선택하세요"
        if (-not (Test-Path $defaultDownloadRoot)) { New-Item -ItemType Directory -Path $defaultDownloadRoot -Force | Out-Null }
        $fbd.SelectedPath = $defaultDownloadRoot
        if ($fbd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $seriesDir = Join-Path $fbd.SelectedPath ($entry.Name -replace '[\\/:*?"<>|]', '_')

        Write-Log "'$($entry.Name)' 전체 다운로드 시작 -> $seriesDir"
        try {
            $extra = if ($entry.ExtraParams) { $entry.ExtraParams } else { "" }
            if ($isJobList) {
                $manifestPath = Save-FssJobSeries -JobPath $entry.JobPath -MenuNo $entry.MenuNo -OutDir $seriesDir -SearchWrd $txtSearch.Text -Sdate $txtSdate.Text -Edate $txtEdate.Text -ExtraParams $extra -OnProgress { param($m) Write-Log $m }
            } else {
                $manifestPath = Save-FssSeries -BbsId $entry.BbsId -MenuNo $entry.MenuNo -OutDir $seriesDir -ExtraParams $extra -OnProgress { param($m) Write-Log $m }
            }

            if ($chkPdf.Checked) {
                Get-ChildItem -Path $seriesDir -Filter "*.hwp" | ForEach-Object {
                    Process-DownloadedFile -SavedPath $_.FullName
                }
            }
            Write-Log "전체 다운로드 완료. 목록: $manifestPath"
        } catch {
            Write-Log "오류: $($_.Exception.Message)"
        }
    })

    $btnNotifySettings.Add_Click({
        Show-FssNotifySettings
    })

    $form.Add_Shown({
        if ($script:FssDefaultItem) {
            foreach ($catNode in $tree.Nodes) {
                foreach ($grpNode in $catNode.Nodes) {
                    foreach ($itemNode in $grpNode.Nodes) {
                        if ($itemNode.Tag -eq $script:FssDefaultItem) {
                            $tree.SelectedNode = $itemNode
                            $itemNode.EnsureVisible()
                        }
                    }
                }
            }
        }
    })

    [System.Windows.Forms.Application]::Run($form)
}

function Show-FssNotifySettings {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "알림 설정 (새 글 이메일 알림)"
    $dlg.ClientSize = New-Object System.Drawing.Size(560, 620)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $config = Get-FssNotifyConfig

    $chkEnabled = New-Object System.Windows.Forms.CheckBox
    $chkEnabled.Left = 15; $chkEnabled.Top = 12; $chkEnabled.Width = 300
    $chkEnabled.Text = "알림 사용"
    $chkEnabled.Checked = [bool]$config.Enabled

    $lblSender = New-Object System.Windows.Forms.Label
    $lblSender.Left = 15; $lblSender.Top = 44; $lblSender.Width = 150
    $lblSender.Text = "발신 Gmail 주소:"
    $txtSender = New-Object System.Windows.Forms.TextBox
    $txtSender.Left = 170; $txtSender.Top = 41; $txtSender.Width = 370
    $txtSender.Text = $config.SenderEmail

    $lblAppPw = New-Object System.Windows.Forms.Label
    $lblAppPw.Left = 15; $lblAppPw.Top = 72; $lblAppPw.Width = 150
    $lblAppPw.Text = "앱 비밀번호:"
    $txtAppPw = New-Object System.Windows.Forms.TextBox
    $txtAppPw.Left = 170; $txtAppPw.Top = 69; $txtAppPw.Width = 370
    $txtAppPw.PasswordChar = '*'
    $txtAppPw.Text = $config.AppPassword

    $lblRecipients = New-Object System.Windows.Forms.Label
    $lblRecipients.Left = 15; $lblRecipients.Top = 100; $lblRecipients.Width = 150
    $lblRecipients.Text = "받는사람(콤마 구분):"
    $txtRecipients = New-Object System.Windows.Forms.TextBox
    $txtRecipients.Left = 170; $txtRecipients.Top = 97; $txtRecipients.Width = 370
    $txtRecipients.Text = ($config.Recipients -join ", ")

    $lblSmtp = New-Object System.Windows.Forms.Label
    $lblSmtp.Left = 15; $lblSmtp.Top = 128; $lblSmtp.Width = 150
    $lblSmtp.Text = "SMTP 서버 / 포트:"
    $txtSmtpHost = New-Object System.Windows.Forms.TextBox
    $txtSmtpHost.Left = 170; $txtSmtpHost.Top = 125; $txtSmtpHost.Width = 260
    $txtSmtpHost.Text = $config.SmtpHost
    $txtSmtpPort = New-Object System.Windows.Forms.TextBox
    $txtSmtpPort.Left = 440; $txtSmtpPort.Top = 125; $txtSmtpPort.Width = 60
    $txtSmtpPort.Text = "$($config.SmtpPort)"

    $chkSsl = New-Object System.Windows.Forms.CheckBox
    $chkSsl.Left = 170; $chkSsl.Top = 153; $chkSsl.Width = 300
    $chkSsl.Text = "SSL/TLS 사용 (Gmail은 반드시 켜야 함)"
    $chkSsl.Checked = [bool]$config.UseSsl

    $lblInterval = New-Object System.Windows.Forms.Label
    $lblInterval.Left = 15; $lblInterval.Top = 182; $lblInterval.Width = 150
    $lblInterval.Text = "자동 확인 주기:"
    $cboInterval = New-Object System.Windows.Forms.ComboBox
    $cboInterval.Left = 170; $cboInterval.Top = 179; $cboInterval.Width = 200
    $cboInterval.DropDownStyle = "DropDownList"
    $cboInterval.Items.AddRange(@("1시간마다", "4시간마다", "하루 1회"))
    $intervalIndex = switch ([int]$config.IntervalHours) { 1 { 0 } 4 { 1 } default { 2 } }
    $cboInterval.SelectedIndex = $intervalIndex

    $lblMenus = New-Object System.Windows.Forms.Label
    $lblMenus.Left = 15; $lblMenus.Top = 214; $lblMenus.Width = 400
    $lblMenus.Text = "모니터링할 메뉴 (표 형태인 '회계법인 정보 통합조회'는 제외):"

    $clbMenus = New-Object System.Windows.Forms.CheckedListBox
    $clbMenus.Left = 15; $clbMenus.Top = 236; $clbMenus.Width = 525; $clbMenus.Height = 220
    $clbMenus.CheckOnClick = $true

    $notifiable = Get-FssNotifiableMenus
    $menuKeyByIndex = @{}
    $i = 0
    foreach ($entry in $notifiable) {
        $key = Get-FssMenuKey $entry
        $label = "$($entry.Category) > $($entry.Group) > $($entry.Name)"
        $clbMenus.Items.Add($label) | Out-Null
        $menuKeyByIndex[$i] = $key
        if ($config.MonitoredKeys -contains $key) {
            $clbMenus.SetItemChecked($i, $true)
        }
        $i++
    }

    $lblTaskStatus = New-Object System.Windows.Forms.Label
    $lblTaskStatus.Left = 15; $lblTaskStatus.Top = 466; $lblTaskStatus.Width = 525
    $lblTaskStatus.Text = if (Test-FssNotifyTaskRegistered) { "작업 스케줄러 등록 상태: 등록됨" } else { "작업 스케줄러 등록 상태: 등록 안 됨" }

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Left = 15; $btnSave.Top = 494; $btnSave.Width = 120; $btnSave.Height = 30
    $btnSave.Text = "설정 저장"

    $btnTestMail = New-Object System.Windows.Forms.Button
    $btnTestMail.Left = 145; $btnTestMail.Top = 494; $btnTestMail.Width = 120; $btnTestMail.Height = 30
    $btnTestMail.Text = "지금 확인/테스트"

    $btnRegisterTask = New-Object System.Windows.Forms.Button
    $btnRegisterTask.Left = 275; $btnRegisterTask.Top = 494; $btnRegisterTask.Width = 130; $btnRegisterTask.Height = 30
    $btnRegisterTask.Text = "자동 확인 등록"

    $btnUnregisterTask = New-Object System.Windows.Forms.Button
    $btnUnregisterTask.Left = 410; $btnUnregisterTask.Top = 494; $btnUnregisterTask.Width = 130; $btnUnregisterTask.Height = 30
    $btnUnregisterTask.Text = "자동 확인 해제"

    $txtNotifyLog = New-Object System.Windows.Forms.TextBox
    $txtNotifyLog.Left = 15; $txtNotifyLog.Top = 532; $txtNotifyLog.Width = 525; $txtNotifyLog.Height = 70
    $txtNotifyLog.Multiline = $true
    $txtNotifyLog.ScrollBars = "Vertical"
    $txtNotifyLog.ReadOnly = $true

    $dlg.Controls.AddRange(@(
        $chkEnabled, $lblSender, $txtSender, $lblAppPw, $txtAppPw, $lblRecipients, $txtRecipients,
        $lblSmtp, $txtSmtpHost, $txtSmtpPort, $chkSsl, $lblInterval, $cboInterval,
        $lblMenus, $clbMenus, $lblTaskStatus, $btnSave, $btnTestMail, $btnRegisterTask, $btnUnregisterTask, $txtNotifyLog
    ))

    function Write-NotifyDlgLog {
        param([string]$Msg)
        $txtNotifyLog.AppendText("$Msg`r`n")
        [System.Windows.Forms.Application]::DoEvents()
    }

    function Get-ConfigFromForm {
        $recipients = $txtRecipients.Text -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        $checkedKeys = New-Object System.Collections.Generic.List[string]
        for ($idx = 0; $idx -lt $clbMenus.Items.Count; $idx++) {
            if ($clbMenus.GetItemChecked($idx)) { $checkedKeys.Add($menuKeyByIndex[$idx]) }
        }
        $intervalHours = switch ($cboInterval.SelectedIndex) { 0 { 1 } 1 { 4 } default { 24 } }

        return [PSCustomObject]@{
            Enabled       = $chkEnabled.Checked
            SenderEmail   = $txtSender.Text.Trim()
            AppPassword   = $txtAppPw.Text
            SmtpHost      = $txtSmtpHost.Text.Trim()
            SmtpPort      = [int]$txtSmtpPort.Text
            UseSsl        = $chkSsl.Checked
            Recipients    = @($recipients)
            MonitoredKeys = @($checkedKeys)
            IntervalHours = $intervalHours
        }
    }

    $btnSave.Add_Click({
        try {
            $newConfig = Get-ConfigFromForm
            Save-FssNotifyConfig -Config $newConfig
            Write-NotifyDlgLog "설정을 저장했습니다: $(Get-FssNotifyConfigPath)"
        } catch {
            Write-NotifyDlgLog "저장 실패: $($_.Exception.Message)"
        }
    })

    $btnTestMail.Add_Click({
        try {
            $newConfig = Get-ConfigFromForm
            Save-FssNotifyConfig -Config $newConfig
            Write-NotifyDlgLog "설정을 저장하고 지금 확인을 실행합니다..."
            Invoke-FssNotificationCheck -SendTestMail -OnProgress { param($m) Write-NotifyDlgLog $m }
        } catch {
            Write-NotifyDlgLog "오류: $($_.Exception.Message)"
        }
    })

    $btnRegisterTask.Add_Click({
        try {
            $newConfig = Get-ConfigFromForm
            Save-FssNotifyConfig -Config $newConfig
            $scriptPath = Get-FssMergedScriptPath
            if (-not (Test-Path $scriptPath)) {
                Write-NotifyDlgLog "실행 스크립트를 찾을 수 없습니다: $scriptPath (build.ps1로 dist를 다시 빌드했는지 확인하세요)"
                return
            }
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Windows 작업 스케줄러에 '$(Get-FssNotifyTaskName)' 작업을 등록합니다.`r`n실행 대상: $scriptPath`r`n주기: $($cboInterval.Text)`r`n`r`n계속할까요?",
                "작업 스케줄러 등록 확인", "YesNo", "Question")
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

            Register-FssNotifyTask -ScriptPath $scriptPath -IntervalHours $newConfig.IntervalHours | Out-Null
            $lblTaskStatus.Text = "작업 스케줄러 등록 상태: 등록됨"
            Write-NotifyDlgLog "작업 스케줄러에 등록했습니다 ($($cboInterval.Text))."
        } catch {
            Write-NotifyDlgLog "등록 실패: $($_.Exception.Message)"
        }
    })

    $btnUnregisterTask.Add_Click({
        try {
            Unregister-FssNotifyTask
            $lblTaskStatus.Text = "작업 스케줄러 등록 상태: 등록 안 됨"
            Write-NotifyDlgLog "작업 스케줄러 등록을 해제했습니다."
        } catch {
            Write-NotifyDlgLog "해제 실패: $($_.Exception.Message)"
        }
    })

    $dlg.ShowDialog() | Out-Null
}
