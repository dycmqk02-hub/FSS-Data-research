Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-FssItemPreview {
    param(
        [string]$Title,
        [string]$RegDate = "",
        [object[]]$Files = @(),
        [string]$ViewUrl = "",
        [scriptblock]$OnDownload = $null
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "미리보기"
    $dlg.ClientSize = New-Object System.Drawing.Size(520, 320)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Left = 15; $lblTitle.Top = 15; $lblTitle.Width = 490; $lblTitle.Height = 50
    $lblTitle.Font = New-Object System.Drawing.Font("맑은 고딕", 10, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Text = $Title

    $lblDate = New-Object System.Windows.Forms.Label
    $lblDate.Left = 15; $lblDate.Top = 70; $lblDate.Width = 490
    $lblDate.Text = if ($RegDate) { "등록일: $RegDate" } else { "" }

    $lblFilesHeader = New-Object System.Windows.Forms.Label
    $lblFilesHeader.Left = 15; $lblFilesHeader.Top = 98; $lblFilesHeader.Width = 200
    $lblFilesHeader.Text = "첨부파일:"

    $lstFiles = New-Object System.Windows.Forms.ListBox
    $lstFiles.Left = 15; $lstFiles.Top = 120; $lstFiles.Width = 490; $lstFiles.Height = 130
    if ($Files.Count -eq 0) {
        $lstFiles.Items.Add("(첨부파일 없음)") | Out-Null
    } else {
        foreach ($f in $Files) { $lstFiles.Items.Add($f.FileName) | Out-Null }
    }

    $btnOpenUrl = New-Object System.Windows.Forms.Button
    $btnOpenUrl.Left = 15; $btnOpenUrl.Top = 260; $btnOpenUrl.Width = 160; $btnOpenUrl.Height = 30
    $btnOpenUrl.Text = "원본 페이지 열기"
    $btnOpenUrl.Enabled = [bool]$ViewUrl

    $btnDownload = New-Object System.Windows.Forms.Button
    $btnDownload.Left = 185; $btnDownload.Top = 260; $btnDownload.Width = 120; $btnDownload.Height = 30
    $btnDownload.Text = "다운로드"
    $btnDownload.Enabled = ([bool]$OnDownload -and $Files.Count -gt 0)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Left = 315; $btnClose.Top = 260; $btnClose.Width = 100; $btnClose.Height = 30
    $btnClose.Text = "닫기"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $dlg.Controls.AddRange(@($lblTitle, $lblDate, $lblFilesHeader, $lstFiles, $btnOpenUrl, $btnDownload, $btnClose))
    $dlg.AcceptButton = $btnClose

    $btnOpenUrl.Add_Click({ if ($ViewUrl) { Start-Process $ViewUrl } })
    $btnDownload.Add_Click({ if ($OnDownload) { & $OnDownload } })

    $dlg.ShowDialog() | Out-Null
}

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
    $tree.PathSeparator = " > "

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
    $lblCurrent.Left = 10; $lblCurrent.Top = 8; $lblCurrent.Width = 550; $lblCurrent.Height = 26
    $lblCurrent.Font = New-Object System.Drawing.Font("맑은 고딕", 11, [System.Drawing.FontStyle]::Bold)
    $lblCurrent.Text = "왼쪽 트리에서 항목을 선택하세요"

    $btnProgress = New-Object System.Windows.Forms.Button
    $btnProgress.Left = 570; $btnProgress.Top = 6; $btnProgress.Width = 140; $btnProgress.Height = 26
    $btnProgress.Text = "대기 중"
    $btnProgress.Enabled = $false
    $btnProgress.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnProgress.ForeColor = [System.Drawing.Color]::White
    $btnProgress.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Left = 715; $btnCancel.Top = 6; $btnCancel.Width = 70; $btnCancel.Height = 26
    $btnCancel.Text = "중지"
    $btnCancel.Enabled = $false
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(196, 43, 28)
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

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
    $lblPage.Left = 495; $lblPage.Top = 46; $lblPage.Width = 300
    $lblPage.Text = "1페이지"

    $lblRange = New-Object System.Windows.Forms.Label
    $lblRange.Left = 10; $lblRange.Top = 80; $lblRange.Width = 405
    $lblRange.Text = "결산연월(예: 202512): 회계법인 통합조회 메뉴 전용, 미기재시 전체 출력"

    $txtSdate = New-Object System.Windows.Forms.TextBox
    # 문구가 길어져서 "다음" 버튼과의 정렬은 포기함(6자리 숫자 칸이라 넓을 필요는 없음)
    $txtSdate.Left = 419; $txtSdate.Top = 78; $txtSdate.Width = 55
    $txtSdate.Text = ""

    $lblRangeDash = New-Object System.Windows.Forms.Label
    $lblRangeDash.Left = 478; $lblRangeDash.Top = 80; $lblRangeDash.Width = 12
    $lblRangeDash.Text = "~"

    $txtEdate = New-Object System.Windows.Forms.TextBox
    $txtEdate.Left = 494; $txtEdate.Top = 78; $txtEdate.Width = 55
    $txtEdate.Text = ""

    $topPanel.Controls.AddRange(@($lblCurrent, $btnProgress, $btnCancel, $txtSearch, $btnSearch, $btnPrev, $btnNext, $lblPage, $lblRange, $txtSdate, $lblRangeDash, $txtEdate))

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

    $btnAiSettings = New-Object System.Windows.Forms.Button
    $btnAiSettings.Left = 770; $btnAiSettings.Top = 8; $btnAiSettings.Width = 110; $btnAiSettings.Height = 30
    $btnAiSettings.Text = "AI 요약 설정"

    $btnPanel.Controls.AddRange(@($btnDownloadSelected, $btnDownloadSeries, $chkPdf, $btnNotifySettings, $btnAiSettings))

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
    $grid.Columns.Add("Snippet", "본문 검색결과 미리보기") | Out-Null
    $grid.Columns.Add("Summary", "AI 요약 (첨부파일)") | Out-Null
    $btnColSummarize = New-Object System.Windows.Forms.DataGridViewButtonColumn
    $btnColSummarize.Name = "SummarizeBtn"
    $btnColSummarize.HeaderText = ""
    $btnColSummarize.Text = "AI요약"
    $btnColSummarize.UseColumnTextForButtonValue = $true
    $btnColSummarize.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
    $btnColSummarize.Width = 70
    $grid.Columns.Add($btnColSummarize) | Out-Null
    $grid.Columns.Add("NttId", "관리번호") | Out-Null
    $grid.Columns["NttId"].Visible = $false
    $grid.Columns["Title"].FillWeight = 25
    $grid.Columns["Snippet"].FillWeight = 35
    $grid.Columns["Summary"].FillWeight = 40
    $grid.BackgroundColor = [System.Drawing.Color]::White

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
    $script:BoardCache = @{}
    $script:DetailCache = @{}
    $script:SummaryCache = @{}
    $script:LastSearchKeyword = ""
    $script:AggregateEntries = $null
    $script:AggregateRowInfo = @{}
    $script:CancelRequested = $false
    $script:TreeLocked = $false

    function Get-FssKeywordSnippet {
        # 본문에서 검색어가 처음 등장한 위치 기준으로 미리보기 문자열을 만듦(하이라이트 렌더링용 원본 텍스트 유지).
        # 그리드 컬럼 폭이 좁아 뒷부분이 잘려도 검색어 자체는 항상 보이도록, 앞쪽 문맥은 짧게(기본 8자)
        # 잡고 검색어를 미리보기 맨 앞쪽에 오게 함(뒤쪽 문맥은 넉넉히 붙여도 어차피 잘리면 그만이라 상관없음).
        param([string]$Text, [string]$Keyword, [int]$Before = 8, [int]$After = 60)
        if (-not $Text) { return "" }
        $idx = $Text.IndexOf($Keyword, [System.StringComparison]::OrdinalIgnoreCase)
        if ($idx -lt 0) { return "" }
        $start = [Math]::Max(0, $idx - $Before)
        $len = [Math]::Min($Text.Length - $start, $Keyword.Length + $Before + $After)
        $snippet = $Text.Substring($start, $len)
        if ($start -gt 0) { $snippet = "..." + $snippet }
        if ($start + $len -lt $Text.Length) { $snippet = $snippet + "..." }
        return $snippet
    }

    function Start-FssBusy {
        # 조회 중 DoEvents()가 메시지 큐를 처리하면서 트리 클릭 등 다른 조작까지 받아버리면,
        # 조회가 끝나기 전에 $script:CurrentEntry가 다른 메뉴로 덮어써지는 재진입 버그가 생김.
        # 조회하는 동안은 트리/검색/페이지 이동을 잠가서 이걸 막음.
        # 트리는 Enabled=false로 잠그면 선택된 노드 강조색이 회색으로 바뀌어 어떤 메뉴를 보고 있는지
        # 안 보이므로, Enabled는 그대로 두고 $script:TreeLocked + BeforeSelect 취소로만 선택 변경을 막음
        # (선택색은 파란색 그대로 유지됨).
        $script:CancelRequested = $false
        $btnProgress.Text = "조회중 (0%)"
        $btnProgress.Enabled = $true
        $btnCancel.Enabled = $true
        $script:TreeLocked = $true
        $txtSearch.Enabled = $false
        $btnSearch.Enabled = $false
        $btnPrev.Enabled = $false
        $btnNext.Enabled = $false
        [System.Windows.Forms.Application]::DoEvents()
    }

    function Stop-FssBusy {
        $btnProgress.Text = "대기 중"
        $btnProgress.Enabled = $false
        $btnCancel.Enabled = $false
        $script:TreeLocked = $false
        $txtSearch.Enabled = $true
        $btnSearch.Enabled = $true
        $btnPrev.Enabled = $true
        $btnNext.Enabled = $true
        [System.Windows.Forms.Application]::DoEvents()
    }

    function Update-FssProgress {
        param($Page, $TotalPages, $Pct)
        if ($null -ne $Pct) {
            $btnProgress.Text = "조회중 ($Pct%)"
        } else {
            $btnProgress.Text = "조회중 ($Page 페이지)"
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    $btnCancel.Add_Click({
        $script:CancelRequested = $true
        Write-Log "중지 요청됨 - 지금까지 모은 결과만 표시합니다."
    })

    function Load-CurrentBoard {
        if (-not $script:CurrentEntry) { return }
        $entry = $script:CurrentEntry
        $grid.Rows.Clear()
        $script:JobRowFiles = @{}

        if ($entry.Type -eq "info") {
            $infoUrl = "https://www.fss.or.kr/fss/main/contents.do?menuNo=$($entry.MenuNo)"
            $lblPage.Text = ""
            Write-Log "정보성 페이지입니다 (목록/첨부파일 없음): $infoUrl"
            Show-FssItemPreview -Title $entry.Name -ViewUrl $infoUrl
            return
        }

        if ($entry.Type -eq "job" -and $entry.SubType -eq "table") {
            Write-Log "'$($entry.Name)'은 결과가 가로로 긴 데이터 표(연번~결산년월)로 제공되어 화면 목록 대신 엑셀로만 내려받습니다. 회계법인명(검색창)/결산연월 범위(선택)를 입력한 뒤 [엑셀 다운로드]를 누르세요."
            $lblPage.Text = ""
            return
        }

        $keyword = $txtSearch.Text.Trim()
        $pageSize = 20

        if ($entry.Type -eq "job" -and $entry.SubType -eq "list") {
            try {
                $extra = if ($entry.ExtraParams) { $entry.ExtraParams } else { "" }
                $cacheKey = "job_$($entry.JobPath)_$($txtSdate.Text)_$($txtEdate.Text)"
                if ($script:BoardCache.ContainsKey($cacheKey)) {
                    $allRows = $script:BoardCache[$cacheKey]
                } else {
                    Write-Log "조회 중: $($entry.Name)"
                    Start-FssBusy
                    # 게시판별 서버 검색조건이 제각각이라 신뢰 불가 -> 전 페이지를 가져와 직접 페이지네이션/필터링
                    $allRows = Get-FssJobAllListRows -JobPath $entry.JobPath -MenuNo $entry.MenuNo -Sdate $txtSdate.Text -Edate $txtEdate.Text -ExtraParams $extra -OnProgress { param($p, $t, $pct) Update-FssProgress $p $t $pct }
                    Stop-FssBusy
                    $script:BoardCache[$cacheKey] = $allRows
                }
                $script:LastSearchKeyword = $keyword
                $filtered = if ($keyword) { @($allRows | Where-Object { $_.Title.IndexOf($keyword, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 }) } else { @($allRows) }
                $totalPages = [Math]::Max(1, [Math]::Ceiling($filtered.Count / $pageSize))
                if ($script:CurrentPage -gt $totalPages) { $script:CurrentPage = $totalPages }
                $startIdx = ($script:CurrentPage - 1) * $pageSize
                $rows = $filtered | Select-Object -Skip $startIdx -First $pageSize
                foreach ($r in $rows) {
                    # 검색폼형(job/list) 게시판은 상세 본문 파싱을 아직 지원하지 않아 제목만 검색(Snippet은 항상 빈 값)
                    $grid.Rows.Add($r.Title, "", "", $null, $r.Slno) | Out-Null
                    $script:JobRowFiles[$r.Slno] = $r.Files
                }
                $lblPage.Text = "$($script:CurrentPage)/$totalPages 페이지 (페이지당 ${pageSize}건 / 총 $($filtered.Count)건)"
                if ($keyword) {
                    Write-Log "검색 완료: 전체 $($allRows.Count)건 중 '$keyword' 포함 $($filtered.Count)건"
                } else {
                    Write-Log "조회 완료: 총 $($filtered.Count)건"
                }
            } catch {
                Stop-FssBusy
                Write-Log "조회 실패: $($_.Exception.Message)"
            }
            return
        }

        try {
            $extra = if ($entry.ExtraParams) { $entry.ExtraParams } else { "" }
            $cacheKey = "bbs_$($entry.BbsId)"
            $busyStarted = $false
            if (-not $script:BoardCache.ContainsKey($cacheKey)) {
                Write-Log "조회 중: $($entry.Name)"
                Start-FssBusy
                $busyStarted = $true
                # 게시판별 서버 검색조건이 제각각이라 신뢰 불가(예: 심사·감리지적사례는 제목이 아니라
                # 쟁점분야/관련기준서/결정년도만 검색됨) -> 전 페이지를 가져와 직접 페이지네이션/필터링
                $allItems = Get-FssBoardAllItems -BbsId $entry.BbsId -MenuNo $entry.MenuNo -ExtraParams $extra -OnProgress { param($p, $t, $pct) Update-FssProgress $p $t $pct }
                $script:BoardCache[$cacheKey] = $allItems
            } else {
                $allItems = $script:BoardCache[$cacheKey]
            }

            $script:LastSearchKeyword = $keyword
            $resultRows = New-Object System.Collections.Generic.List[object]

            if (-not $keyword) {
                foreach ($it in $allItems) { $resultRows.Add([PSCustomObject]@{ Title = $it.Title; NttId = $it.NttId; Snippet = "" }) }
            } else {
                # 1단계: 제목에 검색어가 있는 항목(빠름, 추가 요청 불필요)
                $titleMatchIds = @{}
                foreach ($it in $allItems) {
                    if ($it.Title.IndexOf($keyword, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                        $titleMatchIds[$it.NttId] = $true
                        $resultRows.Add([PSCustomObject]@{ Title = $it.Title; NttId = $it.NttId; Snippet = "" })
                    }
                }
                # 2단계: 제목엔 없지만 본문에 있을 수 있는 나머지 항목 - 상세페이지를 조회해야 하므로
                # 이미 조회한 적 있는 건(DetailCache)은 재사용하고, 없는 건만 병렬로 새로 가져옴
                $bodyCandidates = @($allItems | Where-Object { -not $titleMatchIds.ContainsKey($_.NttId) })
                $toFetch = @($bodyCandidates | Where-Object { -not $script:DetailCache.ContainsKey("$($entry.BbsId)_$($_.NttId)") } | ForEach-Object { [PSCustomObject]@{ BbsId = $entry.BbsId; MenuNo = $entry.MenuNo; NttId = $_.NttId } })
                if ($toFetch.Count -gt 0) {
                    Write-Log "본문 검색 중: $($toFetch.Count)건 상세 조회..."
                    if (-not $busyStarted) { Start-FssBusy; $busyStarted = $true }
                    $fetched = Get-FssBoardBodiesParallel -Items $toFetch -OnProgress { param($c, $t, $pct) Update-FssProgress $c $t $pct }
                    foreach ($kv in $fetched.GetEnumerator()) { $script:DetailCache[$kv.Key] = $kv.Value }
                }
                foreach ($it in $bodyCandidates) {
                    $d = $script:DetailCache["$($entry.BbsId)_$($it.NttId)"]
                    if ($d -and $d.BodyText -and $d.BodyText.IndexOf($keyword, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                        $resultRows.Add([PSCustomObject]@{ Title = $it.Title; NttId = $it.NttId; Snippet = (Get-FssKeywordSnippet -Text $d.BodyText -Keyword $keyword) })
                    }
                }
            }
            if ($busyStarted) { Stop-FssBusy }

            $totalPages = [Math]::Max(1, [Math]::Ceiling($resultRows.Count / $pageSize))
            if ($script:CurrentPage -gt $totalPages) { $script:CurrentPage = $totalPages }
            $startIdx = ($script:CurrentPage - 1) * $pageSize
            $pageRows = $resultRows | Select-Object -Skip $startIdx -First $pageSize
            foreach ($row in $pageRows) {
                $grid.Rows.Add($row.Title, $row.Snippet, "", $null, $row.NttId) | Out-Null
            }
            $lblPage.Text = "$($script:CurrentPage)/$totalPages 페이지 (페이지당 ${pageSize}건 / 총 $($resultRows.Count)건)"
            if ($keyword) {
                Write-Log "검색 완료: 전체 $($allItems.Count)건 중 제목/본문에 '$keyword' 포함 $($resultRows.Count)건"
            } else {
                Write-Log "조회 완료: 총 $($resultRows.Count)건"
            }
        } catch {
            Stop-FssBusy
            Write-Log "조회 실패: $($_.Exception.Message)"
        }
    }

    function Get-AllLeafTags {
        # 공시/회계 같은 상위(카테고리)나 회계기준 같은 중간(그룹) 노드 아래의 모든 리프(실제 조회 대상)를 재귀적으로 수집
        param($node, $list)
        if ($node.Tag) { $list.Add($node.Tag); return }
        foreach ($child in $node.Nodes) { Get-AllLeafTags $child $list }
    }

    function Load-AggregateSearch {
        $grid.Rows.Clear()
        $script:JobRowFiles = @{}
        $script:AggregateRowInfo = @{}
        $keyword = $txtSearch.Text.Trim()

        $script:LastSearchKeyword = $keyword

        if (-not $keyword) {
            $lblPage.Text = ""
            Write-Log "상위 메뉴가 선택된 상태입니다. 검색어를 입력하고 조회하면 하위 전체 게시판에서 통합 검색합니다."
            return
        }

        $pageSize = 20
        $allMatches = New-Object System.Collections.Generic.List[object]
        $totalBoards = $script:AggregateEntries.Count
        $boardIndex = 0
        Start-FssBusy

        foreach ($entry in $script:AggregateEntries) {
            if ($script:CancelRequested) { break }
            $boardIndex++
            try {
                $extra = if ($entry.ExtraParams) { $entry.ExtraParams } else { "" }
                if ($entry.Type -eq "bbs") {
                    $cacheKey = "bbs_$($entry.BbsId)"
                    if ($script:BoardCache.ContainsKey($cacheKey)) {
                        $items = $script:BoardCache[$cacheKey]
                    } else {
                        Write-Log "조회 중 ($boardIndex/$totalBoards): $($entry.Group) > $($entry.Name)"
                        $items = Get-FssBoardAllItems -BbsId $entry.BbsId -MenuNo $entry.MenuNo -ExtraParams $extra
                        $script:BoardCache[$cacheKey] = $items
                    }
                    $titleMatchIds = @{}
                    foreach ($it in $items) {
                        if ($it.Title.IndexOf($keyword, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                            $titleMatchIds[$it.NttId] = $true
                            $key = "bbs_$($entry.BbsId)_$($it.NttId)"
                            $script:AggregateRowInfo[$key] = [PSCustomObject]@{ Entry = $entry; NttId = $it.NttId; IsJobList = $false }
                            $allMatches.Add([PSCustomObject]@{ DisplayTitle = "$($entry.Group) > $($entry.Name) > $($it.Title)"; Key = $key; Snippet = "" })
                        }
                    }
                    # 제목엔 없지만 본문에 있을 수 있는 나머지 - 상세페이지를 병렬로 조회(이미 캐시된 건 재사용)
                    $bodyCandidates = @($items | Where-Object { -not $titleMatchIds.ContainsKey($_.NttId) })
                    $toFetch = @($bodyCandidates | Where-Object { -not $script:DetailCache.ContainsKey("$($entry.BbsId)_$($_.NttId)") } | ForEach-Object { [PSCustomObject]@{ BbsId = $entry.BbsId; MenuNo = $entry.MenuNo; NttId = $_.NttId } })
                    if ($toFetch.Count -gt 0) {
                        Write-Log "  본문 검색 중 ($boardIndex/$totalBoards): $($entry.Name) - $($toFetch.Count)건 상세 조회..."
                        $fetched = Get-FssBoardBodiesParallel -Items $toFetch
                        foreach ($kv in $fetched.GetEnumerator()) { $script:DetailCache[$kv.Key] = $kv.Value }
                    }
                    foreach ($it in $bodyCandidates) {
                        $d = $script:DetailCache["$($entry.BbsId)_$($it.NttId)"]
                        if ($d -and $d.BodyText -and $d.BodyText.IndexOf($keyword, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                            $key = "bbs_$($entry.BbsId)_$($it.NttId)"
                            $script:AggregateRowInfo[$key] = [PSCustomObject]@{ Entry = $entry; NttId = $it.NttId; IsJobList = $false }
                            $allMatches.Add([PSCustomObject]@{ DisplayTitle = "$($entry.Group) > $($entry.Name) > $($it.Title)"; Key = $key; Snippet = (Get-FssKeywordSnippet -Text $d.BodyText -Keyword $keyword) })
                        }
                    }
                } elseif ($entry.Type -eq "job" -and $entry.SubType -eq "list") {
                    $cacheKey = "job_$($entry.JobPath)"
                    if ($script:BoardCache.ContainsKey($cacheKey)) {
                        $rows = $script:BoardCache[$cacheKey]
                    } else {
                        Write-Log "조회 중 ($boardIndex/$totalBoards): $($entry.Group) > $($entry.Name)"
                        $rows = Get-FssJobAllListRows -JobPath $entry.JobPath -MenuNo $entry.MenuNo -Sdate "" -Edate "" -ExtraParams $extra
                        $script:BoardCache[$cacheKey] = $rows
                    }
                    foreach ($r in $rows) {
                        if ($r.Title.IndexOf($keyword, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                            $key = "job_$($entry.JobPath)_$($r.Slno)"
                            $script:AggregateRowInfo[$key] = [PSCustomObject]@{ Entry = $entry; Slno = $r.Slno; Files = $r.Files; IsJobList = $true }
                            $allMatches.Add([PSCustomObject]@{ DisplayTitle = "$($entry.Group) > $($entry.Name) > $($r.Title)"; Key = $key; Snippet = "" })
                        }
                    }
                }
            } catch {
                Write-Log "  오류($($entry.Name)): $($_.Exception.Message)"
            }
            Update-FssProgress $boardIndex $totalBoards ([Math]::Min(100, [int](($boardIndex / $totalBoards) * 100)))
        }
        Stop-FssBusy

        $totalPages = [Math]::Max(1, [Math]::Ceiling($allMatches.Count / $pageSize))
        if ($script:CurrentPage -gt $totalPages) { $script:CurrentPage = $totalPages }
        $startIdx = ($script:CurrentPage - 1) * $pageSize
        $pageItems = $allMatches | Select-Object -Skip $startIdx -First $pageSize
        foreach ($m in $pageItems) {
            $grid.Rows.Add($m.DisplayTitle, $m.Snippet, "", $null, $m.Key) | Out-Null
        }
        $lblPage.Text = "$($script:CurrentPage)/$totalPages 페이지 (페이지당 ${pageSize}건 / 총 $($allMatches.Count)건)"
        if ($script:CancelRequested) {
            Write-Log "중지됨: $boardIndex/$totalBoards 개 게시판까지 조회한 결과 중 '$keyword' 포함 $($allMatches.Count)건"
        } else {
            Write-Log "통합검색 완료: 하위 $totalBoards 개 게시판에서 '$keyword' 포함 $($allMatches.Count)건"
        }
    }

    function Do-Search {
        if ($script:AggregateEntries) { Load-AggregateSearch } else { Load-CurrentBoard }
    }

    $tree.Add_BeforeSelect({
        param($s, $e)
        if ($script:TreeLocked) { $e.Cancel = $true }
    })

    $tree.Add_AfterSelect({
        param($s, $e)
        $node = $e.Node
        if ($node.Tag) {
            $script:CurrentEntry = $node.Tag
            $script:AggregateEntries = $null
            $script:CurrentPage = 1
            $lblCurrent.Text = $node.FullPath
            if ($node.Tag.Type -eq "job" -and $node.Tag.SubType -eq "table") {
                $btnDownloadSeries.Text = "엑셀 다운로드"
                $btnDownloadSeries.Enabled = $true
                $btnDownloadSelected.Enabled = $false
            } else {
                $btnDownloadSeries.Text = "이 게시판 전체(시리즈) 다운로드"
                $btnDownloadSeries.Enabled = $true
                $btnDownloadSelected.Enabled = $true
            }
            Load-CurrentBoard
        } else {
            $leaves = New-Object System.Collections.Generic.List[object]
            Get-AllLeafTags $node $leaves
            $script:AggregateEntries = @($leaves | Where-Object { $_.Type -ne "info" -and -not ($_.Type -eq "job" -and $_.SubType -eq "table") })
            $script:CurrentEntry = $null
            $script:CurrentPage = 1
            $lblCurrent.Text = $node.FullPath
            $btnDownloadSeries.Enabled = $false
            $btnDownloadSelected.Enabled = $false
            $grid.Rows.Clear()
            $lblPage.Text = ""
            Write-Log "'$($node.Text)' 선택됨 - 검색어를 입력하고 조회하면 하위 $($script:AggregateEntries.Count)개 게시판을 통합 검색합니다."
        }
    })

    $btnSearch.Add_Click({ $script:CurrentPage = 1; Do-Search })
    $btnNext.Add_Click({ $script:CurrentPage++; Do-Search })
    $btnPrev.Add_Click({ if ($script:CurrentPage -gt 1) { $script:CurrentPage--; Do-Search } })

    $txtSearch.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $e.SuppressKeyPress = $true
            $script:CurrentPage = 1
            Do-Search
        }
    })

    function Process-DownloadedFile {
        param([string]$SavedPath)
        if ($chkPdf.Checked -and $SavedPath -match '\.hwp$') {
            $pdfPath = [System.IO.Path]::ChangeExtension($SavedPath, ".pdf")
            $result = Convert-HwpToPdf -HwpPath $SavedPath -PdfPath $pdfPath
            Write-Log "  $([System.IO.Path]::GetFileName($SavedPath)) -> $($result.Message)"
        }
    }

    function Invoke-FssFilesDownload {
        param([object[]]$Files)
        if (-not $Files -or $Files.Count -eq 0) {
            Write-Log "첨부파일이 없습니다."
            return
        }
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = "다운로드 받을 폴더를 선택하세요"
        if (-not (Test-Path $defaultDownloadRoot)) { New-Item -ItemType Directory -Path $defaultDownloadRoot -Force | Out-Null }
        $fbd.SelectedPath = $defaultDownloadRoot
        if ($fbd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $outDir = $fbd.SelectedPath
        foreach ($f in $Files) {
            try {
                $saved = Save-FssAttachment -DownloadUrl $f.DownloadUrl -OutDir $outDir -PreferredFileName $f.FileName
                Write-Log "저장됨: $saved"
                Process-DownloadedFile -SavedPath $saved
            } catch {
                Write-Log "오류: $($_.Exception.Message)"
            }
        }
        Write-Log "다운로드 완료: $outDir"
    }

    $grid.Add_CellPainting({
        # 제목/본문미리보기 셀 안에서 검색어와 일치하는 부분만 노란색 음영 + 검정 글씨로 그려서
        # 어디에 검색어가 있는지 한눈에 보이게 함. 검색어가 없거나 해당 셀에 검색어가 없으면
        # $e.Handled를 건드리지 않고 그냥 반환해서 DataGridView 기본 렌더링이 그대로 동작하게 둠.
        param($s, $e)
        if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
        $colName = $grid.Columns[$e.ColumnIndex].Name
        if ($colName -ne "Title" -and $colName -ne "Snippet") { return }
        $keyword = $script:LastSearchKeyword
        if ([string]::IsNullOrEmpty($keyword)) { return }
        $text = [string]$e.FormattedValue
        if ([string]::IsNullOrEmpty($text)) { return }
        $idx = $text.IndexOf($keyword, [System.StringComparison]::OrdinalIgnoreCase)
        if ($idx -lt 0) { return }

        $parts = [System.Windows.Forms.DataGridViewPaintParts]([int][System.Windows.Forms.DataGridViewPaintParts]::Background -bor `
            [int][System.Windows.Forms.DataGridViewPaintParts]::Border -bor `
            [int][System.Windows.Forms.DataGridViewPaintParts]::Focus -bor `
            [int][System.Windows.Forms.DataGridViewPaintParts]::SelectionBackground)
        $e.Paint($e.ClipBounds, $parts)

        $isSelected = (([int]$e.State -band [int][System.Windows.Forms.DataGridViewElementStates]::Selected) -ne 0)
        $foreColor = if ($isSelected) { $e.CellStyle.SelectionForeColor } else { $e.CellStyle.ForeColor }
        $font = $e.CellStyle.Font
        $g = $e.Graphics
        $bounds = $e.CellBounds
        $y = $bounds.Top + [Math]::Max(0, [int](($bounds.Height - $font.Height) / 2))
        $x = $bounds.Left + 2
        $flags = [System.Windows.Forms.TextFormatFlags]::NoPadding -bor [System.Windows.Forms.TextFormatFlags]::SingleLine

        $before = $text.Substring(0, $idx)
        $match = $text.Substring($idx, $keyword.Length)
        $after = $text.Substring($idx + $keyword.Length)
        $measureSize = New-Object System.Drawing.Size(2000, 50)

        $g.SetClip($bounds)
        if ($before) {
            [System.Windows.Forms.TextRenderer]::DrawText($g, $before, $font, (New-Object System.Drawing.Point($x, $y)), $foreColor, $flags)
            $x += [System.Windows.Forms.TextRenderer]::MeasureText($g, $before, $font, $measureSize, $flags).Width
        }
        $matchWidth = [System.Windows.Forms.TextRenderer]::MeasureText($g, $match, $font, $measureSize, $flags).Width
        $g.FillRectangle([System.Drawing.Brushes]::Yellow, $x, $bounds.Top + 1, $matchWidth, $bounds.Height - 2)
        [System.Windows.Forms.TextRenderer]::DrawText($g, $match, $font, (New-Object System.Drawing.Point($x, $y)), [System.Drawing.Color]::Black, $flags)
        $x += $matchWidth
        if ($after) {
            [System.Windows.Forms.TextRenderer]::DrawText($g, $after, $font, (New-Object System.Drawing.Point($x, $y)), $foreColor, $flags)
        }
        $g.ResetClip()
        $e.Handled = $true
    })

    $grid.Add_CellDoubleClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $row = $grid.Rows[$e.RowIndex]
        $id = $row.Cells["NttId"].Value
        $title = $row.Cells["Title"].Value
        try {
            if ($script:AggregateEntries) {
                $info = $script:AggregateRowInfo[$id]
                if (-not $info) { Write-Log "미리보기 정보를 찾을 수 없습니다."; return }
                if ($info.IsJobList) {
                    $viewUrl = "$script:FssBaseUrl/fss/job/$($info.Entry.JobPath)/view.do?menuNo=$($info.Entry.MenuNo)&acntnWrkSlno=$($info.Slno)"
                    if ($info.Entry.ExtraParams) { $viewUrl += "&$($info.Entry.ExtraParams)" }
                    $files = $info.Files
                    Show-FssItemPreview -Title $title -Files $files -ViewUrl $viewUrl -OnDownload { Invoke-FssFilesDownload -Files $files }
                } else {
                    $detailKey = "$($info.Entry.BbsId)_$($info.NttId)"
                    $detail = $script:DetailCache[$detailKey]
                    if (-not $detail) {
                        $detail = Get-FssBoardDetail -BbsId $info.Entry.BbsId -MenuNo $info.Entry.MenuNo -NttId $info.NttId
                        $script:DetailCache[$detailKey] = $detail
                    }
                    $viewUrl = "$script:FssBaseUrl/fss/bbs/$($info.Entry.BbsId)/view.do?nttId=$($info.NttId)&menuNo=$($info.Entry.MenuNo)"
                    $files = $detail.Files
                    Show-FssItemPreview -Title $title -RegDate $detail.RegDate -Files $files -ViewUrl $viewUrl -OnDownload { Invoke-FssFilesDownload -Files $files }
                }
            } elseif ($script:CurrentEntry) {
                $entry = $script:CurrentEntry
                if ($entry.Type -eq "job" -and $entry.SubType -eq "list") {
                    $files = $script:JobRowFiles[$id]
                    $viewUrl = "$script:FssBaseUrl/fss/job/$($entry.JobPath)/view.do?menuNo=$($entry.MenuNo)&acntnWrkSlno=$id"
                    if ($entry.ExtraParams) { $viewUrl += "&$($entry.ExtraParams)" }
                    Show-FssItemPreview -Title $title -Files $files -ViewUrl $viewUrl -OnDownload { Invoke-FssFilesDownload -Files $files }
                } elseif ($entry.Type -eq "bbs") {
                    $detailKey = "$($entry.BbsId)_$id"
                    $detail = $script:DetailCache[$detailKey]
                    if (-not $detail) {
                        $detail = Get-FssBoardDetail -BbsId $entry.BbsId -MenuNo $entry.MenuNo -NttId $id
                        $script:DetailCache[$detailKey] = $detail
                    }
                    $viewUrl = "$script:FssBaseUrl/fss/bbs/$($entry.BbsId)/view.do?nttId=$id&menuNo=$($entry.MenuNo)"
                    $files = $detail.Files
                    Show-FssItemPreview -Title $title -RegDate $detail.RegDate -Files $files -ViewUrl $viewUrl -OnDownload { Invoke-FssFilesDownload -Files $files }
                }
            }
        } catch {
            Write-Log "미리보기 실패: $($_.Exception.Message)"
        }
    })

    $btnDownloadSelected.Add_Click({
        $entry = $script:CurrentEntry
        $isAggregate = [bool]$script:AggregateEntries
        $isJobList = ($entry -and $entry.Type -eq "job" -and $entry.SubType -eq "list")
        if (-not $isAggregate -and (-not $entry -or ($entry.Type -ne "bbs" -and -not $isJobList))) {
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
                if ($isAggregate) {
                    $info = $script:AggregateRowInfo[$id]
                    if (-not $info) {
                        Write-Log "  정보를 찾을 수 없음(캐시가 갱신되었을 수 있습니다. 다시 조회해주세요)"
                        continue
                    }
                    $files = if ($info.IsJobList) { $info.Files } else { (Get-FssBoardDetail -BbsId $info.Entry.BbsId -MenuNo $info.Entry.MenuNo -NttId $info.NttId).Files }
                } else {
                    $files = if ($isJobList) { $script:JobRowFiles[$id] } else { (Get-FssBoardDetail -BbsId $entry.BbsId -MenuNo $entry.MenuNo -NttId $id).Files }
                }
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

    $btnAiSettings.Add_Click({
        Show-FssAiSettings
    })

    $grid.Add_CellContentClick({
        # 그리드의 "AI요약" 버튼(SummarizeBtn 컬럼) 클릭 시에만 동작. 온디맨드 호출이라 여기서 클릭될 때만
        # Gemini API를 호출함(페이지 로드/검색으로는 절대 자동 호출되지 않음 - 비용 통제).
        param($s, $e)
        if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
        if ($grid.Columns[$e.ColumnIndex].Name -ne "SummarizeBtn") { return }
        $row = $grid.Rows[$e.RowIndex]
        $id = $row.Cells["NttId"].Value

        $files = $null
        $summaryKey = $null
        try {
            if ($script:AggregateEntries) {
                $info = $script:AggregateRowInfo[$id]
                if (-not $info) { Write-Log "요약 대상 정보를 찾을 수 없습니다."; return }
                if ($info.IsJobList) {
                    $files = $info.Files
                    $summaryKey = "job_$($info.Entry.JobPath)_$($info.Slno)"
                } else {
                    $detailKey = "$($info.Entry.BbsId)_$($info.NttId)"
                    $detail = $script:DetailCache[$detailKey]
                    if (-not $detail) {
                        $detail = Get-FssBoardDetail -BbsId $info.Entry.BbsId -MenuNo $info.Entry.MenuNo -NttId $info.NttId
                        $script:DetailCache[$detailKey] = $detail
                    }
                    $files = $detail.Files
                    $summaryKey = $detailKey
                }
            } elseif ($script:CurrentEntry) {
                $entry = $script:CurrentEntry
                if ($entry.Type -eq "job" -and $entry.SubType -eq "list") {
                    $files = $script:JobRowFiles[$id]
                    $summaryKey = "job_$($entry.JobPath)_$id"
                } elseif ($entry.Type -eq "bbs") {
                    $detailKey = "$($entry.BbsId)_$id"
                    $detail = $script:DetailCache[$detailKey]
                    if (-not $detail) {
                        $detail = Get-FssBoardDetail -BbsId $entry.BbsId -MenuNo $entry.MenuNo -NttId $id
                        $script:DetailCache[$detailKey] = $detail
                    }
                    $files = $detail.Files
                    $summaryKey = $detailKey
                }
            }
        } catch {
            $row.Cells["Summary"].Value = "요약 실패(첨부 조회): $($_.Exception.Message)"
            return
        }

        if (-not $summaryKey) { Write-Log "요약할 게시글 정보를 찾을 수 없습니다."; return }

        if ($script:SummaryCache.ContainsKey($summaryKey)) {
            $row.Cells["Summary"].Value = $script:SummaryCache[$summaryKey]
            return
        }
        if (-not $files -or $files.Count -eq 0) {
            $row.Cells["Summary"].Value = "(첨부파일 없음)"
            return
        }
        $hwpFiles = @($files | Where-Object { $_.FileName -match '\.hwp$' })
        if ($hwpFiles.Count -eq 0) {
            $row.Cells["Summary"].Value = "(HWP 첨부파일이 없어 요약 불가 - 현재 HWP만 지원)"
            return
        }
        if (-not (Test-HwpInstalled)) {
            $row.Cells["Summary"].Value = "(한컴오피스 미설치로 첨부파일 텍스트 추출 불가)"
            return
        }
        $aiConfig = Get-FssAiConfig
        if (-not $aiConfig.ApiKey) {
            $row.Cells["Summary"].Value = "(API 키 미설정 - 'AI 요약 설정'에서 입력하세요)"
            return
        }

        # 로컬에서 누적 집계한 오늘 사용량이 설정한 일일 한도의 90%를 넘으면 호출 자체를 막고 경고창을 띄움
        # (구글 서버의 실시간 잔여 할당량을 조회하는 게 아니라 이 도구가 지금까지 호출한 만큼만 더한 추정치).
        $usageBefore = Get-FssAiUsage
        $budget = [int]$aiConfig.DailyTokenBudget
        if ($budget -gt 0 -and $usageBefore.TokensUsed -ge [Math]::Floor($budget * 0.9)) {
            $row.Cells["Summary"].Value = "(일일 토큰 한도 90% 도달로 요약 중지됨)"
            [System.Windows.Forms.MessageBox]::Show(
                "토큰이 모자랍니다.`r`n`r`nGemini API 오늘 누적 사용량이 설정한 일일 한도의 90%를 넘어 AI 요약 기능을 멈췄습니다.`r`n(오늘 사용량: $($usageBefore.TokensUsed) / 한도: $budget)`r`n`r`n이 값은 로컬 집계 추정치이며 자정에 초기화됩니다. 'AI 요약 설정'에서 한도를 조정할 수 있습니다.",
                "토큰 한도 임박", "OK", "Warning") | Out-Null
            return
        }

        $row.Cells["Summary"].Value = "요약 중..."
        $grid.Refresh()
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $tempDir = Join-Path $env:TEMP "FSS-DataTool-aitmp"
            if (-not (Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir -Force | Out-Null }
            $combinedText = New-Object System.Text.StringBuilder
            foreach ($f in $hwpFiles) {
                $saved = Save-FssAttachment -DownloadUrl $f.DownloadUrl -OutDir $tempDir -PreferredFileName $f.FileName
                $extracted = Get-HwpPlainText -HwpPath $saved
                if ($extracted) { [void]$combinedText.AppendLine($extracted) }
                Remove-Item -Path $saved -Force -ErrorAction SilentlyContinue
            }
            $plainText = $combinedText.ToString().Trim()
            if (-not $plainText) {
                $row.Cells["Summary"].Value = "(첨부파일에서 텍스트를 추출하지 못했습니다)"
                return
            }
            $result = Invoke-FssGeminiSummary -Text $plainText -ApiKey $aiConfig.ApiKey -Model $aiConfig.Model
            if ($result.TokensUsed -gt 0) {
                $usageAfter = Add-FssAiUsage -Tokens $result.TokensUsed
                if ($budget -gt 0 -and $usageAfter.TokensUsed -ge [Math]::Floor($budget * 0.9) -and $usageAfter.TokensUsed - $result.TokensUsed -lt [Math]::Floor($budget * 0.9)) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "토큰이 모자랍니다.`r`n`r`n방금 호출로 오늘 누적 사용량이 일일 한도의 90%를 넘었습니다(오늘 사용량: $($usageAfter.TokensUsed) / 한도: $budget). 다음 요약 요청부터는 자동으로 중지됩니다.",
                        "토큰 한도 임박", "OK", "Warning") | Out-Null
                }
            }
            $script:SummaryCache[$summaryKey] = $result.Summary
            $row.Cells["Summary"].Value = $result.Summary
        } catch {
            $errMsg = $_.Exception.Message
            if ($errMsg -match '429|RESOURCE_EXHAUSTED|quota') {
                $row.Cells["Summary"].Value = "(구글 무료 한도 초과 - 잠시 후 다시 시도하세요)"
                [System.Windows.Forms.MessageBox]::Show("Gemini API가 무료 등급 요청 한도 초과(429) 응답을 반환했습니다. 잠시 후 다시 시도하거나 'AI 요약 설정'에서 한도/모델을 확인하세요.`r`n`r`n상세: $errMsg", "API 한도 초과", "OK", "Warning") | Out-Null
            } else {
                $row.Cells["Summary"].Value = "요약 실패: $errMsg"
            }
        }
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

function Show-FssUnsavedChangesPrompt {
    $prompt = New-Object System.Windows.Forms.Form
    $prompt.Text = "변경사항 확인"
    $prompt.ClientSize = New-Object System.Drawing.Size(360, 120)
    $prompt.StartPosition = "CenterParent"
    $prompt.FormBorderStyle = "FixedDialog"
    $prompt.MaximizeBox = $false
    $prompt.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Left = 20; $lbl.Top = 20; $lbl.Width = 320; $lbl.Height = 30
    $lbl.Text = "변경사항을 저장하시겠습니까?"

    $btnSaveConfirm = New-Object System.Windows.Forms.Button
    $btnSaveConfirm.Left = 90; $btnSaveConfirm.Top = 65; $btnSaveConfirm.Width = 80; $btnSaveConfirm.Height = 30
    $btnSaveConfirm.Text = "저장"
    $btnSaveConfirm.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancelConfirm = New-Object System.Windows.Forms.Button
    $btnCancelConfirm.Left = 190; $btnCancelConfirm.Top = 65; $btnCancelConfirm.Width = 80; $btnCancelConfirm.Height = 30
    $btnCancelConfirm.Text = "취소"
    $btnCancelConfirm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $prompt.Controls.AddRange(@($lbl, $btnSaveConfirm, $btnCancelConfirm))
    $prompt.AcceptButton = $btnSaveConfirm
    $prompt.CancelButton = $btnCancelConfirm

    return $prompt.ShowDialog()
}

function Show-FssAiSettings {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "AI 요약 설정 (Google Gemini API)"
    $dlg.ClientSize = New-Object System.Drawing.Size(480, 300)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $config = Get-FssAiConfig

    $lblKey = New-Object System.Windows.Forms.Label
    $lblKey.Left = 15; $lblKey.Top = 15; $lblKey.Width = 130
    $lblKey.Text = "Gemini API 키:"
    $txtKey = New-Object System.Windows.Forms.TextBox
    $txtKey.Left = 150; $txtKey.Top = 12; $txtKey.Width = 315
    $txtKey.PasswordChar = '*'
    $txtKey.Text = $config.ApiKey

    $lblModel = New-Object System.Windows.Forms.Label
    $lblModel.Left = 15; $lblModel.Top = 48; $lblModel.Width = 130
    $lblModel.Text = "모델 이름:"
    $txtModel = New-Object System.Windows.Forms.TextBox
    $txtModel.Left = 150; $txtModel.Top = 45; $txtModel.Width = 315
    $txtModel.Text = $config.Model

    $lblBudget = New-Object System.Windows.Forms.Label
    $lblBudget.Left = 15; $lblBudget.Top = 81; $lblBudget.Width = 130
    $lblBudget.Text = "일일 토큰 한도:"
    $txtBudget = New-Object System.Windows.Forms.TextBox
    $txtBudget.Left = 150; $txtBudget.Top = 78; $txtBudget.Width = 150
    $txtBudget.Text = "$($config.DailyTokenBudget)"

    $usage = Get-FssAiUsage
    $lblUsage = New-Object System.Windows.Forms.Label
    $lblUsage.Left = 15; $lblUsage.Top = 108; $lblUsage.Width = 450
    $lblUsage.Text = "오늘($($usage.Date)) 누적 사용량: $($usage.TokensUsed) 토큰"

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Left = 15; $lblHint.Top = 136; $lblHint.Width = 450; $lblHint.Height = 150
    $lblHint.Text = "API 키는 https://aistudio.google.com/apikey 에서 구글 계정으로 무료 발급 가능합니다.`r`n그리드의 'AI요약' 버튼을 누른 항목에 한해서만, 그 순간에만 API가 호출됩니다(자동 호출 없음).`r`n`r`n일일 토큰 한도는 이 도구가 로컬에서 누적 집계한 사용량과 비교하는 값이며(자정 기준 매일 초기화), 구글 서버의 실제 잔여 할당량을 실시간으로 조회하는 것은 아닙니다. 누적 사용량이 이 한도의 90%를 넘으면 요약 기능이 자동으로 멈추고 경고창이 뜹니다. 무료 등급의 실제 한도는 모델/시점마다 다르니 https://ai.google.dev/gemini-api/docs/rate-limits 에서 확인 후 이 값을 조정하세요.`r`n(무료 등급은 한도 초과 시 요청이 거부(오류)될 뿐 자동으로 유료 결제되지 않습니다 - 별도로 유료 결제를 연결한 경우만 과금됩니다.)"

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Left = 15; $btnSave.Top = 255; $btnSave.Width = 100; $btnSave.Height = 30
    $btnSave.Text = "저장"
    $btnSave.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Left = 125; $btnClose.Top = 255; $btnClose.Width = 100; $btnClose.Height = 30
    $btnClose.Text = "닫기"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $dlg.Controls.AddRange(@($lblKey, $txtKey, $lblModel, $txtModel, $lblBudget, $txtBudget, $lblUsage, $lblHint, $btnSave, $btnClose))
    $dlg.AcceptButton = $btnSave
    $dlg.CancelButton = $btnClose

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $budgetVal = 1000000
        [void][int]::TryParse($txtBudget.Text.Trim(), [ref]$budgetVal)
        if ($budgetVal -le 0) { $budgetVal = 1000000 }
        $newConfig = [PSCustomObject]@{
            ApiKey           = $txtKey.Text.Trim()
            Model            = $(if ($txtModel.Text.Trim()) { $txtModel.Text.Trim() } else { "gemini-2.0-flash" })
            DailyTokenBudget = $budgetVal
        }
        Save-FssAiConfig -Config $newConfig
    }
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
    $script:NotifyDirty = $false

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
    $lblMenus.Left = 15; $lblMenus.Top = 214; $lblMenus.Width = 330
    $lblMenus.Text = "모니터링할 메뉴 (표 형태인 '회계법인 정보 통합조회'는 제외):"

    $btnSelectAllMenus = New-Object System.Windows.Forms.Button
    $btnSelectAllMenus.Left = 345; $btnSelectAllMenus.Top = 211; $btnSelectAllMenus.Width = 85; $btnSelectAllMenus.Height = 22
    $btnSelectAllMenus.Text = "전체선택"

    $btnDeselectAllMenus = New-Object System.Windows.Forms.Button
    $btnDeselectAllMenus.Left = 435; $btnDeselectAllMenus.Top = 211; $btnDeselectAllMenus.Width = 105; $btnDeselectAllMenus.Height = 22
    $btnDeselectAllMenus.Text = "전체선택해제"

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
        $lblMenus, $btnSelectAllMenus, $btnDeselectAllMenus, $clbMenus, $lblTaskStatus, $btnSave, $btnTestMail, $btnRegisterTask, $btnUnregisterTask, $txtNotifyLog
    ))

    $btnSelectAllMenus.Add_Click({
        for ($idx = 0; $idx -lt $clbMenus.Items.Count; $idx++) { $clbMenus.SetItemChecked($idx, $true) }
    })

    $btnDeselectAllMenus.Add_Click({
        for ($idx = 0; $idx -lt $clbMenus.Items.Count; $idx++) { $clbMenus.SetItemChecked($idx, $false) }
    })

    # 변경 여부 추적 (X로 닫을 때 저장 확인 팝업을 띄우기 위함)
    $markDirty = { $script:NotifyDirty = $true }
    $chkEnabled.Add_CheckedChanged($markDirty)
    $txtSender.Add_TextChanged($markDirty)
    $txtAppPw.Add_TextChanged($markDirty)
    $txtRecipients.Add_TextChanged($markDirty)
    $txtSmtpHost.Add_TextChanged($markDirty)
    $txtSmtpPort.Add_TextChanged($markDirty)
    $chkSsl.Add_CheckedChanged($markDirty)
    $cboInterval.Add_SelectedIndexChanged($markDirty)
    $clbMenus.Add_ItemCheck($markDirty)

    $dlg.Add_FormClosing({
        param($s, $e)
        if ($script:NotifyDirty) {
            $choice = Show-FssUnsavedChangesPrompt
            if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
                try {
                    $newConfig = Get-ConfigFromForm
                    Save-FssNotifyConfig -Config $newConfig
                    $script:NotifyDirty = $false
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("저장 실패: $($_.Exception.Message)", "오류", "OK", "Error") | Out-Null
                    $e.Cancel = $true
                }
            } else {
                $e.Cancel = $true
            }
        }
    })

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
            $script:NotifyDirty = $false
            Write-NotifyDlgLog "설정을 저장했습니다: $(Get-FssNotifyConfigPath)"
        } catch {
            Write-NotifyDlgLog "저장 실패: $($_.Exception.Message)"
        }
    })

    $btnTestMail.Add_Click({
        try {
            $newConfig = Get-ConfigFromForm
            Save-FssNotifyConfig -Config $newConfig
            $script:NotifyDirty = $false
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
            $script:NotifyDirty = $false
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
