$script:FssBaseUrl = "https://www.fss.or.kr"

function Get-FssBoardList {
    param(
        [Parameter(Mandatory)] [string]$BbsId,
        [Parameter(Mandatory)] [string]$MenuNo,
        [int]$PageIndex = 1,
        [string]$SearchWrd = "",
        [string]$ExtraParams = ""
    )

    $url = "$script:FssBaseUrl/fss/bbs/$BbsId/list.do?menuNo=$MenuNo&pageIndex=$PageIndex"
    if ($ExtraParams) { $url += "&$ExtraParams" }
    if ($SearchWrd) { $url += "&searchWrd=" + [uri]::EscapeDataString($SearchWrd) }

    $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30
    $html = $resp.Content

    # 페이지 하단 페이지네이션 위젯(예: <a ... data-pageindex='24'>끝 목록</a>)에 총 페이지 수가 이미 노출되어 있어서
    # 진행률 계산용으로 파싱해 둠(전 페이지를 실제로 걷지 않고도 첫 페이지에서 총 페이지 수를 알 수 있음).
    $script:LastListTotalPages = $null
    $pageNums = [regex]::Matches($html, "data-pageindex='(\d+)'") | ForEach-Object { [int]$_.Groups[1].Value }
    if ($pageNums) { $script:LastListTotalPages = ($pageNums | Measure-Object -Maximum).Maximum }

    $rows = New-Object System.Collections.Generic.List[object]

    # 게시글 제목/링크: <a href="/fss/bbs/{BbsId}/view.do?nttId=12345&...">제목</a>
    $pattern = '<a[^>]*href="[^"]*/bbs/' + [regex]::Escape($BbsId) + '/view\.do\?nttId=(\d+)[^"]*"[^>]*>\s*([^<]+?)\s*</a>'
    $matches = [regex]::Matches($html, $pattern)

    foreach ($m in $matches) {
        $nttId = $m.Groups[1].Value
        $title = [System.Net.WebUtility]::HtmlDecode($m.Groups[2].Value.Trim())
        if ([string]::IsNullOrWhiteSpace($title)) { continue }

        # 같은 nttId가 목록 안에서 중복 매칭되는 경우(모바일용 마크업 등) 제거
        if ($rows | Where-Object { $_.NttId -eq $nttId }) { continue }

        $rows.Add([PSCustomObject]@{
            Title    = $title
            NttId    = $nttId
            BbsId    = $BbsId
            MenuNo   = $MenuNo
            ViewUrl  = "$script:FssBaseUrl/fss/bbs/$BbsId/view.do?nttId=$nttId&menuNo=$MenuNo"
        })
    }

    return $rows
}

function Get-FssBoardAllItems {
    # 게시판마다 서버 자체 검색(searchWrd/searchCnd) 구성이 달라 신뢰할 수 없으므로(예: 심사·감리지적사례는
    # 제목이 아니라 쟁점분야/관련기준서/결정년도만 검색됨), 전 페이지를 가져와 호출부에서 제목으로 직접 필터링할 때 사용.
    # $OnProgress -> & $OnProgress <현재페이지> <총페이지 또는 $null> <퍼센트(int) 또는 $null>
    # $script:CancelRequested가 $true가 되면 중간에 중단하고 지금까지 모은 항목만 반환.
    param(
        [Parameter(Mandatory)] [string]$BbsId,
        [Parameter(Mandatory)] [string]$MenuNo,
        [string]$ExtraParams = "",
        [scriptblock]$OnProgress = $null
    )

    $pageIndex = 1
    $totalPages = $null
    $allItems = New-Object System.Collections.Generic.List[object]
    while ($true) {
        if ($script:CancelRequested) { break }
        $items = Get-FssBoardList -BbsId $BbsId -MenuNo $MenuNo -PageIndex $pageIndex -ExtraParams $ExtraParams
        if ($pageIndex -eq 1 -and $script:LastListTotalPages) { $totalPages = $script:LastListTotalPages }
        if ($items.Count -eq 0) { break }
        foreach ($it in $items) { $allItems.Add($it) }
        if ($OnProgress) {
            $pct = if ($totalPages) { [Math]::Min(100, [int](($pageIndex / $totalPages) * 100)) } else { $null }
            & $OnProgress $pageIndex $totalPages $pct
        }
        if ($items.Count -lt 10) { break }
        $pageIndex++
        if ($pageIndex -gt 300) { break }
    }
    return $allItems
}

function ConvertFrom-FssDetailHtml {
    # 상세페이지 HTML -> RegDate/Files/BodyText 파싱. 순수 함수(네트워크 호출/스크립트 스코프 의존 없음)로 만들어
    # Get-FssBoardDetail과 Get-FssBoardBodiesParallel(병렬 본문조회, 별도 실행 스레드) 양쪽에서 재사용.
    param(
        [Parameter(Mandatory)] [string]$Html,
        [Parameter(Mandatory)] [string]$NttId,
        [Parameter(Mandatory)] [string]$BaseUrl
    )

    # 등록일: 상세 페이지 내 '등록일' 레이블 뒤 날짜 패턴
    $regDate = ""
    $dateMatch = [regex]::Match($Html, '등록일[^0-9]{0,20}(\d{4}[-.]\d{2}[-.]\d{2})')
    if ($dateMatch.Success) { $regDate = $dateMatch.Groups[1].Value }

    # 첨부파일: fileDown.do?menuNo=...&atchFileId=...&fileSn=...
    # 목록 페이지 스타일: <a ... title="파일명 다운로드">
    # 상세 페이지 스타일: <a ...><span class="file-name"><span class="name">파일명 ... <span>
    $files = New-Object System.Collections.Generic.List[object]
    $seen = @{}

    $filePatterns = @(
        '<a[^>]*href="([^"]*fileDown\.do\?[^"]*atchFileId=([a-zA-Z0-9]+)&fileSn=(\d+)[^"]*)"[^>]*title="([^"]+?)\s*다운로드"',
        '<a[^>]*href="([^"]*fileDown\.do\?[^"]*atchFileId=([a-zA-Z0-9]+)&fileSn=(\d+)[^"]*)"[^>]*>\s*<span class="file-name">\s*(?:<i[^>]*></i>\s*)?<span class="name">\s*([^<]+?)\s*(?:\r?\n|<span)'
    )

    foreach ($filePattern in $filePatterns) {
        $fileMatches = [regex]::Matches($Html, $filePattern)
        foreach ($fm in $fileMatches) {
            $key = "$($fm.Groups[2].Value)_$($fm.Groups[3].Value)"
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true

            $relUrl = [System.Net.WebUtility]::HtmlDecode($fm.Groups[1].Value)
            $downloadUrl = if ($relUrl -match '^https?://') { $relUrl } else { "$BaseUrl$relUrl" }
            $files.Add([PSCustomObject]@{
                AtchFileId  = $fm.Groups[2].Value
                FileSn      = $fm.Groups[3].Value
                FileName    = [System.Net.WebUtility]::HtmlDecode($fm.Groups[4].Value.Trim())
                DownloadUrl = $downloadUrl
            })
        }
    }

    # 본문 텍스트: 게시글 본문은 <div class="dbdata">...</div> 안에 있고(실사이트 확인),
    # 바로 뒤에 담당부서/문의 <dl>이 오는 게 일반적이라 그 지점까지를 본문으로 간주.
    # 담당부서 블록이 없는 게시글(일부 정보성/짧은 글)은 그 다음 목록 버튼(class="btn-set")까지로 대체.
    $bodyText = ""
    $bodyMatch = [regex]::Match($Html, '<div class="dbdata">(.*?)<dl>\s*<dt>담당부서', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if (-not $bodyMatch.Success) {
        $bodyMatch = [regex]::Match($Html, '<div class="dbdata">(.*?)<div class="btn-set"', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    }
    if ($bodyMatch.Success) {
        $stripped = [regex]::Replace($bodyMatch.Groups[1].Value, '<[^>]+>', ' ')
        $stripped = [System.Net.WebUtility]::HtmlDecode($stripped)
        $bodyText = ([regex]::Replace($stripped, '\s+', ' ')).Trim()
    }

    return [PSCustomObject]@{
        NttId    = $NttId
        RegDate  = $regDate
        Files    = $files
        BodyText = $bodyText
    }
}

function Get-FssBoardDetail {
    param(
        [Parameter(Mandatory)] [string]$BbsId,
        [Parameter(Mandatory)] [string]$MenuNo,
        [Parameter(Mandatory)] [string]$NttId
    )

    $url = "$script:FssBaseUrl/fss/bbs/$BbsId/view.do?nttId=$NttId&menuNo=$MenuNo"
    $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30
    return ConvertFrom-FssDetailHtml -Html $resp.Content -NttId $NttId -BaseUrl $script:FssBaseUrl
}

function Get-FssBoardBodiesParallel {
    # 검색어가 제목에 없는 게시글들의 "본문"까지 검색하려면 상세페이지를 건마다 조회해야 해서
    # 순차 처리하면 게시판 하나(수백 건)당 수십 초~수 분이 걸림. HttpClient 비동기 요청을
    # 동시에 여러 건(기본 8개) 진행시켜 시간을 단축. 파싱은 가벼운 정규식이라 메인 스레드에서
    # 완료된 요청을 폴링하며 처리(별도 러너스페이스/스레드 없이 WinForms STA 스레드 안에서 안전하게 동작).
    # 반환: "{BbsId}_{NttId}" -> (ConvertFrom-FssDetailHtml 결과) 해시테이블. 일부 항목이 실패해도
    # 나머지는 그대로 반환(본문검색은 최선노력이므로 개별 실패가 전체를 막지 않음).
    param(
        [Parameter(Mandatory)] [object[]]$Items,   # 각 항목: BbsId, MenuNo, NttId
        [int]$MaxConcurrent = 8,
        [scriptblock]$OnProgress = $null
    )
    Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

    $results = @{}
    if ($Items.Count -eq 0) { return $results }

    $client = New-Object System.Net.Http.HttpClient
    $client.Timeout = [TimeSpan]::FromSeconds(30)

    $queue = New-Object System.Collections.Generic.Queue[object]
    foreach ($it in $Items) { $queue.Enqueue($it) }
    $inFlight = New-Object System.Collections.Generic.List[object]
    $completed = 0
    $total = $Items.Count

    try {
        while ($queue.Count -gt 0 -or $inFlight.Count -gt 0) {
            if ($script:CancelRequested) { break }
            while ($inFlight.Count -lt $MaxConcurrent -and $queue.Count -gt 0) {
                $it = $queue.Dequeue()
                $url = "$script:FssBaseUrl/fss/bbs/$($it.BbsId)/view.do?nttId=$($it.NttId)&menuNo=$($it.MenuNo)"
                $task = $client.GetStringAsync($url)
                $inFlight.Add([PSCustomObject]@{ Item = $it; Task = $task })
            }
            Start-Sleep -Milliseconds 30
            for ($i = $inFlight.Count - 1; $i -ge 0; $i--) {
                $ent = $inFlight[$i]
                if ($ent.Task.IsCompleted) {
                    $inFlight.RemoveAt($i)
                    $completed++
                    if (-not $ent.Task.IsFaulted -and -not $ent.Task.IsCanceled) {
                        try {
                            $html = $ent.Task.GetAwaiter().GetResult()
                            $detail = ConvertFrom-FssDetailHtml -Html $html -NttId $ent.Item.NttId -BaseUrl $script:FssBaseUrl
                            $results["$($ent.Item.BbsId)_$($ent.Item.NttId)"] = $detail
                        } catch {
                            # 개별 항목 파싱 실패 - 건너뜀
                        }
                    }
                    if ($OnProgress) { & $OnProgress $completed $total ([int](($completed / [Math]::Max(1, $total)) * 100)) }
                }
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
    } finally {
        $client.Dispose()
    }
    return $results
}

function Save-FssAttachment {
    param(
        [Parameter(Mandatory)] [string]$DownloadUrl,
        [Parameter(Mandatory)] [string]$OutDir,
        [string]$PreferredFileName = ""
    )

    if (-not (Test-Path $OutDir)) {
        New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
    }

    $tempPath = Join-Path $OutDir ([System.IO.Path]::GetRandomFileName())
    $resp = Invoke-WebRequest -Uri $DownloadUrl -UseBasicParsing -TimeoutSec 60 -OutFile $tempPath -PassThru

    $fileName = $PreferredFileName
    if (-not $fileName) {
        $cd = $resp.Headers["Content-Disposition"]
        if ($cd -match 'filename\*?=(?:UTF-8'')?"?([^;"]+)"?') {
            $fileName = [System.Net.WebUtility]::UrlDecode($matches[1])
        } else {
            $fileName = [System.IO.Path]::GetFileName($tempPath)
        }
    }

    $fileName = ($fileName -replace '[\\/:*?"<>|]', '_')
    $finalPath = Join-Path $OutDir $fileName

    if (Test-Path $finalPath) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        $ext = [System.IO.Path]::GetExtension($fileName)
        $i = 1
        while (Test-Path $finalPath) {
            $finalPath = Join-Path $OutDir "$base($i)$ext"
            $i++
        }
    }

    Move-Item -Path $tempPath -Destination $finalPath -Force
    return $finalPath
}

function Get-FssJobListRows {
    param(
        [Parameter(Mandatory)] [string]$JobPath,
        [Parameter(Mandatory)] [string]$MenuNo,
        [int]$PageIndex = 1,
        [string]$SearchWrd = "",
        [string]$Sdate = "",
        [string]$Edate = "",
        [string]$ExtraParams = ""
    )

    $url = "$script:FssBaseUrl/fss/job/$JobPath/list.do?menuNo=$MenuNo&pageIndex=$PageIndex"
    if ($ExtraParams) { $url += "&$ExtraParams" }
    if ($Sdate) { $url += "&sdate=" + [uri]::EscapeDataString($Sdate) }
    if ($Edate) { $url += "&edate=" + [uri]::EscapeDataString($Edate) }
    if ($SearchWrd) { $url += "&searchWrd=" + [uri]::EscapeDataString($SearchWrd) }

    $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30
    $html = $resp.Content

    $script:LastListTotalPages = $null
    $pageNums = [regex]::Matches($html, "data-pageindex='(\d+)'") | ForEach-Object { [int]$_.Groups[1].Value }
    if ($pageNums) { $script:LastListTotalPages = ($pageNums | Measure-Object -Maximum).Maximum }

    $rows = New-Object System.Collections.Generic.List[object]

    # accnutAdtorInfo형 목록 행: <td class="num">N</td><td><a href=".../view.do?...XxxSlno=123">제목</a></td> 이후 일반 <td>들 + 첨부파일(class="file-single")
    $rowPattern = '<tr>\s*<td class="num">(\d+)</td>\s*<td><a href="[^"]*[a-zA-Z]+Slno=(\d+)"[^>]*>\s*([^<]+?)\s*</a></td>(.*?)</tr>'
    $rowMatches = [regex]::Matches($html, $rowPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)

    foreach ($m in $rowMatches) {
        $slno = $m.Groups[2].Value
        $title = [System.Net.WebUtility]::HtmlDecode($m.Groups[3].Value.Trim())
        $rest = $m.Groups[4].Value

        $plainTds = [regex]::Matches($rest, '<td>([^<]*)</td>')
        $cols = New-Object System.Collections.Generic.List[string]
        foreach ($p in $plainTds) { $cols.Add([System.Net.WebUtility]::HtmlDecode($p.Groups[1].Value.Trim())) }

        # 조치일(YYYYMMDD 8자리)로 보이는 컬럼을 우선 채택, 없으면 두번째 컬럼을 날짜로 간주
        $actionDate = $cols | Where-Object { $_ -match '^\d{8}$' } | Select-Object -First 1
        if (-not $actionDate -and $cols.Count -ge 2) { $actionDate = $cols[1] }

        $filePattern = '<a[^>]*href="([^"]*fileDown\.do\?[^"]*atchFileId=([a-zA-Z0-9]+)&fileSn=(\d+)[^"]*)"[^>]*class="file-single"[^>]*title="([^"]+?)\s*다운로드"'
        $fileMatches = [regex]::Matches($rest, $filePattern)
        $files = New-Object System.Collections.Generic.List[object]
        foreach ($fm in $fileMatches) {
            $relUrl = [System.Net.WebUtility]::HtmlDecode($fm.Groups[1].Value)
            $downloadUrl = if ($relUrl -match '^https?://') { $relUrl } else { "$script:FssBaseUrl$relUrl" }
            $files.Add([PSCustomObject]@{
                AtchFileId  = $fm.Groups[2].Value
                FileSn      = $fm.Groups[3].Value
                FileName    = [System.Net.WebUtility]::HtmlDecode($fm.Groups[4].Value.Trim())
                DownloadUrl = $downloadUrl
            })
        }

        $rows.Add([PSCustomObject]@{
            Title      = $title
            Slno       = $slno
            ActionDate = $actionDate
            Columns    = $cols
            Files      = $files
            JobPath    = $JobPath
            MenuNo     = $MenuNo
        })
    }

    return $rows
}

function Get-FssJobAllListRows {
    # Get-FssBoardAllItems와 동일한 이유(게시판별 서버 검색 신뢰 불가)로 전 페이지를 가져와
    # 호출부에서 제목으로 직접 필터링할 때 사용. Sdate/Edate는 서버가 실제로 필터링하는 것으로
    # 확인됐으므로(Phase 2 검증) 그대로 전달.
    param(
        [Parameter(Mandatory)] [string]$JobPath,
        [Parameter(Mandatory)] [string]$MenuNo,
        [string]$Sdate = "",
        [string]$Edate = "",
        [string]$ExtraParams = "",
        [scriptblock]$OnProgress = $null
    )

    $pageIndex = 1
    $totalPages = $null
    $allRows = New-Object System.Collections.Generic.List[object]
    while ($true) {
        if ($script:CancelRequested) { break }
        $rows = Get-FssJobListRows -JobPath $JobPath -MenuNo $MenuNo -PageIndex $pageIndex -Sdate $Sdate -Edate $Edate -ExtraParams $ExtraParams
        if ($pageIndex -eq 1 -and $script:LastListTotalPages) { $totalPages = $script:LastListTotalPages }
        if ($rows.Count -eq 0) { break }
        foreach ($r in $rows) { $allRows.Add($r) }
        if ($OnProgress) {
            $pct = if ($totalPages) { [Math]::Min(100, [int](($pageIndex / $totalPages) * 100)) } else { $null }
            & $OnProgress $pageIndex $totalPages $pct
        }
        if ($rows.Count -lt 10) { break }
        $pageIndex++
        if ($pageIndex -gt 300) { break }
    }
    return $allRows
}

function Save-FssJobSeries {
    param(
        [Parameter(Mandatory)] [string]$JobPath,
        [Parameter(Mandatory)] [string]$MenuNo,
        [Parameter(Mandatory)] [string]$OutDir,
        [string]$SearchWrd = "",
        [string]$Sdate = "",
        [string]$Edate = "",
        [string]$ExtraParams = "",
        [scriptblock]$OnProgress = $null
    )

    if (-not (Test-Path $OutDir)) {
        New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
    }

    $manifest = New-Object System.Collections.Generic.List[string]
    $manifest.Add("# $JobPath (menuNo=$MenuNo) 다운로드 목록")
    $manifest.Add("생성일시: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $manifest.Add("")
    $manifest.Add("| 조치일 | 제목 | 저장된 파일 |")
    $manifest.Add("|---|---|---|")

    $pageIndex = 1
    $allRows = New-Object System.Collections.Generic.List[object]
    while ($true) {
        $rows = Get-FssJobListRows -JobPath $JobPath -MenuNo $MenuNo -PageIndex $pageIndex -SearchWrd $SearchWrd -Sdate $Sdate -Edate $Edate -ExtraParams $ExtraParams
        if ($rows.Count -eq 0) { break }
        foreach ($r in $rows) { $allRows.Add($r) }
        if ($rows.Count -lt 10) { break }
        $pageIndex++
        if ($pageIndex -gt 50) { break }
    }

    foreach ($row in $allRows) {
        if ($OnProgress) { & $OnProgress "처리 중: $($row.Title)" }
        $savedNames = New-Object System.Collections.Generic.List[string]

        foreach ($f in $row.Files) {
            if ($OnProgress) { & $OnProgress "다운로드 중: $($f.FileName)" }
            $saved = Save-FssAttachment -DownloadUrl $f.DownloadUrl -OutDir $OutDir -PreferredFileName $f.FileName
            $savedNames.Add([System.IO.Path]::GetFileName($saved))
        }

        $manifest.Add("| $($row.ActionDate) | $($row.Title) | $($savedNames -join ', ') |")
    }

    $manifestPath = Join-Path $OutDir "manifest.md"
    $manifest | Out-File -FilePath $manifestPath -Encoding utf8
    return $manifestPath
}

function Save-FssJobExcel {
    param(
        [Parameter(Mandatory)] [string]$JobPath,
        [Parameter(Mandatory)] [string]$MenuNo,
        [Parameter(Mandatory)] [string]$OutPath,
        [string]$SearchStr = "",
        [string]$Sdate = "",
        [string]$Edate = "",
        [string]$ExtraParams = ""
    )

    # cprCoreInfo형: excelDown.do는 화면의 페이지네이션과 무관하게 검색조건(sdate/edate/searchStr)에 해당하는 전체 결과를 반환함.
    # sdate/edate를 비우면 전체 기간이 조회됨(사이트 화면의 "1년 이내" 제한은 클라이언트 JS 검증일 뿐 서버는 강제하지 않음, 실사이트 확인 완료).
    $url = "$script:FssBaseUrl/fss/job/$JobPath/excelDown.do?menuNo=$MenuNo&pageIndex=1&viewType=CONTBODY"
    if ($ExtraParams) { $url += "&$ExtraParams" }
    $url += "&sdate=" + [uri]::EscapeDataString($Sdate)
    $url += "&edate=" + [uri]::EscapeDataString($Edate)
    $url += "&searchStr=" + [uri]::EscapeDataString($SearchStr)

    Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 90 -OutFile $OutPath | Out-Null
    return $OutPath
}

function Get-FssRevisionHint {
    param([Parameter(Mandatory)] [string]$Title)

    if ($Title -match '공개초안|의견\s*조회|의견\s*요청') { return "공개초안/의견조회 (개정 예고, 확정본 아님)" }
    if ($Title -match '개정') { return "개정 - 전문 여부 확인 필요" }
    if ($Title -match '전문|제정') { return "전문(추정)" }
    return ""
}

function Save-FssSeries {
    param(
        [Parameter(Mandatory)] [string]$BbsId,
        [Parameter(Mandatory)] [string]$MenuNo,
        [Parameter(Mandatory)] [string]$OutDir,
        [string]$ExtraParams = "",
        [scriptblock]$OnProgress = $null
    )

    if (-not (Test-Path $OutDir)) {
        New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
    }

    $manifest = New-Object System.Collections.Generic.List[string]
    $manifest.Add("# $BbsId (menuNo=$MenuNo) 다운로드 목록")
    $manifest.Add("생성일시: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $manifest.Add("")
    $manifest.Add("| 등록일 | 제목 | 추정유형 | 저장된 파일 |")
    $manifest.Add("|---|---|---|---|")

    $pageIndex = 1
    $allItems = New-Object System.Collections.Generic.List[object]
    while ($true) {
        $items = Get-FssBoardList -BbsId $BbsId -MenuNo $MenuNo -PageIndex $pageIndex -ExtraParams $ExtraParams
        if ($items.Count -eq 0) { break }
        foreach ($it in $items) { $allItems.Add($it) }
        if ($items.Count -lt 10) { break }
        $pageIndex++
        if ($pageIndex -gt 50) { break }
    }

    foreach ($item in $allItems) {
        if ($OnProgress) { & $OnProgress "조회 중: $($item.Title)" }
        $detail = Get-FssBoardDetail -BbsId $BbsId -MenuNo $MenuNo -NttId $item.NttId
        $hint = Get-FssRevisionHint -Title $item.Title
        $savedNames = New-Object System.Collections.Generic.List[string]

        foreach ($f in $detail.Files) {
            if ($OnProgress) { & $OnProgress "다운로드 중: $($f.FileName)" }
            $saved = Save-FssAttachment -DownloadUrl $f.DownloadUrl -OutDir $OutDir -PreferredFileName $f.FileName
            $savedNames.Add([System.IO.Path]::GetFileName($saved))
        }

        $manifest.Add("| $($detail.RegDate) | $($item.Title) | $hint | $($savedNames -join ', ') |")
    }

    $manifestPath = Join-Path $OutDir "manifest.md"
    $manifest | Out-File -FilePath $manifestPath -Encoding utf8
    return $manifestPath
}
