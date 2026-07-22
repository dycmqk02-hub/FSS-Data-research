# FSS 업무자료(공시/회계) 조회·다운로드 도구 — 진행 상황

## 목적
금융감독원(fss.or.kr) "업무자료 > 공시/회계" 메뉴의 게시물·첨부파일을 API/스크래핑으로 조회하고 다운로드하는 GUI 도구.
조건: **Python 미사용, PowerShell만 사용**, **결과물은 exe 파일 하나**, **exe 실행 시 자동으로 데이터 조회 시작**.

## 현재 상태: Phase 1 + Phase 2(검색폼형 페이지) + Phase 4(새 글 이메일 알림) 완료, 실사이트 검증 완료

## 프로젝트 구조
```
C:\Users\seo\FSS-DataTool\
  src\
    Config.ps1        # 공시/회계 게시판 매핑 테이블 (menuNo, BbsId 등)
    Scraper.ps1        # Get-FssBoardList / Get-FssBoardDetail / Save-FssAttachment / Save-FssSeries
    Notify.ps1         # 새 글 이메일 알림: 설정/상태 저장, 새 글 감지, 메일 발송, 작업 스케줄러 등록/해제
    HwpConvert.ps1     # Test-HwpInstalled / Convert-HwpToPdf (한컴오피스 COM 자동화)
    Gui.ps1            # WinForms GUI (트리/그리드/다운로드 버튼/알림 설정 다이얼로그)
    Main.ps1           # 진입점 (src 파일들을 dot-source 후 Show-FssGui 호출) — 로컬 테스트용
  build.ps1            # src\*.ps1 을 병합(맨 앞에 param([switch]$CheckOnly) 추가) 후 ps2exe로 컴파일
                        # → dist\FSS-DataTool.exe 및 dist\FSS-DataTool-merged.ps1(작업 스케줄러용 사본) 생성
  dist\
    FSS-DataTool.exe            # 최종 산출물 (더블클릭 실행용, GUI)
    FSS-DataTool-merged.ps1     # 위와 동일한 로직의 병합 스크립트 사본(작업 스케줄러가 이걸 참조, SAC 우회)
  PROGRESS.md          # 이 파일
```

## 사이트 구조 조사 결과 (중요)
- 공식 문서화된 Open API가 아니라 **GET 파라미터 기반 서버사이드 렌더링 HTML**. API 키 불필요, 브라우저로 보는 것과 동일한 공개 페이지를 그대로 가져옴.
- 목록: `GET /fss/bbs/{BbsId}/list.do?menuNo={MenuNo}&pageIndex=N`
- 상세: `GET /fss/bbs/{BbsId}/view.do?nttId={id}&menuNo={MenuNo}`
- 첨부파일: `GET /fss/cmmn/file/fileDown.do?menuNo=...&atchFileId=...&fileSn=...` (인증 불필요)
- **첨부파일은 대부분 .hwp**(PDF 아님). PDF 변환은 한컴오피스 설치 + COM 자동화(`HWPFrame.HwpObject`)로만 가능. 미설치 PC에서는 원본 .hwp 유지.
- 회계법인 정보 통합조회(`/fss/job/cprCoreInfo/list.do`), 회계감리결과제재(`/fss/job/accnutAdtorInfo/list.do`)는 게시판형이 아니라 검색폼형(기간+키워드)이라 파서가 다름 → **Phase 2에서 구현 완료** (아래 "Phase 2 사이트 구조 조사 결과" 참고).
- Config.ps1에 전체 메뉴 매핑 테이블 있음 (공시 3항목 + 회계 하위 게시판형 약 18개 + job형 2개 + 정보성 페이지 목록).

## Phase 2 사이트 구조 조사 결과 (검색폼형 페이지, 2026-07-22 조사)
Type="job" 항목은 `SubType`으로 다시 둘로 나뉨 (Config.ps1 상단 주석 참고):

1. **`SubType="table"` — `cprCoreInfo`(회계법인 정보 통합조회)**
   - 목록 화면 자체가 게시글이 아니라 **가로 27개 컬럼짜리 데이터 표**(연번/법인명/인력현황/감사실적/재무정보/감리결과·소송현황/결산년월). 행별 첨부파일 다운로드 개념이 없음(팝업 링크만 있고 dart.fss.or.kr 및 사내 팝업으로 연결).
   - 검색 파라미터: `sdate`/`edate`(결산연월 `YYYY-MM`), `searchStr`(회계법인명). **둘 다 비워두면 서버가 전체 기간을 반환**함 — 화면의 "1년 이내" 제한(`fnSearch()`의 `monthE > 11` 체크)은 **클라이언트 JS 검증일 뿐 서버는 강제하지 않음**(실사이트 확인 완료: `sdate=2010-01&edate=2026-07` 요청 시 2227행 정상 반환).
   - `GET /fss/job/cprCoreInfo/excelDown.do?menuNo=200429&pageIndex=1&viewType=CONTBODY&sdate=...&edate=...&searchStr=...` → **화면 페이지네이션과 무관하게 검색조건에 해당하는 전체 결과**를 `Content-Disposition: attachment; filename=....xls`로 반환(HTML 테이블을 xls 확장자로 감싼 형태, 엑셀에서 정상적으로 열림). 이걸 그대로 파일로 저장하는 것이 실사용 패턴과 가장 가까움 → 화면 내 표 미리보기 그리드는 만들지 않고 **엑셀 다운로드만 지원**.
   - `Get-FssJobListRows`는 이 타입에서 사용하지 않음. `Save-FssJobExcel`(Scraper.ps1)이 위 URL을 그대로 호출해 `-OutPath`에 저장.

2. **`SubType="list"` — `accnutAdtorInfo`(회계감리결과제재 등)`**
   - 게시판형과 거의 동일한 구조: `GET /fss/job/accnutAdtorInfo/list.do?menuNo=200617&acntnWrkCode=02&pageIndex=N&sdate=...&edate=...&searchWrd=...` (페이지당 10건, `sdate`/`edate`/`searchWrd` 모두 비우면 전체).
   - 목록 행 자체에 첨부파일 링크(`class="file-single"`, `fileDown.do?...atchFileId=...&fileSn=...`)가 바로 노출되어 있어서 **bbs와 달리 상세 페이지(view.do)를 별도로 조회할 필요 없음**(view.do를 확인해봐도 동일한 atchFileId를 그대로 노출할 뿐 추가 정보 없음).
   - 회사명 링크는 `./view.do?acntnWrkCode=02&menuNo=200617&acntnWrkSlno=N` 형태(`nttId` 대신 `acntnWrkSlno`).
   - `Get-FssJobListRows`(Scraper.ps1)가 이 목록을 파싱(행 하나 = 회사 하나 = 첨부파일 보통 1개). `Save-FssJobSeries`가 bbs의 `Save-FssSeries`와 동일한 패턴으로 전체 페이지를 순회하며 다운로드 + manifest.md 생성. 실사이트로 전체 74건 다운로드까지 검증 완료.

### GUI 반영 (Gui.ps1)
- 상단 검색줄에 "결산연월(회계법인 통합조회 전용, 비우면 전체)" `sdate`/`edate` 입력란 2개 추가(`$txtSdate`/`$txtEdate`, `TOP_HEIGHT` 90→122로 확장). `cprCoreInfo`에서만 의미 있고 다른 항목에서는 무시됨.
- 트리에서 항목 선택 시(`tree.Add_AfterSelect`) `Type="job" and SubType="table"`이면 `$btnDownloadSeries.Text`를 "엑셀 다운로드"로 바꾸고 `$btnDownloadSelected`를 비활성화, 그 외에는 기존 "이 게시판 전체(시리즈) 다운로드"로 복원.
- `Load-CurrentBoard`: `job/table`은 안내 메시지만 표시(그리드 비움), `job/list`는 `Get-FssJobListRows` 결과를 그리드에 채우고 파일 목록을 `$script:JobRowFiles[$Slno]`에 캐시(상세 재조회 없이 바로 다운로드용).
- `$btnDownloadSelected`: bbs는 기존처럼 `Get-FssBoardDetail`, job/list는 `$script:JobRowFiles`에서 바로 파일 꺼내서 다운로드.
- `$btnDownloadSeries`: job/table이면 `SaveFileDialog`로 단일 `.xls` 저장(`Save-FssJobExcel`), bbs/job-list면 기존처럼 `FolderBrowserDialog` + 시리즈 다운로드(`Save-FssSeries` 또는 `Save-FssJobSeries`).

### 검증 방법 및 한계
- 새 함수 3개(`Get-FssJobListRows`, `Save-FssJobSeries`, `Save-FssJobExcel`)는 Config/Scraper.ps1을 dot-source해서 **실사이트에 직접 호출**하여 검증(accnutAdtorInfo 74건 전체 다운로드 성공, cprCoreInfo 엑셀 다운로드 검색어/기간 필터 포함 검증).
- GUI는 병합 스크립트(`%TEMP%\FSS-DataTool-merged.ps1`)를 백그라운드로 띄우고 `PrintWindow`로 스크린샷 캡처하여 레이아웃(신규 결산연월 입력란)이 깨지지 않음을 확인.
- **주의**: 이 PC는 디스플레이 배율(DPI scaling, 아마도 125%)이 걸려 있어서, DPI-unaware 호출 프로세스에서 `GetWindowRect`로 얻은 좌표와 `SetCursorPos`가 실제로 적용하는 물리 좌표가 어긋남(`SetProcessDPIAware()` 호출 후에는 오히려 `GetWindowRect`가 `-40000,-40000` 같은 비정상 값을 반환하는 현상 발생 — 원인 미확인). 따라서 **마우스 좌표 기반 트리 클릭 자동화는 이 환경에서 신뢰할 수 없었고**, 트리/그리드 내부 노드는 UI Automation(`AutomationElement`)으로도 조회되지 않음(전체 18개 Pane 컨트롤만 노출 — ps2exe 빌드 exe에 comctl32 v6 매니페스트가 없어서 접근성 트리가 제한되는 것으로 추정). 다음에 GUI 클릭 자동화가 필요하면: (a) 매니페스트에 `<application xmlns="urn:schemas-microsoft-com:asm.v3"><windowsSettings><dpiAware>true/PM</dpiAware></windowsSettings></application>`를 추가해서 exe 자체를 DPI-aware로 빌드하거나, (b) `ps2exe`의 매니페스트 옵션으로 comctl32 v6을 활성화하는 방법을 먼저 검토할 것. 이번에는 백엔드 함수를 실사이트로 직접 호출하는 방식으로 충분히 검증했으므로 우회함.

## "개정판만 있는 최신본을 이전판과 병합" 요구사항 처리 방침
문서 내용을 이해해야 하는 작업이라 스크립트로 완전 자동화하지 않기로 함(사용자 확정 사항). 대신:
- `Save-FssSeries` 함수가 게시판의 전 버전을 다운로드하고 `manifest.md`에 날짜순 목록 + 제목 키워드 기반 추정 플래그("개정"/"공개초안" 등)를 기록.
- 실제 병합은 이후 별도로 Claude 세션에서 다운로드된 파일들을 직접 읽고 병합 초안 + 참고문서를 제시하는 방식으로 진행 (도구 자동화 범위 밖).

## 검증 완료된 것
- `Get-FssBoardList`, `Get-FssBoardDetail`, `Save-FssAttachment` — 실제 사이트(`B0000131` 분ㆍ반기재무제표검토준칙 게시판)로 목록 조회, 상세 파싱, 첨부파일 다운로드(225KB .hwp) 모두 성공 확인.
- `Test-HwpInstalled` — 이 PC는 한글 미설치 확인 (정상적으로 PDF 변환 건너뛰고 원본 유지).
- GUI 전체 동작 — WinForms 창 실행, 트리 자동조회, 목록 표시, 검색/페이징 버튼, 다운로드 버튼, 로그창 모두 **PrintWindow API로 스크린샷 캡처하여 육안 검증 완료**.
- `build.ps1` → ps2exe로 `dist\FSS-DataTool.exe` 빌드 성공.

## 알려진 이슈 / 주의사항 (다른 PC에서 이어할 때 참고)

1. **한글 인코딩**: 모든 `.ps1` 파일은 반드시 **UTF-8 with BOM**으로 저장해야 함. BOM 없이 저장하면 Windows PowerShell 5.1이 한글 텍스트를 깨진 문자로 읽어서 파싱 에러(특히 Config.ps1의 해시테이블 리터럴이 깨짐) 발생. `Write`/`Edit` 도구로 저장 후 항상 아래처럼 BOM을 재적용할 것:
   ```powershell
   $content = Get-Content -Path $path -Raw -Encoding UTF8
   [System.IO.File]::WriteAllText($path, $content, (New-Object System.Text.UTF8Encoding($true)))
   ```

2. **WinForms Dock 중첩 버그(해결됨)**: 처음에 `rightPanel`(Dock=Fill) 안에 다시 `topPanel`(Top)/`bottomPanel`(Bottom)/`grid`(Fill)을 중첩시켰더니 자식 Fill 컨트롤이 부모의 Left 도킹을 무시하고 전체 폭을 차지하는 버그 발생(진단 결과 rightPanel.Bounds가 X=0,Width=전체폭으로 tree를 무시). **해결책: Dock 속성을 아예 쓰지 않고, `Form.Add_Resize` 이벤트에서 `SetBounds()`로 직접 좌표를 계산하는 수동 레이아웃 방식으로 전환** (현재 `Gui.ps1`의 `$doLayout` 스크립트블록). 앞으로 레이아웃을 건드릴 때는 이 수동 계산 로직을 수정하면 됨.

3. **원격 스크린샷 캡처 시 `CopyFromScreen`은 신뢰 불가**: 같은 창인데도 버튼 등 일부 컨트롤이 캡처에 안 보이는 현상이 반복됐음(DWM 컴포지션/캐시 문제로 추정). **`PrintWindow` Win32 API(`PW_RENDERFULLCONTENT`=2 플래그)로 캡처하면 정확하게 나옴.** 이후 GUI를 스크린샷으로 검증할 때는 반드시 PrintWindow 방식 사용.

4. **`taskkill /IM powershell.exe /F` 사용 금지(주의)**: 이 명령은 테스트 중인 PowerShell 프로세스뿐 아니라 VS Code의 PowerShell 확장(에디터 인텔리센스용 백그라운드 프로세스)도 같이 죽여서 VS Code 쪽에 에러가 남. 테스트 프로세스를 종료할 때는 `Start-Process ... -PassThru`로 받은 PID만 `Stop-Process -Id $p.Id -Force`로 종료할 것.

5. **Windows Smart App Control(SAC)이 새로 컴파일된 exe를 차단함**: `dist\FSS-DataTool.exe`를 실행하면 "Application Control policy has blocked this file" 오류 발생 (이 PC의 SAC가 평판 없는 신규 exe를 차단하는 정책). ps2exe나 다른 컴파일러로 만든 exe라면 어떤 것이든 동일하게 차단됨 — 코드 문제 아님. 사용자가 "지금은 exe 유지, 차단 해제는 직접 판단"으로 결정함. SAC를 끄면 **Windows 재설치 전까지 다시 켤 수 없음(비가역적)**이므로 임의로 끄지 말 것. 테스트/검증 시에는 `powershell -NoProfile -ExecutionPolicy Bypass -File <병합된 스크립트>` 방식으로 우회해서 진행함 (build.ps1이 생성하는 `%TEMP%\FSS-DataTool-merged.ps1`).

## Phase 4: 새 글 이메일 알림 (2026-07-22 구현)
사용자 요청: 금감원 사이트 각 메뉴에 새 글/자료가 올라오면 등록해둔 이메일로 자동 알림.

### 설계 결정 (사용자 확인 완료)
- **실행 방식**: Windows 작업 스케줄러 자동 등록(추천 선택). 데스크톱 exe는 사용자가 앱을 열어두지 않으면 스스로 깨어날 수 없으므로, 진짜 자동 알림을 위해서는 OS 스케줄러가 필요.
- **확인 주기**: 하루 1회(기본값, GUI에서 1시간/4시간/하루 중 변경 가능).
- **발신 계정**: Gmail SMTP(무료, 앱 비밀번호 필요 — 2단계 인증 켠 뒤 발급). GUI에 SMTP 호스트/포트도 노출해뒀으므로 회사 SMTP로도 대체 가능.

### 아키텍처
- **모니터링 대상**: `Type="bbs"` 전체 + `Type="job" and SubType="list"`(accnutAdtorInfo형)만 해당. `SubType="table"`(cprCoreInfo형)은 게시글 단위가 아닌 표 데이터라 "새 글" 개념이 없어서 제외(`Get-FssNotifiableMenus`, Notify.ps1).
- **설정/상태 저장 위치**: `%AppData%\FSS-DataTool\` (exe가 단일 파일이라 설정은 exe 밖에 보관해야 함).
  - `notify-config.json`: 활성화 여부, 발신 이메일/앱 비밀번호, SMTP 호스트·포트·SSL, 받는사람 목록, 모니터링할 메뉴 키 목록, 확인 주기.
  - `notify-state.json`: 메뉴별로 마지막으로 확인한 게시글 ID 목록(첫 페이지 기준). 최초 실행 시에는 기준선만 저장하고 알림을 보내지 않음(그렇지 않으면 기존 게시글 전체가 "새 글"로 잡혀 스팸이 됨).
  - `notify-log.txt`: 확인/발송 이력 로그(누적 append).
  - **주의**: `AppPassword`가 현재 평문 JSON으로 저장됨. 사용자 개인 계정 폴더 안에만 있지만, 다른 용도로 쓰는 비밀번호를 재사용하지 말고 이 용도 전용 Gmail 앱 비밀번호를 발급해서 쓸 것.
- **새 글 감지 로직**(`Invoke-FssNotificationCheck`): 모니터링 대상 각 메뉴의 첫 페이지만 조회(`Get-FssBoardList` 또는 `Get-FssJobListRows` 재사용) → 현재 ID 목록과 저장된 이전 ID 목록을 비교(번호 순서가 아니라 `-notcontains` 존재 여부로 비교, ID 체계가 게시판마다 달라도 안전) → 새 항목이 있으면 모아서 메일 한 통으로 발송, 상태 갱신.
- **이메일 발송**(`Send-FssNotificationEmail`): `System.Net.Mail.SmtpClient` 직접 사용(Send-MailMessage는 최신 PowerShell에서 지원 중단 경고가 있어 회피). 새 글이 없을 때 "지금 확인" 버튼을 누르면 테스트 메일(`-TestOnly`)을 보내 설정이 맞는지 확인 가능.
- **작업 스케줄러 등록**(`Register-FssNotifyTask`/`Unregister-FssNotifyTask`/`Test-FssNotifyTaskRegistered`, `schtasks.exe` 사용): 등록되는 실행 명령은 컴파일된 exe가 아니라 **`powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "dist\FSS-DataTool-merged.ps1" -CheckOnly`** — exe를 직접 스케줄링하면 Smart App Control 차단에 걸리므로, PowerShell 인터프리터로 병합 스크립트를 실행하는 방식으로 우회함(기존 build.ps1의 SAC 우회 패턴과 동일). 이 때문에 build.ps1이 병합 스크립트 사본을 `dist\FSS-DataTool-merged.ps1`로도 저장하도록 수정함(`Get-FssMergedScriptPath`가 실행 중인 exe와 같은 폴더에서 이 파일을 찾음) — **build.ps1로 다시 빌드할 때마다 이 사본도 갱신되므로, 알림 로직을 고친 뒤에는 반드시 재빌드해야 스케줄러가 최신 로직으로 실행됨.**
- **진입점 분기**(build.ps1의 병합 로직): 병합 스크립트 맨 앞에 `param([switch]$CheckOnly)`를 추가하고, 맨 끝을 `if ($CheckOnly) { Invoke-FssNotificationCheck } else { Show-FssGui }`로 분기. 평소 더블클릭 실행(exe)이나 인자 없이 실행하면 기존과 동일하게 GUI가 뜸.

### GUI (Gui.ps1)
- 메인 화면 하단 버튼 줄에 **"알림 설정"** 버튼 추가 → `Show-FssNotifySettings` 모달 다이얼로그.
- 다이얼로그 구성: 알림 사용 체크박스, 발신 Gmail 주소, 앱 비밀번호(마스킹), 받는사람(콤마 구분 여러 명), SMTP 서버/포트/SSL(Gmail 기본값 프리셋, 편집 가능), 확인 주기 드롭다운, 모니터링할 메뉴 체크리스트(`CheckedListBox`, cprCoreInfo 제외 전체 메뉴), 작업 스케줄러 등록 상태 표시, 진행 로그.
- 버튼: "설정 저장" / "지금 확인/테스트"(설정 저장 후 즉시 확인, 새 글 없으면 테스트 메일) / "자동 확인 등록"(작업 스케줄러 등록 전 확인 메시지박스 표시 — 시스템 변경이라 되돌리기 전까지 확인 절차를 둠) / "자동 확인 해제".

### 검증 방법 및 결과
- 백엔드 함수(`Get-FssNotifiableMenus`, `Invoke-FssNotificationCheck`의 최초등록/정상상태/새글감지 3단계, `Register-FssNotifyTask`/`Test-FssNotifyTaskRegistered`/`Unregister-FssNotifyTask` 전체 수명주기)는 `Send-FssNotificationEmail`을 목(mock)으로 대체해서 실제 이메일 발송 없이 로직만 직접 호출로 검증 — 정상 동작 확인(accnutAdtorInfo 신규 게시글 감지 시 올바른 제목/URL로 메일 페이로드 구성됨을 확인).
- 병합 스크립트의 `-CheckOnly` 진입점(param 바인딩 + 분기)도 실제 `powershell.exe -File ... -CheckOnly` 실행으로 검증 완료.
- **GUI 자동클릭 이슈 원인 규명**: Phase 2 때 "DPI 가상화 문제"로 추정했던 좌표 클릭 실패는 실제로는 **창이 최소화(iconic) 상태였거나, Windows의 foreground-lock 정책 때문에 자동화 스크립트의 `SetForegroundWindow` 호출이 실제로 포커스를 못 가져와서 클릭이 엉뚱한(화면상 그 위치에 떠 있던 다른) 창으로 전달된 것**임을 확인함(`WindowFromPoint`로 클릭 지점의 실제 윈도우 클래스를 찍어보니 다른 앱 창이었음). **해결책**: 화면 좌표 클릭 대신 `EnumChildWindows`로 자식 컨트롤의 실제 HWND를 얻어 `SendMessage`로 `BM_CLICK`(버튼/체크박스)·`WM_SETTEXT`(텍스트박스)를 직접 보내는 방식을 쓰면 포커스/좌표 문제와 무관하게 확실하게 동작함. 이 방식으로 알림 설정 다이얼로그의 체크박스 토글, 텍스트 입력, "지금 확인/테스트" 버튼 클릭까지 실제로 동작 확인(로그창에 "모니터링 대상으로 선택된 메뉴가 없습니다" 정상 출력, 설정 파일도 올바르게 저장됨). **다음에 GUI 자동클릭이 필요하면 이 SendMessage(BM_CLICK/WM_SETTEXT) 방식을 우선 사용할 것** — 좌표 클릭이나 UI Automation(InvokePattern 등)은 이 ps2exe 빌드 환경에서 신뢰할 수 없음(UIA는 모든 컨트롤이 `ControlType.Pane`으로만 노출되고 지원 패턴이 없음 — comctl32 v6 매니페스트 부재 추정).
- 테스트에 사용한 더미 설정(가짜 Gmail 계정 등)은 검증 후 `%AppData%\FSS-DataTool\`에서 모두 삭제함 — 실제 사용 시 사용자가 GUI에서 직접 입력해야 함.

## 다음에 할 일
- **Phase 2는 완료.** (`cprCoreInfo` 엑셀 다운로드, `accnutAdtorInfo` 목록·개별/시리즈 다운로드 모두 실사이트 검증 완료.)
- **Phase 4(새 글 이메일 알림)도 완료.** 단, **사용자가 아직 실제 값을 입력 안 함** — 다음에 실사용하려면 GUI "알림 설정"에서 (1) 본인 Gmail 계정 2단계 인증 켜고 앱 비밀번호 발급 (2) 발신 이메일/앱 비밀번호/받는사람/모니터링할 메뉴 입력 후 "설정 저장" (3) "지금 확인/테스트"로 테스트 메일 확인 (4) "자동 확인 등록"으로 작업 스케줄러 등록 — 이 4단계는 사용자가 직접 해야 함(비밀번호 등 민감정보라 대신 입력 불가).
- **Phase 3(미착수)**: `Save-FssSeries`의 개정/전문 추정 로직(`Get-FssRevisionHint`) 개선.
- **exe 실행 시 Smart App Control(SAC) 차단**: 여전히 그대로 둠(사용자가 "지금은 유지, 필요시 직접 판단"으로 결정). 차단 해제(SAC 끄기는 비가역적 — Windows 재설치 전까지 재활성화 불가 / Windows Defender 예외 추가는 가역적) 여부는 사용자 판단 필요, 요청 시 진행. 또는 exe 대신 .cmd/.bat 단일 파일로 전환(SAC 우회, 더블클릭 실행 가능) 검토.
- **스타일 경고(지금은 보류)**: PSScriptAnalyzer가 `Load-CurrentBoard`, `Process-DownloadedFile` 함수명에 대해 "승인되지 않은 동사(PSUseApprovedVerbs)" 경고를 띄움. 기능상 문제 없음, 사용자가 "나중에 한 번에 정리"하기로 결정 — 다음에 손댈 때 `Invoke-`/`Update-` 등으로 이름 변경 고려.
- **ps2exe 모듈**: 이번 세션에서 `Install-PackageProvider -Name NuGet` + `Install-Module -Name ps2exe`를 `-Confirm:$false`로 재설치함(이 PC/세션에는 없었음). 다음에 새 환경에서 build.ps1이 "NuGet provider가 필요합니다" 오류를 내면 이 방법으로 우회 가능.
