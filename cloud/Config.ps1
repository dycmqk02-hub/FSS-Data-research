# Invoke-WebRequest의 기본 진행률 표시(별도 팝업/창으로 뜸)를 끔 - 전 페이지를 순회하며 여러 번 호출할 때 방해됨
$ProgressPreference = 'SilentlyContinue'

# FSS 업무자료(공시/회계) 메뉴 매핑 테이블
# Type: "bbs"  -> /fss/bbs/{BbsId}/list.do?menuNo={MenuNo}   (게시판형, Phase 1)
#       "job"  -> /fss/job/{JobPath}/list.do?menuNo={MenuNo} (검색폼형, Phase 2)
#       "info" -> 목록/첨부 없는 정보성 페이지 (스크래핑 대상 아님, 안내만 표시)
# Type="job"일 때 SubType:
#       "table" -> cprCoreInfo형. 결과가 게시글이 아니라 가로로 긴 데이터 표(연번~결산년월 27개 컬럼) + Excel 전체 다운로드(excelDown.do)만 제공. 개별 첨부파일 없음.
#       "list"  -> accnutAdtorInfo형. 게시판형과 비슷하게 행 단위 목록 + 행마다 첨부파일(hwp/hwpx) 링크가 목록 페이지에 바로 노출됨(상세 페이지 별도 조회 불필요).

$script:FssMenuMap = @(
    # ---- 공시 ----
    [PSCustomObject]@{ Category = "공시"; Group = "공시"; Name = "기업공시제도일반"; Type = "bbs"; BbsId = "B0000145"; MenuNo = "200152" }
    [PSCustomObject]@{ Category = "공시"; Group = "공시"; Name = "공시유의사항";     Type = "bbs"; BbsId = "B0000146"; MenuNo = "200153" }
    [PSCustomObject]@{ Category = "공시"; Group = "공시"; Name = "FAQ";             Type = "bbs"; BbsId = "B0000147"; MenuNo = "200154" }

    # ---- 회계 : 회계법인 정보 통합조회 (job-app, Phase 2) ----
    [PSCustomObject]@{ Category = "회계"; Group = "회계법인 정보"; Name = "회계법인 정보 통합조회"; Type = "job"; SubType = "table"; JobPath = "cprCoreInfo"; MenuNo = "200429" }

    # ---- 회계 : 회계기준 ----
    [PSCustomObject]@{ Category = "회계"; Group = "회계기준"; Name = "기업회계기준해석";     Type = "bbs"; BbsId = "B0000125"; MenuNo = "200432" }
    [PSCustomObject]@{ Category = "회계"; Group = "회계기준"; Name = "재무보고 실무의견서"; Type = "bbs"; BbsId = "B0000126"; MenuNo = "200433" }
    [PSCustomObject]@{ Category = "회계"; Group = "회계기준"; Name = "감독지침";           Type = "bbs"; BbsId = "B0000127"; MenuNo = "200434" }
    [PSCustomObject]@{ Category = "회계"; Group = "회계기준"; Name = "주석공시 모범사례"; Type = "bbs"; BbsId = "B0000128"; MenuNo = "200435" }

    # ---- 회계 : 감사기준 ----
    [PSCustomObject]@{ Category = "회계"; Group = "감사기준"; Name = "회계감사기준";           Type = "bbs"; BbsId = "B0000129"; MenuNo = "200437" }
    [PSCustomObject]@{ Category = "회계"; Group = "감사기준"; Name = "회계감사 실무의견서";    Type = "bbs"; BbsId = "B0000130"; MenuNo = "200438" }
    [PSCustomObject]@{ Category = "회계"; Group = "감사기준"; Name = "분ㆍ반기재무제표검토준칙"; Type = "bbs"; BbsId = "B0000131"; MenuNo = "200439" }

    # ---- 회계 : 회계질의 ----
    [PSCustomObject]@{ Category = "회계"; Group = "회계질의"; Name = "K-IFRS 질의회신요약";           Type = "bbs"; BbsId = "B0000132"; MenuNo = "200442" }
    [PSCustomObject]@{ Category = "회계"; Group = "회계질의"; Name = "일반기업회계기준 질의회신요약"; Type = "bbs"; BbsId = "B0000133"; MenuNo = "200443" }
    [PSCustomObject]@{ Category = "회계"; Group = "회계질의"; Name = "과거 국제회계기준 Q&A";         Type = "bbs"; BbsId = "B0000134"; MenuNo = "200445" }

    # ---- 회계 : 회계감리 ----
    [PSCustomObject]@{ Category = "회계"; Group = "회계감리"; Name = "심사ㆍ감리지적사례";       Type = "bbs"; BbsId = "B0000135"; MenuNo = "200448" }
    [PSCustomObject]@{ Category = "회계"; Group = "회계감리"; Name = "회계법인품질관리매뉴얼"; Type = "bbs"; BbsId = "B0000136"; MenuNo = "200449" }
    [PSCustomObject]@{ Category = "회계"; Group = "회계감리"; Name = "회계법인사업보고서";     Type = "bbs"; BbsId = "B0000137"; MenuNo = "200450" }
    [PSCustomObject]@{ Category = "회계"; Group = "회계감리"; Name = "감사인 감리결과 개선권고사항 등"; Type = "bbs"; BbsId = "B0000291"; MenuNo = "200620"; ExtraParams = "cl1Cd=01" }
    [PSCustomObject]@{ Category = "회계"; Group = "회계감리"; Name = "회계감리결과제재 등"; Type = "job"; SubType = "list"; JobPath = "accnutAdtorInfo"; MenuNo = "200617"; ExtraParams = "acntnWrkCode=02" }

    # ---- 회계 : 감사인 등록 ----
    [PSCustomObject]@{ Category = "회계"; Group = "감사인 등록"; Name = "주권상장법인 감사인 등록현황"; Type = "bbs"; BbsId = "B0000142"; MenuNo = "200462" }
    [PSCustomObject]@{ Category = "회계"; Group = "감사인 등록"; Name = "등록신청서 서식 및 작성 매뉴얼"; Type = "bbs"; BbsId = "B0000143"; MenuNo = "200464" }
    [PSCustomObject]@{ Category = "회계"; Group = "감사인 등록"; Name = "주권상장법인 감사인 등록신청 FAQ"; Type = "bbs"; BbsId = "B0000144"; MenuNo = "200465" }

    # ---- 회계 : 자료실 ----
    [PSCustomObject]@{ Category = "회계"; Group = "자료실"; Name = "회계감독 동향자료"; Type = "bbs"; BbsId = "B0000154"; MenuNo = "200467" }
    [PSCustomObject]@{ Category = "회계"; Group = "자료실"; Name = "연구용역 보고서";   Type = "bbs"; BbsId = "B0000156"; MenuNo = "200469" }

    # ---- 회계 : 정보성 페이지 (다운로드 대상 없음, Phase 1 제외) ----
    [PSCustomObject]@{ Category = "회계"; Group = "회계감독권역"; Name = "회계감독권역 소개"; Type = "info"; MenuNo = "200427" }
    [PSCustomObject]@{ Category = "회계"; Group = "회계질의";     Name = "질의회신 업무절차"; Type = "info"; MenuNo = "200441" }
    [PSCustomObject]@{ Category = "회계"; Group = "회계감리";     Name = "회계감리업무절차"; Type = "info"; MenuNo = "200446" }
    [PSCustomObject]@{ Category = "회계"; Group = "외부감사인 지정"; Name = "외부감사인 지정 절차";     Type = "info"; MenuNo = "200457" }
    [PSCustomObject]@{ Category = "회계"; Group = "외부감사인 지정"; Name = "외부감사인 지정신청 안내"; Type = "info"; MenuNo = "200458" }
)

# 실행 시 자동으로 조회할 기본 항목
$script:FssDefaultItem = $script:FssMenuMap | Where-Object { $_.BbsId -eq "B0000129" } | Select-Object -First 1
