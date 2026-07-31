# 첨부파일(HWP) 본문을 Google Gemini API로 요약하는 기능 (온디맨드 - 그리드의 "AI요약" 버튼을 누를 때만 호출됨).
# 설정(API 키/모델)은 Notify.ps1과 같은 위치(%AppData%\FSS-DataTool\)에 평문 JSON으로 저장.
# 주의: API 키는 현재 평문 JSON으로 저장됨(Notify.ps1의 앱 비밀번호와 동일한 수준) - 본인 PC 계정 폴더 안에만 저장됨.

function Get-FssAiConfigPath { Join-Path (Get-FssNotifyDir) "ai-config.json" }
function Get-FssAiUsagePath { Join-Path (Get-FssNotifyDir) "ai-usage.json" }

# 이 값은 구글이 실시간으로 알려주는 실제 남은 할당량이 아니라, 이 도구가 로컬에서 누적 집계한 추정치를
# 사용자가 설정한 한도와 비교하는 것뿐임(Gemini API 응답의 usageMetadata.totalTokenCount를 매 호출마다 누적).
# 무료 등급의 실제 한도는 모델/시점에 따라 바뀌므로 https://ai.google.dev/gemini-api/docs/rate-limits 에서 확인 필요.
function Get-FssAiConfig {
    $path = Get-FssAiConfigPath
    $default = [PSCustomObject]@{
        ApiKey           = ""
        Model            = "gemini-2.0-flash"
        DailyTokenBudget = 1000000
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
    if (-not $loaded.Model) { $loaded.Model = $default.Model }
    if (-not $loaded.DailyTokenBudget -or [int]$loaded.DailyTokenBudget -le 0) { $loaded.DailyTokenBudget = $default.DailyTokenBudget }
    return $loaded
}

function Save-FssAiConfig {
    param([Parameter(Mandatory)] $Config)
    $path = Get-FssAiConfigPath
    ($Config | ConvertTo-Json -Depth 5) | Out-File -FilePath $path -Encoding utf8
}

function Get-FssAiUsage {
    # 오늘 날짜 기준 누적 토큰 사용량. 날짜가 바뀌면 자동으로 0부터 다시 셈(구글 무료 등급 일일 한도가
    # 매일 초기화되는 것과 대략 맞추기 위함 - 정확한 초기화 시각/타임존까지 구글과 일치시키진 않음).
    $path = Get-FssAiUsagePath
    $today = Get-Date -Format "yyyy-MM-dd"
    if (Test-Path $path) {
        try {
            $loaded = Get-Content -Path $path -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($loaded.Date -eq $today) {
                return [PSCustomObject]@{ Date = $today; TokensUsed = [int]$loaded.TokensUsed }
            }
        } catch {}
    }
    return [PSCustomObject]@{ Date = $today; TokensUsed = 0 }
}

function Add-FssAiUsage {
    param([Parameter(Mandatory)] [int]$Tokens)
    $usage = Get-FssAiUsage
    $usage.TokensUsed += $Tokens
    ($usage | ConvertTo-Json) | Out-File -FilePath (Get-FssAiUsagePath) -Encoding utf8
    return $usage
}

function Invoke-FssGeminiSummary {
    # Gemini REST API(generateContent) 호출. 실패 시 예외를 던짐(호출부에서 그리드 셀에 오류 메시지로 표시).
    param(
        [Parameter(Mandatory)] [string]$Text,
        [Parameter(Mandatory)] [string]$ApiKey,
        [string]$Model = "gemini-2.0-flash",
        [int]$MaxInputChars = 12000
    )

    if ($Text.Length -gt $MaxInputChars) { $Text = $Text.Substring(0, $MaxInputChars) }
    $prompt = "다음은 금융감독원 회계 관련 게시물의 첨부문서 본문입니다. 핵심 내용을 한국어로 5줄 이내로 간결하게 요약해줘. 문서 형식이나 서두 인사말은 요약하지 말고 실질적인 내용만 다뤄줘.`n`n$Text"

    $bodyObj = [PSCustomObject]@{
        contents         = @(
            [PSCustomObject]@{ parts = @([PSCustomObject]@{ text = $prompt }) }
        )
        generationConfig = [PSCustomObject]@{ maxOutputTokens = 500 }
    }
    $bodyJson = $bodyObj | ConvertTo-Json -Depth 10
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($bodyJson)

    $uri = "https://generativelanguage.googleapis.com/v1beta/models/$($Model):generateContent?key=$ApiKey"
    $resp = Invoke-RestMethod -Uri $uri -Method Post -ContentType "application/json; charset=utf-8" -Body $bodyBytes -TimeoutSec 60

    $summary = $resp.candidates[0].content.parts[0].text
    if (-not $summary) { throw "API 응답에서 요약 텍스트를 찾을 수 없습니다." }

    $tokensUsed = 0
    if ($resp.usageMetadata -and $resp.usageMetadata.totalTokenCount) { $tokensUsed = [int]$resp.usageMetadata.totalTokenCount }

    return [PSCustomObject]@{
        Summary    = $summary.Trim()
        TokensUsed = $tokensUsed
    }
}
