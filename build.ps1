$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcDir = Join-Path $root "src"
$distDir = Join-Path $root "dist"
if (-not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir -Force | Out-Null }

if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Output "ps2exe 모듈 설치 중..."
    Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber
}
Import-Module ps2exe -Force

# 여러 소스 파일을 하나의 스크립트로 병합 (exe 단일 파일화)
# 맨 앞에 param() 블록을 둬서 "-CheckOnly" 인자로 GUI 없이 알림 확인만 실행 가능하게 함(작업 스케줄러용).
$mergedPath = Join-Path $env:TEMP "FSS-DataTool-merged.ps1"
$paramBlock = "param([switch]`$CheckOnly)`n"
$parts = @("Config.ps1", "Scraper.ps1", "Notify.ps1", "HwpConvert.ps1", "Gui.ps1") | ForEach-Object {
    Get-Content -Path (Join-Path $srcDir $_) -Raw
}
$dispatch = "`nif (`$CheckOnly) { Invoke-FssNotificationCheck } else { Show-FssGui }`n"
($paramBlock + ($parts -join "`n`n") + $dispatch) | Out-File -FilePath $mergedPath -Encoding utf8

# dist\ 안에도 사본을 둬서 Windows 작업 스케줄러가 참조할 안정적인 경로를 제공(exe는 Smart App Control에 막히므로
# 스케줄러는 이 .ps1을 powershell.exe -File로 직접 실행하는 방식을 씀 - Notify.ps1의 Register-FssNotifyTask 참고).
$mergedDistPath = Join-Path $distDir "FSS-DataTool-merged.ps1"
Copy-Item -Path $mergedPath -Destination $mergedDistPath -Force

$exePath = Join-Path $distDir "FSS-DataTool.exe"

Invoke-ps2exe -inputFile $mergedPath -outputFile $exePath `
    -title "FSS 업무자료 조회·다운로드" `
    -company "" -product "FSS-DataTool" `
    -noConsole -STA

Write-Output "빌드 완료: $exePath"
