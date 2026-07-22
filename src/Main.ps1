$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptDir) { $scriptDir = "." }

. (Join-Path $scriptDir "Config.ps1")
. (Join-Path $scriptDir "Scraper.ps1")
. (Join-Path $scriptDir "Notify.ps1")
. (Join-Path $scriptDir "HwpConvert.ps1")
. (Join-Path $scriptDir "Gui.ps1")

Show-FssGui
