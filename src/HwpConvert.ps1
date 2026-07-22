function Test-HwpInstalled {
    try {
        $regPath = "HKLM:\SOFTWARE\Classes\HWPFrame.HwpObject"
        if (Test-Path $regPath) { return $true }
        $regPath32 = "HKLM:\SOFTWARE\WOW6432Node\Classes\HWPFrame.HwpObject"
        if (Test-Path $regPath32) { return $true }
    } catch {}
    return $false
}

function Convert-HwpToPdf {
    param(
        [Parameter(Mandatory)] [string]$HwpPath,
        [Parameter(Mandatory)] [string]$PdfPath
    )

    if (-not (Test-HwpInstalled)) {
        return [PSCustomObject]@{ Success = $false; Message = "한컴오피스(한글) 미설치 - PDF 변환 불가, 원본 HWP 유지" }
    }

    $hwp = $null
    try {
        $hwp = New-Object -ComObject HWPFrame.HwpObject
        try {
            $hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
        } catch {}

        $hwp.Open($HwpPath, "HWP", "")
        $hwp.SaveAs($PdfPath, "PDF", "")
        $hwp.Clear(1)
        $hwp.Quit()
        return [PSCustomObject]@{ Success = $true; Message = "PDF 변환 완료" }
    } catch {
        try { if ($hwp) { $hwp.Quit() } } catch {}
        return [PSCustomObject]@{ Success = $false; Message = "PDF 변환 실패: $($_.Exception.Message)" }
    } finally {
        if ($hwp) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($hwp) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
