function Test-HwpInstalled {
    try {
        $regPath = "HKLM:\SOFTWARE\Classes\HWPFrame.HwpObject"
        if (Test-Path $regPath) { return $true }
        $regPath32 = "HKLM:\SOFTWARE\WOW6432Node\Classes\HWPFrame.HwpObject"
        if (Test-Path $regPath32) { return $true }
    } catch {}
    return $false
}

function Get-HwpPlainText {
    # HWP 본문을 일반 텍스트로 추출(AI 요약 입력용). 한컴오피스 COM의 GetTextFile("TEXT","") 사용.
    # 미설치거나 추출 실패 시 $null 반환(호출부에서 안내 메시지 처리).
    param([Parameter(Mandatory)] [string]$HwpPath)

    if (-not (Test-HwpInstalled)) { return $null }

    $hwp = $null
    try {
        $hwp = New-Object -ComObject HWPFrame.HwpObject
        try {
            $hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
        } catch {}

        $hwp.Open($HwpPath, "HWP", "")
        $text = $hwp.GetTextFile("TEXT", "")
        $hwp.Clear(1)
        $hwp.Quit()
        return $text
    } catch {
        try { if ($hwp) { $hwp.Quit() } } catch {}
        return $null
    } finally {
        if ($hwp) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($hwp) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
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
