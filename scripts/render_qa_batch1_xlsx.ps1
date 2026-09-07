param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,
    [Parameter(Mandatory = $true)]
    [string]$RenderDirectory
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
New-Item -ItemType Directory -Force -Path $RenderDirectory | Out-Null
$resolvedRenderDirectory = (Resolve-Path -LiteralPath $RenderDirectory).Path

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    # Excel can briefly reject automation calls while its add-ins finish
    # initializing. A short warm-up makes batch rendering deterministic.
    Start-Sleep -Seconds 2
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $workbook = $excel.Workbooks.Open($resolvedWorkbook)
    Start-Sleep -Seconds 1
    $excel.CalculateFullRebuild()
    $workbook.Save()

    for ($sheetIndex = 1; $sheetIndex -le $workbook.Worksheets.Count; $sheetIndex++) {
        $sheet = $workbook.Worksheets.Item($sheetIndex)
        $sheetName = [string]$sheet.Name
        [void]$sheet.Activate()
        $sheet.PageSetup.Zoom = $false
        $sheet.PageSetup.FitToPagesWide = 1
        $sheet.PageSetup.FitToPagesTall = $false
        $pdfPath = Join-Path $resolvedRenderDirectory ($sheetName + '.pdf')
        $sheet.ExportAsFixedFormat(0, $pdfPath)
        Start-Sleep -Milliseconds 700

        $used = $sheet.UsedRange
        if ($null -eq $used) { throw "UsedRange unavailable for sheet $sheetName." }
        $copied = $false
        for ($attempt = 1; $attempt -le 4 -and -not $copied; $attempt++) {
            try {
                [void]$used.CopyPicture(1, 2)
                $copied = $true
            }
            catch {
                if ($attempt -eq 4) { throw }
                Start-Sleep -Milliseconds 750
            }
        }
        Start-Sleep -Milliseconds 400
        $image = [System.Windows.Forms.Clipboard]::GetImage()
        if ($null -eq $image) {
            throw "Excel did not place an image on the clipboard for sheet $sheetName."
        }
        $pngPath = Join-Path $resolvedRenderDirectory ($sheetName + '.png')
        $image.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $image.Dispose()
        [System.Windows.Forms.Clipboard]::Clear()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($used) | Out-Null
        Write-Output "$sheetName|$pdfPath|$pngPath"
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
    }
}
catch {
    Write-Error ("Excel render failed at: " + $_.InvocationInfo.PositionMessage + "`n" + $_.Exception.Message)
    throw
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
