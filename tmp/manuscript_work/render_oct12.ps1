$ErrorActionPreference = 'Stop'
$docx = (Resolve-Path -LiteralPath '1. FLOODSENSE-MAIN\WORKING_REVISION_2026-10-11\FloodSense_Chapters_1_to_3_Introduction_Chapter2_Working_Revision_2026-10-12.docx').Path
$outDir = Join-Path (Resolve-Path -LiteralPath 'tmp\manuscript_work').Path 'render_oct12'
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$pdf = Join-Path $outDir 'FloodSense_Chapters_1_to_3_Introduction_Chapter2_Working_Revision_2026-10-12.pdf'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $document = $word.Documents.Open($docx, $false, $true)
    try {
        $document.ExportAsFixedFormat($pdf, 17)
    }
    finally {
        $document.Close($false)
    }
}
finally {
    $word.Quit()
}
Write-Output $pdf
