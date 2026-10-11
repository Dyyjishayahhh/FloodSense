$ErrorActionPreference = 'Stop'
$docxPath = (Resolve-Path -LiteralPath '1. FLOODSENSE-MAIN\WORKING_REVISION_2026-10-11\FloodSense_Chapters_1_to_3_Survey_Ongoing_Working_Revision_2026-10-11.docx').Path
$outputDirectory = Join-Path (Resolve-Path -LiteralPath 'tmp\current_revision_work').Path 'render_final_word'
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}
$pdfPath = Join-Path $outputDirectory 'FloodSense_Chapters_1_to_3_Survey_Ongoing_Working_Revision_2026-10-11.pdf'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $document = $word.Documents.Open($docxPath, $false, $true)
    try {
        $document.ExportAsFixedFormat($pdfPath, 17)
    }
    finally {
        $document.Close($false)
    }
}
finally {
    $word.Quit()
}
Write-Output $pdfPath
