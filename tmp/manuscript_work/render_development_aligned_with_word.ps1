$ErrorActionPreference = 'Stop'
$docx = (Resolve-Path -LiteralPath '1. FLOODSENSE-MAIN\MAIN FILESS\FloodSense_Chapters_1_to_3_Development_Aligned_2026-09-20.docx').Path
$outDir = Join-Path (Resolve-Path -LiteralPath 'tmp\manuscript_work').Path 'rendered_development_aligned'
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$pdf = Join-Path $outDir 'FloodSense_Chapters_1_to_3_Development_Aligned_2026-09-20.pdf'
$tempDocx = Join-Path $env:TEMP 'floodsense-development-aligned-2026-09-20.docx'
$tempPdf = Join-Path $env:TEMP 'floodsense-development-aligned-2026-09-20.pdf'
Copy-Item -LiteralPath $docx -Destination $tempDocx -Force
Remove-Item -LiteralPath $tempPdf -Force -ErrorAction SilentlyContinue
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $document = $word.Documents.Open($tempDocx, $false, $true)
    try {
        $document.Fields.Update() | Out-Null
        foreach ($section in $document.Sections) {
            foreach ($header in $section.Headers) {
                $header.Range.Fields.Update() | Out-Null
            }
        }
        $document.ExportAsFixedFormat($tempPdf, 17)
        Copy-Item -LiteralPath $tempPdf -Destination $pdf -Force
    }
    finally {
        $document.Close($false)
    }
}
finally {
    $word.Quit()
}
Write-Output $pdf
