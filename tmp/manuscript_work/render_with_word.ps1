$ErrorActionPreference = 'Stop'
$docx = (Resolve-Path -LiteralPath '1. FLOODSENSE-MAIN\MAIN FILESS\FloodSense_Manuscript_Chapters_1_to_3_Final.docx').Path
$outDir = Join-Path (Resolve-Path -LiteralPath 'tmp\manuscript_work').Path 'rendered_word'
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$pdf = Join-Path $outDir 'FloodSense_Manuscript_Chapters_1_to_3_Final.pdf'
$tempDocx = Join-Path $env:TEMP 'floodsense-final.docx'
$tempPdf = Join-Path $env:TEMP 'floodsense-final.pdf'
Copy-Item -LiteralPath $docx -Destination $tempDocx -Force
Remove-Item -LiteralPath $tempPdf -Force -ErrorAction SilentlyContinue
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    Write-Output "Opening $tempDocx"
    $document = $word.Documents.Open($tempDocx, $false, $true)
    try {
        Write-Output "Rendering $tempPdf"
        $document.ExportAsFixedFormat($tempPdf, 17)
        Copy-Item -LiteralPath $tempPdf -Destination $pdf -Force
        Write-Output "Rendered $pdf"
    }
    finally {
        $document.Close($false)
    }
}
finally {
    $word.Quit()
}
Write-Output $pdf
