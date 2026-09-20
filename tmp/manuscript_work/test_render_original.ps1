$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath '1. FLOODSENSE-MAIN\MAIN FILESS\FloodSense_Chapter_1_REvision_Version_1.docx').Path
$docx = Join-Path $env:TEMP 'floodsense-original.docx'
$pdf = Join-Path $env:TEMP 'floodsense-original.pdf'
Copy-Item -LiteralPath $source -Destination $docx -Force
Remove-Item -LiteralPath $pdf -Force -ErrorAction SilentlyContinue
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    Write-Output 'OPEN'
    $document = $word.Documents.Open($docx, $false, $true)
    try {
        Write-Output 'EXPORT'
        $document.ExportAsFixedFormat($pdf, 17)
        Write-Output ((Get-Item -LiteralPath $pdf).Length)
    }
    finally { $document.Close($false) }
}
finally { $word.Quit() }
