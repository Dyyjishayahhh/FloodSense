$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath '.').Path
$sourceDocx = Join-Path $root '1. FLOODSENSE-MAIN\MAIN FILESS\FloodSense_Chapters_1_to_3_Chapter_2_Enhanced_2026-09-21.docx'
$finalDocx = Join-Path $root '1. FLOODSENSE-MAIN\MAIN FILESS\FloodSense_Chapters_1_to_3_Chapter_2_Applications_and_Narrative_Citations_2026-09-21.docx'
$outDir = Join-Path $root 'tmp\ch2_citation_revision_2026-09-21\word-render'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    foreach ($item in @(
        @{ Input = $sourceDocx; Name = 'baseline' },
        @{ Input = $finalDocx; Name = 'final' }
    )) {
        $tempDocx = Join-Path $env:TEMP ("floodsense-ch2-" + $item.Name + '.docx')
        $tempPdf = Join-Path $env:TEMP ("floodsense-ch2-" + $item.Name + '.pdf')
        Copy-Item -LiteralPath $item.Input -Destination $tempDocx -Force
        Remove-Item -LiteralPath $tempPdf -Force -ErrorAction SilentlyContinue
        $document = $word.Documents.Open($tempDocx, $false, $true)
        try {
            $document.Repaginate()
            $document.Fields.Update() | Out-Null
            foreach ($section in $document.Sections) {
                foreach ($header in $section.Headers) {
                    $header.Range.Fields.Update() | Out-Null
                }
                foreach ($footer in $section.Footers) {
                    $footer.Range.Fields.Update() | Out-Null
                }
            }
            $document.ExportAsFixedFormat($tempPdf, 17)
            Copy-Item -LiteralPath $tempPdf -Destination (Join-Path $outDir ($item.Name + '.pdf')) -Force
        }
        finally {
            $document.Close($false)
        }
    }
}
finally {
    $word.Quit()
}
Write-Output $outDir
