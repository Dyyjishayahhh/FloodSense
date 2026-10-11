$ErrorActionPreference = 'Stop'
$pdfPath = Join-Path (Resolve-Path -LiteralPath 'tmp\current_revision_work\render_final_word').Path 'word_pdf_smoke_test.pdf'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $document = $word.Documents.Add()
    $document.Content.Text = 'Word PDF smoke test'
    $document.ExportAsFixedFormat($pdfPath, 17)
    $document.Close($false)
}
finally { $word.Quit() }
Write-Output $pdfPath
