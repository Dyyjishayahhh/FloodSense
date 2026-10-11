$ErrorActionPreference = 'Stop'
$docxPath = (Resolve-Path -LiteralPath '1. FLOODSENSE-MAIN\WORKING_REVISION_2026-10-11\FloodSense_Chapters_1_to_3_Survey_Ongoing_Working_Revision_2026-10-11.docx').Path
$pdfPath = Join-Path (Resolve-Path -LiteralPath 'tmp\current_revision_work\render_final_word').Path 'FloodSense_qa_after_table_fix.pdf'
Write-Output 'START'
$word = New-Object -ComObject Word.Application
$word.Visible = $true
$word.DisplayAlerts = -1
Write-Output 'WORD_READY'
try {
    $document = $word.Documents.Open($docxPath, $false, $true)
    Write-Output ('OPENED pages=' + $document.ComputeStatistics(2))
    try {
        $document.SaveAs2($pdfPath, 17)
        Write-Output 'EXPORTED'
    }
    finally { $document.Close($false) }
}
finally { $word.Quit() }
Write-Output 'DONE'
