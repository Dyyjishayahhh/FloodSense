$ErrorActionPreference = 'Stop'
$docxPath = (Resolve-Path -LiteralPath '1. FLOODSENSE-MAIN\WORKING_REVISION_2026-10-11\FloodSense_Chapters_1_to_3_Survey_Ongoing_Working_Revision_2026-10-11.docx').Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $document = $word.Documents.Open($docxPath, $false, $true)
    try {
        Write-Output ('pages=' + $document.ComputeStatistics(2))
        for ($tableIndex = 1; $tableIndex -le $document.Tables.Count; $tableIndex++) {
            $table = $document.Tables.Item($tableIndex)
            $firstPage = $table.Range.Information(3)
            $lastPage = $table.Range.Duplicate.Information(3)
            $rows = @()
            for ($rowIndex = 1; $rowIndex -le $table.Rows.Count; $rowIndex++) {
                $row = $table.Rows.Item($rowIndex)
                $rowStart = $row.Range.Duplicate
                $rowStart.Collapse(1)
                $rowEnd = $row.Range.Duplicate
                $rowEnd.Collapse(0)
                $rows += ("$($rowStart.Information(3))-$($rowEnd.Information(3))")
            }
            Write-Output ("table=${tableIndex}; rows=$($table.Rows.Count); columns=$($table.Columns.Count); rowPages=$($rows -join ',')")
        }
        foreach ($paragraph in $document.Paragraphs) {
            $headingText = $paragraph.Range.Text.Trim()
            if ($headingText -eq 'APPENDICES' -or $headingText -like 'Appendix A.*' -or $headingText -like 'Appendix B.*' -or $headingText -like 'Appendix C.*' -or $headingText -eq 'METHODOLOGY') {
                Write-Output ("heading=$headingText; page=$($paragraph.Range.Information(3))")
            }
        }
    }
    finally { $document.Close($false) }
}
finally { $word.Quit() }
