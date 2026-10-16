$ErrorActionPreference = 'Stop'

$outputPath = 'E:\School Downloads-Jiraah\00 FILES IN SCHOOL\3RD YEAR 2ND SEM\DCIT 60\FLOODSENSE\1.FLOODSENSE-200A\FloodSense\1. FLOODSENSE-MAIN\WORKING_REVISION_2026-10-11\FloodSense_System_Architecture_Editable_A4_2026-10-14.docx'

function Color-Rgb([int]$red, [int]$green, [int]$blue) {
    return ($red + 256 * $green + 65536 * $blue)
}

$script:word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$script:document = $word.Documents.Add()
$document.PageSetup.PageWidth = 595.3
$document.PageSetup.PageHeight = 841.9
$document.PageSetup.TopMargin = 18
$document.PageSetup.BottomMargin = 18
$document.PageSetup.LeftMargin = 18
$document.PageSetup.RightMargin = 18

function Set-ShapeText($shape, [string]$label, [double]$size = 9, [bool]$bold = $false) {
    $shape.TextFrame.TextRange.Text = $label
    $shape.TextFrame.MarginLeft = 3
    $shape.TextFrame.MarginRight = 3
    $shape.TextFrame.MarginTop = 2
    $shape.TextFrame.MarginBottom = 2
    $shape.TextFrame.VerticalAnchor = 3
    $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 1
    $shape.TextFrame.TextRange.Font.Name = 'Arial'
    $shape.TextFrame.TextRange.Font.Size = $size
    $shape.TextFrame.TextRange.Font.Bold = [int]$bold
    $shape.TextFrame.TextRange.Font.Color = Color-Rgb 26 26 26
}

function Place-Shape($shape) {
    $shape.RelativeHorizontalPosition = 1
    $shape.RelativeVerticalPosition = 1
    $shape.WrapFormat.Type = 3
    $shape.LockAnchor = $true
    return $shape
}

function Add-Box([int]$shapeType, [double]$left, [double]$top, [double]$width, [double]$height,
                 [int]$fill, [int]$outline, [string]$label = '', [double]$fontSize = 9, [bool]$bold = $false) {
    $shape = Place-Shape ($document.Shapes.AddShape($shapeType, $left, $top, $width, $height))
    $shape.Fill.ForeColor.RGB = $fill
    $shape.Line.ForeColor.RGB = $outline
    $shape.Line.Weight = 1
    if ($label) { Set-ShapeText $shape $label $fontSize $bold }
    return $shape
}

function Add-Label([double]$left, [double]$top, [double]$width, [double]$height,
                   [string]$label, [double]$fontSize = 9, [bool]$bold = $false) {
    $shape = Place-Shape ($document.Shapes.AddTextbox(1, $left, $top, $width, $height))
    $shape.Fill.Visible = 0
    $shape.Line.Visible = 0
    Set-ShapeText $shape $label $fontSize $bold
    return $shape
}

function Add-Arrow([double]$x1, [double]$y1, [double]$x2, [double]$y2) {
    $line = Place-Shape ($document.Shapes.AddConnector(1, $x1, $y1, $x2, $y2))
    $line.Line.ForeColor.RGB = Color-Rgb 38 38 38
    $line.Line.Weight = 1.15
    $line.Line.EndArrowheadStyle = 3
    return $line
}

function Add-Rule([double]$x1, [double]$y1, [double]$x2, [double]$y2) {
    $line = Place-Shape ($document.Shapes.AddLine($x1, $y1, $x2, $y2))
    $line.Line.ForeColor.RGB = Color-Rgb 38 38 38
    $line.Line.Weight = 1.15
    return $line
}

try {
    $white = Color-Rgb 255 255 255
    $systemFill = Color-Rgb 255 249 206
    $systemLine = Color-Rgb 240 181 47
    $moduleFill = Color-Rgb 217 217 217
    $moduleLine = Color-Rgb 125 125 125
    $databaseFill = Color-Rgb 169 209 245
    $databaseLine = Color-Rgb 66 118 175
    $black = Color-Rgb 37 37 37

    # Native, editable Word shapes follow the supplied image's pale-yellow
    # system boundary, stacked gray modules, people, cloud, and blue database.
    $null = Add-Box 5 151 91 293 550 $systemFill $systemLine
    try {
        $null = Add-Box 179 249 12 99 63 $white (Color-Rgb 145 145 145) 'INTERNET' 8.5 $false
    } catch {
        $null = Add-Box 9 249 12 99 63 $white (Color-Rgb 145 145 145) 'INTERNET' 8.5 $false
    }
    $null = Add-Arrow 298 76 298 99
    $null = Add-Label 315 74 70 19 'HTTPS' 7.5 $false

    $null = Add-Label 173 103 249 43 "FloodSense`nAndroid Expert System and Web Administration" 10 $true

    $modules = @(
        @{y=158; label="Access and Identity Module`nAuthentication, onboarding, role permissions"},
        @{y=228; label="Scenario and Location Module`nRainfall choice, barangay, map pin, optional foreground GPS"},
        @{y=298; label="Susceptibility Assessment Module`nDeterministic rule matching and assessment"},
        @{y=368; label="Map and Results Module`nGeographic context, susceptibility class, explanation"},
        @{y=438; label="Preparedness and Resources Module`nSeparate DSS guidance and evacuation-center reference"},
        @{y=508; label="Administration and Data Governance Module`nControlled review, provenance, and audit records"}
    )
    foreach ($module in $modules) {
        $null = Add-Box 5 177 $module.y 241 55 $moduleFill $moduleLine $module.label 8.6 $false
    }

    # Resident and administrator are the only human actors. The administrator
    # is placed on the right, as stated in the October 7 TA consultation.
    $null = Add-Box 9 52 281 23 23 $black $black
    $null = Add-Box 5 39 309 49 53 $black $black
    $null = Add-Label 12 367 103 25 'RESIDENT' 8.6 $true
    $null = Add-Box 9 520 281 23 23 $black $black
    $null = Add-Box 5 507 309 49 53 $black $black
    $null = Add-Label 479 367 107 25 'ADMINISTRATOR' 8.1 $true

    $null = Add-Rule 88 335 138 335
    $null = Add-Rule 138 185 138 465
    $null = Add-Arrow 138 185 177 185
    $null = Add-Arrow 138 255 177 255
    $null = Add-Arrow 138 395 177 395
    $null = Add-Arrow 138 465 177 465
    $null = Add-Label 76 151 73 29 "account /`nlogin" 7 $false
    $null = Add-Label 76 221 73 29 "scenario /`nlocation" 7 $false
    $null = Add-Label 76 361 73 29 "map /`nresult" 7 $false
    $null = Add-Label 76 430 73 29 'guidance' 7 $false

    $null = Add-Rule 457 185 457 535
    $null = Add-Rule 457 335 507 335
    $null = Add-Arrow 457 185 418 185
    $null = Add-Arrow 457 535 418 535
    $null = Add-Label 445 151 76 29 "authorized`naccess" 7 $false
    $null = Add-Label 445 497 76 31 "review /`nmanage" 7 $false

    $null = Add-Arrow 284 641 284 667
    $null = Add-Arrow 311 667 311 641
    $null = Add-Label 179 643 88 23 'stored records' 7 $false
    $null = Add-Label 326 643 103 23 'retrieved records' 7 $false
    $null = Add-Box 13 218 669 159 65 $databaseFill $databaseLine "PostgreSQL + PostGIS`nGeographic and System Records" 8.5 $false

    $null = Add-Label 161 769 273 24 'Figure 1. System Architecture of FloodSense.' 9 $false

    $document.SaveAs2($outputPath, 16)
    Write-Output "Created editable architecture: $outputPath"
    Write-Output "Native Word shape count: $($document.Shapes.Count)"
} finally {
    $document.Close($false)
    $word.Quit()
}
