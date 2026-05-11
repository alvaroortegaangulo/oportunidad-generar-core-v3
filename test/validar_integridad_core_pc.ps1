param(
    [string]$CorePath = ".\.local\test-output\CORE_PC_20250256445_test.xlsx",
    [string]$TemplatePath = ".\ficheros-auxiliares\PMBox_plantilla CORE_PC_v1_04 (2).xlsx",
    [string]$BaselinePath = ".\test\baseline_core_pc.csv",
    [ValidateSet("Completo", "DatosIncompletos")]
    [string]$Scenario = "Completo",
    [switch]$SkipBaseline
)

$ErrorActionPreference = "Stop"

$requiredSheets = @(
    "Project Infor",
    "Cost Overview",
    "Resources",
    "Cost Planning",
    "Monthly View",
    "Cost Summary"
)

$forbiddenSheets = @(
    "Trazabilidad_RPA"
)

$xlCellTypeFormulas = -4123
$xlErrors = 16

function Resolve-RequiredFile {
    param(
        [string]$Path,
        [string]$Label,
        [string]$MissingHint
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$Label no existe: $Path. $MissingHint"
    }

    return (Resolve-Path -LiteralPath $Path).Path
}

function Get-WorksheetNames {
    param([object]$Workbook)

    $names = New-Object System.Collections.Generic.List[string]

    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $names.Add([string]$Workbook.Worksheets.Item($index).Name)
    }

    return $names
}

function Test-WorksheetExists {
    param(
        [object]$Workbook,
        [string]$Name
    )

    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        if ([string]::Equals([string]$Workbook.Worksheets.Item($index).Name, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Get-WorksheetByName {
    param(
        [object]$Workbook,
        [string]$Name
    )

    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $worksheet = $Workbook.Worksheets.Item($index)

        if ([string]::Equals([string]$worksheet.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $worksheet
        }
    }

    throw "No existe la hoja esperada: $Name"
}

function Get-UsedDimensions {
    param([object]$Worksheet)

    $usedRange = $Worksheet.UsedRange

    return [pscustomobject]@{
        Rows = [int]$usedRange.Rows.Count
        Columns = [int]$usedRange.Columns.Count
    }
}

function Get-FormulaCount {
    param([object]$Worksheet)

    try {
        return [int]$Worksheet.UsedRange.SpecialCells($xlCellTypeFormulas).Count
    }
    catch {
        return 0
    }
}

function Get-FormulaErrorAddress {
    param([object]$Worksheet)

    try {
        $errorCells = $Worksheet.UsedRange.SpecialCells($xlCellTypeFormulas, $xlErrors)
        return [string]$errorCells.Address()
    }
    catch {
        return ""
    }
}

function Get-NotaShapeCount {
    param([object]$Worksheet)

    $count = 0

    for ($index = 1; $index -le $Worksheet.Shapes.Count; $index++) {
        $shape = $Worksheet.Shapes.Item($index)
        $text = ""

        try {
            if ($shape.TextFrame2.HasText) {
                $text = [string]$shape.TextFrame2.TextRange.Text
            }
        }
        catch {
            $text = ""
        }

        if ($text -match "NOTA|NOTE") {
            $count++
        }
    }

    return $count
}

function Get-WorksheetCellText {
    param(
        [object]$Workbook,
        [string]$Sheet,
        [string]$Cell
    )

    $worksheet = Get-WorksheetByName -Workbook $Workbook -Name $Sheet
    $value = $worksheet.Range($Cell).Value2

    if ($null -eq $value) {
        return ""
    }

    return ([string]$value).Trim()
}

function Test-WorksheetCellRed {
    param(
        [object]$Workbook,
        [string]$Sheet,
        [string]$Cell
    )

    $worksheet = Get-WorksheetByName -Workbook $Workbook -Name $Sheet
    $range = $worksheet.Range($Cell)

    try {
        $fontColor = [int]$range.Font.Color
        $fontColorIndex = [int]$range.Font.ColorIndex

        return ($fontColor -eq 255 -or $fontColorIndex -eq 3)
    }
    catch {
        return $false
    }
}

function Assert-CellTextRed {
    param(
        [object]$Workbook,
        [string]$Sheet,
        [string]$Cell,
        [string]$Expected,
        [System.Collections.Generic.List[string]]$Failures
    )

    $actual = Get-WorksheetCellText -Workbook $Workbook -Sheet $Sheet -Cell $Cell

    if ($actual -ne $Expected) {
        $Failures.Add("$Sheet!$Cell esperado='$Expected' actual='$actual'")
    }

    if (-not (Test-WorksheetCellRed -Workbook $Workbook -Sheet $Sheet -Cell $Cell)) {
        $Failures.Add("$Sheet!$Cell debe estar en fuente roja para el escenario de datos incompletos.")
    }
}

function Test-ScenarioSpecificCells {
    param(
        [object]$Workbook,
        [string]$ScenarioName,
        [System.Collections.Generic.List[string]]$Failures
    )

    if ($ScenarioName -ne "DatosIncompletos") {
        return
    }

    $expectedRedCells = @(
        @{ Sheet = "Project Infor"; Cell = "C10"; Text = "WBS no disponible" },
        @{ Sheet = "Project Infor"; Cell = "C16"; Text = "Cliente SAP no disponible" },
        @{ Sheet = "Project Infor"; Cell = "C19"; Text = "Sales Order no disponible" },
        @{ Sheet = "Project Infor"; Cell = "C22"; Text = "PM cod no disponible" },
        @{ Sheet = "Project Infor"; Cell = "D22"; Text = "PM no disponible" },
        @{ Sheet = "Project Infor"; Cell = "C23"; Text = "PL cod no disponible" },
        @{ Sheet = "Project Infor"; Cell = "D23"; Text = "Practice leader no disponible" },
        @{ Sheet = "Project Infor"; Cell = "C24"; Text = "Sales manager cod no disponible" },
        @{ Sheet = "Resources"; Cell = "B7"; Text = "Codigo empleado no disponible" },
        @{ Sheet = "Resources"; Cell = "D7"; Text = "Intercompany no determinable" },
        @{ Sheet = "Cost Planning"; Cell = "B64"; Text = "Proveedor no disponible" },
        @{ Sheet = "Cost Planning"; Cell = "D64"; Text = "PO compra no disponible" },
        @{ Sheet = "Cost Planning"; Cell = "E64"; Text = "Ariba no disponible" }
    )

    foreach ($case in $expectedRedCells) {
        Assert-CellTextRed -Workbook $Workbook -Sheet $case.Sheet -Cell $case.Cell -Expected $case.Text -Failures $Failures
    }
}

$coreResolved = Resolve-RequiredFile -Path $CorePath -Label "CORE generado" -MissingHint "Ejecuta primero test\test_generar_core_pc.xaml desde UiPath Studio 23.10.4."
$templateResolved = Resolve-RequiredFile -Path $TemplatePath -Label "Plantilla CORE PC" -MissingHint "Revisa la ruta de la plantilla oficial."
$baselineResolved = Resolve-RequiredFile -Path $BaselinePath -Label "Baseline CORE PC" -MissingHint "Revisa test\baseline_core_pc.csv."

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$failures = New-Object System.Collections.Generic.List[string]
$templateWorkbook = $null
$coreWorkbook = $null

try {
    $templateWorkbook = $excel.Workbooks.Open($templateResolved, 0, $true)

    $coreWorkbook = $excel.Workbooks.Open($coreResolved, 0, $false)
    try {
        $coreWorkbook.ForceFullCalculation = $true
    }
    catch {
    }

    $excel.CalculateFullRebuild()
    $coreWorkbook.Save()
    $coreWorkbook.Close($true) | Out-Null
    $coreWorkbook = $null

    $coreWorkbook = $excel.Workbooks.Open($coreResolved, 0, $true)

    $templateSheets = Get-WorksheetNames -Workbook $templateWorkbook
    $coreSheets = Get-WorksheetNames -Workbook $coreWorkbook

    foreach ($sheet in $requiredSheets) {
        if (-not (Test-WorksheetExists -Workbook $coreWorkbook -Name $sheet)) {
            $failures.Add("Falta hoja obligatoria en CORE generado: $sheet")
        }
    }

    foreach ($sheet in $forbiddenSheets) {
        if (Test-WorksheetExists -Workbook $coreWorkbook -Name $sheet) {
            $failures.Add("El CORE generado contiene una hoja no permitida: $sheet")
        }
    }

    if ($coreSheets.Count -ne $templateSheets.Count) {
        $failures.Add("Numero de hojas distinto a plantilla. Plantilla=$($templateSheets.Count); CORE=$($coreSheets.Count)")
    }

    foreach ($sheet in $templateSheets) {
        if (-not (Test-WorksheetExists -Workbook $coreWorkbook -Name $sheet)) {
            $failures.Add("El CORE generado no conserva la hoja de plantilla: $sheet")
            continue
        }

        $templateSheet = Get-WorksheetByName -Workbook $templateWorkbook -Name $sheet
        $coreSheet = Get-WorksheetByName -Workbook $coreWorkbook -Name $sheet
        $templateDimensions = Get-UsedDimensions -Worksheet $templateSheet
        $coreDimensions = Get-UsedDimensions -Worksheet $coreSheet

        if ($coreDimensions.Rows -ne $templateDimensions.Rows -or $coreDimensions.Columns -ne $templateDimensions.Columns) {
            $failures.Add("$sheet dimensiones distintas. Plantilla=$($templateDimensions.Rows)x$($templateDimensions.Columns); CORE=$($coreDimensions.Rows)x$($coreDimensions.Columns)")
        }

        $templateFormulaCount = Get-FormulaCount -Worksheet $templateSheet
        $coreFormulaCount = Get-FormulaCount -Worksheet $coreSheet

        if ($coreFormulaCount -ne $templateFormulaCount) {
            $failures.Add("$sheet formulas distintas. Plantilla=$templateFormulaCount; CORE=$coreFormulaCount")
        }

        $formulaErrors = Get-FormulaErrorAddress -Worksheet $coreSheet

        if (-not [string]::IsNullOrWhiteSpace($formulaErrors)) {
            $failures.Add("$sheet contiene formulas con error en: $formulaErrors")
        }

        $templateNotas = Get-NotaShapeCount -Worksheet $templateSheet
        $coreNotas = Get-NotaShapeCount -Worksheet $coreSheet

        if ($coreNotas -ne $templateNotas) {
            $failures.Add("$sheet cuadros NOTA distintos. Plantilla=$templateNotas; CORE=$coreNotas")
        }
    }

    Test-ScenarioSpecificCells -Workbook $coreWorkbook -ScenarioName $Scenario -Failures $failures
}
finally {
    if ($coreWorkbook) {
        $coreWorkbook.Close($false) | Out-Null
    }

    if ($templateWorkbook) {
        $templateWorkbook.Close($false) | Out-Null
    }

    $excel.Quit() | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Integridad CORE PC no superada. Fallos: $($failures.Count)"
}

if (-not $SkipBaseline) {
    & (Join-Path $PSScriptRoot "validar_baseline_core_pc.ps1") -CorePath $coreResolved -BaselinePath $baselineResolved
}

$scopeText = if ($SkipBaseline) {
    "hojas, dimensiones, notas, formulas, recalculo y escenario $Scenario"
}
else {
    "hojas, dimensiones, notas, formulas, recalculo, baseline y escenario $Scenario"
}

Write-Host "Integridad CORE PC superada: $scopeText validados en $coreResolved"
