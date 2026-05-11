param(
    [string]$IpfPath = ".\.local\test-output\IPF_20250256445_test.xlsx",
    [ValidateSet("Completo", "DatosIncompletos")]
    [string]$Scenario = "Completo",
    [string]$TemplatePath = ".\ficheros-auxiliares\Plantilla IPF.xlsx",
    [string]$ExamplePath = ".\ficheros-auxiliares\IPF - ICT - GAP - ZEV-PCE0012 - Despliegue Smart Airport nueva terminal-  FEB26 v1.xlsx",
    [string]$ReferenceBaselinePath = ".\test\baseline_ipf_referencias.csv",
    [string]$GeneratedBaselinePath = ".\test\baseline_ipf_generado.csv",
    [switch]$SkipGenerated
)

$ErrorActionPreference = "Stop"

$requiredSheets = @(
    "Invoice Request",
    "UC Maint Contract+Inv Request",
    "Billing template instructions",
    "Lists",
    "UC WBS codes"
)

$forbiddenSheets = @(
    "Trazabilidad_RPA",
    "Auditoria_RPA",
    "Audit_RPA"
)

$validationCells = @("C3", "C6", "C8", "F16", "D40", "D41", "C61")
$xlCellTypeFormulas = -4123
$xlCellTypeAllValidation = -4174
$xlErrors = 16
$xlValidateList = 3

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

function Convert-ToCellText {
    param([object]$Value)

    if ($null -eq $Value) {
        return ""
    }

    return ([string]$Value).Trim()
}

function Convert-ToDoubleInvariant {
    param([object]$Value)

    if ($null -eq $Value -or $Value -eq "") {
        return $null
    }

    if ($Value -is [double] -or $Value -is [int] -or $Value -is [decimal]) {
        return [double]$Value
    }

    $text = ([string]$Value).Trim()
    $styles = [System.Globalization.NumberStyles]::Any
    $cultures = @(
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.CultureInfo]::GetCultureInfo("es-ES"),
        [System.Globalization.CultureInfo]::CurrentCulture
    )

    foreach ($culture in $cultures) {
        $number = 0.0
        if ([double]::TryParse($text, $styles, $culture, [ref]$number)) {
            return $number
        }
    }

    throw "No se pudo interpretar como numero: $Value"
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

function Get-ValidationCount {
    param([object]$Worksheet)

    try {
        return [int]$Worksheet.UsedRange.SpecialCells($xlCellTypeAllValidation).Count
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

function Test-IsRedFont {
    param([object]$Range)

    try {
        $color = [int]$Range.Font.Color
        $colorIndex = [int]$Range.Font.ColorIndex
        return ($color -eq 255 -or $colorIndex -eq 3)
    }
    catch {
        return $false
    }
}

function Assert-Baseline {
    param(
        [object]$Workbook,
        [array]$Baseline,
        [string]$Label,
        [System.Collections.Generic.List[string]]$Failures
    )

    foreach ($item in $Baseline) {
        $worksheet = Get-WorksheetByName -Workbook $Workbook -Name $item.Hoja
        $range = $worksheet.Range($item.Celda)
        $actual = $range.Value2
        $formula = [string]$range.Formula
        $ok = $false

        switch ($item.Tipo) {
            "Text" {
                $ok = ((Convert-ToCellText $actual) -eq $item.Esperado)
            }
            "Number" {
                $expected = Convert-ToDoubleInvariant $item.Esperado
                $actualNumber = Convert-ToDoubleInvariant $actual
                $tolerance = if ([string]::IsNullOrWhiteSpace($item.Tolerancia)) { 0.000001 } else { Convert-ToDoubleInvariant $item.Tolerancia }
                $ok = ($null -ne $actualNumber -and [math]::Abs($actualNumber - $expected) -le $tolerance)
            }
            "Blank" {
                $ok = ($null -eq $actual -or [string]::IsNullOrWhiteSpace([string]$actual))
            }
            "NotBlank" {
                $ok = -not ($null -eq $actual -or [string]::IsNullOrWhiteSpace([string]$actual))
            }
            "Formula" {
                $ok = ($formula -eq $item.Esperado)
            }
            "RedText" {
                $ok = ((Convert-ToCellText $actual) -eq $item.Esperado -and (Test-IsRedFont -Range $range))
            }
            default {
                throw "Tipo de comparacion no soportado en baseline IPF: $($item.Tipo)"
            }
        }

        if (-not $ok) {
            $failures.Add("$Label $($item.Hoja)!$($item.Celda) [$($item.Tipo)] esperado='$($item.Esperado)' actual='$actual' formula='$formula'")
        }
    }
}

function Assert-StructureAgainstTemplate {
    param(
        [object]$Workbook,
        [object]$TemplateWorkbook,
        [string]$Label,
        [System.Collections.Generic.List[string]]$Failures
    )

    $templateSheets = Get-WorksheetNames -Workbook $TemplateWorkbook
    $candidateSheets = Get-WorksheetNames -Workbook $Workbook

    foreach ($sheet in $requiredSheets) {
        if (-not (Test-WorksheetExists -Workbook $Workbook -Name $sheet)) {
            $failures.Add("$Label falta hoja obligatoria: $sheet")
        }
    }

    foreach ($sheet in $forbiddenSheets) {
        if (Test-WorksheetExists -Workbook $Workbook -Name $sheet) {
            $failures.Add("$Label contiene una hoja no permitida: $sheet")
        }
    }

    if ($candidateSheets.Count -ne $templateSheets.Count) {
        $failures.Add("$Label numero de hojas distinto a plantilla. Plantilla=$($templateSheets.Count); IPF=$($candidateSheets.Count)")
    }

    foreach ($sheet in $templateSheets) {
        if (-not (Test-WorksheetExists -Workbook $Workbook -Name $sheet)) {
            $failures.Add("$Label no conserva la hoja de plantilla: $sheet")
            continue
        }

        $templateSheet = Get-WorksheetByName -Workbook $TemplateWorkbook -Name $sheet
        $candidateSheet = Get-WorksheetByName -Workbook $Workbook -Name $sheet

        if ([int]$candidateSheet.Visible -ne [int]$templateSheet.Visible) {
            $failures.Add("$Label visibilidad distinta en hoja $sheet. Plantilla=$($templateSheet.Visible); IPF=$($candidateSheet.Visible)")
        }

        $templateDimensions = Get-UsedDimensions -Worksheet $templateSheet
        $candidateDimensions = Get-UsedDimensions -Worksheet $candidateSheet

        if ($candidateDimensions.Rows -ne $templateDimensions.Rows -or $candidateDimensions.Columns -ne $templateDimensions.Columns) {
            $failures.Add("$Label $sheet dimensiones distintas. Plantilla=$($templateDimensions.Rows)x$($templateDimensions.Columns); IPF=$($candidateDimensions.Rows)x$($candidateDimensions.Columns)")
        }

        $templateFormulaCount = Get-FormulaCount -Worksheet $templateSheet
        $candidateFormulaCount = Get-FormulaCount -Worksheet $candidateSheet

        if ($candidateFormulaCount -ne $templateFormulaCount) {
            $failures.Add("$Label $sheet formulas distintas. Plantilla=$templateFormulaCount; IPF=$candidateFormulaCount")
        }

        $templateValidationCount = Get-ValidationCount -Worksheet $templateSheet
        $candidateValidationCount = Get-ValidationCount -Worksheet $candidateSheet

        if ($candidateValidationCount -ne $templateValidationCount) {
            $failures.Add("$Label $sheet validaciones distintas. Plantilla=$templateValidationCount; IPF=$candidateValidationCount")
        }

        $formulaErrors = Get-FormulaErrorAddress -Worksheet $candidateSheet

        if (-not [string]::IsNullOrWhiteSpace($formulaErrors)) {
            $failures.Add("$Label $sheet contiene formulas con error en: $formulaErrors")
        }
    }

    $invoiceSheet = Get-WorksheetByName -Workbook $Workbook -Name "Invoice Request"
    foreach ($cell in $validationCells) {
        $validationType = $null

        try {
            $validationType = [int]$invoiceSheet.Range($cell).Validation.Type
        }
        catch {
            $validationType = $null
        }

        if ($validationType -ne $xlValidateList) {
            $failures.Add("$Label Invoice Request!$cell no conserva validacion de lista. Tipo=$validationType")
        }
    }
}

$templateResolved = Resolve-RequiredFile -Path $TemplatePath -Label "Plantilla IPF" -MissingHint "Revisa la ruta de la plantilla oficial."
$exampleResolved = Resolve-RequiredFile -Path $ExamplePath -Label "Ejemplo GAP IPF" -MissingHint "Revisa la ruta del ejemplo oficial."
$referenceBaselineResolved = Resolve-RequiredFile -Path $ReferenceBaselinePath -Label "Baseline IPF de referencias" -MissingHint "Revisa test\baseline_ipf_referencias.csv."
$generatedBaselineResolved = Resolve-RequiredFile -Path $GeneratedBaselinePath -Label "Baseline IPF generado" -MissingHint "Revisa test\baseline_ipf_generado.csv."

if (-not $SkipGenerated) {
    $ipfResolved = Resolve-RequiredFile -Path $IpfPath -Label "IPF generado" -MissingHint "Ejecuta primero test\test_generar_ipf.xaml o test\test_generar_ipf_datos_incompletos.xaml desde UiPath Studio 23.10.4."
}

$referenceBaseline = Import-Csv -LiteralPath $referenceBaselineResolved -Encoding UTF8
$generatedBaseline = Import-Csv -LiteralPath $generatedBaselineResolved -Encoding UTF8 | Where-Object { $_.Escenario -eq $Scenario }

if (-not $SkipGenerated -and $generatedBaseline.Count -eq 0) {
    throw "No hay baseline IPF para el escenario: $Scenario"
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$failures = New-Object System.Collections.Generic.List[string]
$templateWorkbook = $null
$exampleWorkbook = $null
$ipfWorkbook = $null

try {
    $templateWorkbook = $excel.Workbooks.Open($templateResolved, 0, $true)
    $exampleWorkbook = $excel.Workbooks.Open($exampleResolved, 0, $true)

    Assert-StructureAgainstTemplate -Workbook $exampleWorkbook -TemplateWorkbook $templateWorkbook -Label "Ejemplo GAP" -Failures $failures
    Assert-Baseline -Workbook $templateWorkbook -Baseline ($referenceBaseline | Where-Object { $_.Libro -eq "PlantillaBase" }) -Label "PlantillaBase" -Failures $failures
    Assert-Baseline -Workbook $exampleWorkbook -Baseline ($referenceBaseline | Where-Object { $_.Libro -eq "EjemploGAP" }) -Label "EjemploGAP" -Failures $failures

    if (-not $SkipGenerated) {
        $ipfWorkbook = $excel.Workbooks.Open($ipfResolved, 0, $false)
        try {
            $ipfWorkbook.ForceFullCalculation = $true
        }
        catch {
        }

        $excel.CalculateFullRebuild()
        $ipfWorkbook.Save()
        $ipfWorkbook.Close($true) | Out-Null
        $ipfWorkbook = $null

        $ipfWorkbook = $excel.Workbooks.Open($ipfResolved, 0, $true)
        Assert-StructureAgainstTemplate -Workbook $ipfWorkbook -TemplateWorkbook $templateWorkbook -Label "IPF generado $Scenario" -Failures $failures
        Assert-Baseline -Workbook $ipfWorkbook -Baseline $generatedBaseline -Label "IPF generado $Scenario" -Failures $failures
    }
}
finally {
    if ($ipfWorkbook) {
        $ipfWorkbook.Close($false) | Out-Null
    }

    if ($exampleWorkbook) {
        $exampleWorkbook.Close($false) | Out-Null
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
    throw "Integridad IPF no superada. Fallos: $($failures.Count)"
}

if ($SkipGenerated) {
    Write-Host "Integridad IPF de referencias superada: plantilla base y ejemplo GAP validados."
}
else {
    Write-Host "Integridad IPF superada para escenario $Scenario en $ipfResolved"
}
