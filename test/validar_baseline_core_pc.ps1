param(
    [string]$CorePath = ".\.local\test-output\CORE_PC_20250256445_test.xlsx",
    [string]$BaselinePath = ".\test\baseline_core_pc.csv"
)

$ErrorActionPreference = "Stop"

function Convert-ToComparableDate {
    param([object]$Value)

    if ($null -eq $Value -or $Value -eq "") {
        return ""
    }

    if ($Value -is [datetime]) {
        return $Value.ToString("yyyy-MM-dd")
    }

    if ($Value -is [double] -or $Value -is [int] -or $Value -is [decimal]) {
        return [datetime]::FromOADate([double]$Value).ToString("yyyy-MM-dd")
    }

    return ([datetime]::Parse([string]$Value)).ToString("yyyy-MM-dd")
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

if (-not (Test-Path -LiteralPath $CorePath)) {
    throw "No existe el CORE a validar: $CorePath. Ejecuta primero test\test_generar_core_pc.xaml desde UiPath Studio."
}

if (-not (Test-Path -LiteralPath $BaselinePath)) {
    throw "No existe el baseline: $BaselinePath"
}

$baseline = Import-Csv -LiteralPath $BaselinePath -Encoding UTF8
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$failures = New-Object System.Collections.Generic.List[string]

try {
    $workbook = $excel.Workbooks.Open((Resolve-Path -LiteralPath $CorePath).Path, 0, $true)

    foreach ($item in $baseline) {
        $worksheet = $workbook.Worksheets.Item($item.Hoja)
        $actual = $worksheet.Range($item.Celda).Value2
        $ok = $false

        switch ($item.Tipo) {
            "Text" {
                $ok = (([string]$actual).Trim() -eq $item.Esperado)
            }
            "Number" {
                $expected = Convert-ToDoubleInvariant $item.Esperado
                $actualNumber = Convert-ToDoubleInvariant $actual
                $tolerance = if ([string]::IsNullOrWhiteSpace($item.Tolerancia)) { 0.000001 } else { Convert-ToDoubleInvariant $item.Tolerancia }
                $ok = ($null -ne $actualNumber -and [math]::Abs($actualNumber - $expected) -le $tolerance)
            }
            "Date" {
                $ok = ((Convert-ToComparableDate $actual) -eq $item.Esperado)
            }
            "Today" {
                $ok = ((Convert-ToComparableDate $actual) -eq (Get-Date).ToString("yyyy-MM-dd"))
            }
            "Blank" {
                $ok = ($null -eq $actual -or [string]::IsNullOrWhiteSpace([string]$actual))
            }
            default {
                throw "Tipo de comparacion no soportado: $($item.Tipo)"
            }
        }

        if (-not $ok) {
            $failures.Add("$($item.Hoja)!$($item.Celda) [$($item.Tipo)] esperado='$($item.Esperado)' actual='$actual'")
        }
    }
}
finally {
    if ($workbook) {
        $workbook.Close($false) | Out-Null
    }

    $excel.Quit() | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Baseline CORE PC no superado. Fallos: $($failures.Count)"
}

Write-Host "Baseline CORE PC superado: $($baseline.Count) celdas validadas en $CorePath"
