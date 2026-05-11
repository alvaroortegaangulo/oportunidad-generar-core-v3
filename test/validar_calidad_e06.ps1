param(
    [string]$UiPathCliPath = "C:\Program Files\UiPath\Studio\UiPath.Studio.CommandLine.exe",
    [switch]$SkipAnalyzer,
    [switch]$SkipGeneratedWorkbooks,
    [switch]$RequireGeneratedWorkbooks
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$failures = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]

$acceptedAnalyzerCodes = @(
    "ST-ANA-009",
    "ST-DBP-002",
    "ST-MRD-009",
    "ST-NMG-002",
    "ST-NMG-009",
    "ST-NMG-011",
    "ST-NMG-016",
    "ST-USG-020"
)

$acceptedAnalyzerReasons = @{
    "ST-ANA-009" = "Informativo de recuento de actividades."
    "ST-DBP-002" = "Contrato publico CORE/common con argumentos individuales mantenido por decision de integracion."
    "ST-MRD-009" = "Anidacion heredada de actividades modernas Excel; no impide carga ni ejecucion focalizada."
    "ST-NMG-002" = "Nombres humanos con segmentos SAP/SF/WBS aceptados por claridad de contrato."
    "ST-NMG-009" = "Nombres descriptivos largos aceptados por legibilidad funcional."
    "ST-NMG-011" = "Variables/argumentos auxiliares de subworkflows aceptados si no hay error de carga."
    "ST-NMG-016" = "Argumentos largos aceptados cuando reducen ambiguedad para el proceso padre."
    "ST-USG-020" = "Common workflows puros sin logs propios; la fachada registra inicio, rendimiento, fin y error."
}

function Get-RepoPath {
    param([string]$Path)

    $resolved = (Resolve-Path -LiteralPath $Path).Path

    if ($resolved.StartsWith($repoRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $resolved.Substring($repoRoot.Length).TrimStart("\")
    }

    return $resolved
}

function Invoke-QualityStep {
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    Write-Host "== $Name"

    try {
        & $Action
    }
    catch {
        $failures.Add("$Name fallo: $($_.Exception.Message)")
    }
}

function Test-XamlLoad {
    param([string[]]$Paths)

    foreach ($path in $Paths) {
        $resolved = Join-Path $repoRoot $path

        if (-not (Test-Path -LiteralPath $resolved)) {
            $failures.Add("No existe XAML focalizado: $path")
            continue
        }

        [xml]$null = Get-Content -Raw -LiteralPath $resolved
        Write-Host ("- XML OK: {0}" -f $path)
    }
}

function Invoke-WorkflowAnalyzer {
    param([string[]]$Paths)

    if ($SkipAnalyzer) {
        $warnings.Add("Workflow Analyzer omitido por parametro -SkipAnalyzer.")
        return
    }

    if (-not (Test-Path -LiteralPath $UiPathCliPath)) {
        $failures.Add("No existe UiPath Studio CommandLine en: $UiPathCliPath")
        return
    }

    foreach ($path in $Paths) {
        $resolved = (Resolve-Path -LiteralPath (Join-Path $repoRoot $path)).Path
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $UiPathCliPath
        $psi.Arguments = "analyze-file -p `"$resolved`""
        $psi.WorkingDirectory = $repoRoot
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true

        $process = [System.Diagnostics.Process]::Start($psi)
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()

        $combined = ($stdout + [Environment]::NewLine + $stderr).Trim()
        $codes = [regex]::Matches($combined, '"[^"]+-ErrorCode"\s*:\s*"(?<Code>ST-[A-Z]+-\d+)"') |
            ForEach-Object { $_.Groups["Code"].Value } |
            Sort-Object -Unique

        $unexpected = @($codes | Where-Object { $acceptedAnalyzerCodes -notcontains $_ })
        $loadFailure = $combined -match 'proyecto ya esta abierto|proyecto ya está abierto|project already open|no se pudo cargar|failed to load|parseo|parsing'

        if ($unexpected.Count -gt 0) {
            $failures.Add("Analyzer $path devuelve reglas no aceptadas: $($unexpected -join ', ')")
            continue
        }

        if ($loadFailure) {
            $failures.Add("Analyzer $path no pudo cargar de forma estable: $combined")
            continue
        }

        if ($process.ExitCode -ne 0 -and $codes.Count -eq 0) {
            $failures.Add("Analyzer $path fallo sin codigos de regla parseables. Exit=$($process.ExitCode). Salida=$combined")
            continue
        }

        if ($codes.Count -eq 0) {
            Write-Host ("- Analyzer OK: {0} | sin incidencias" -f $path)
        }
        else {
            $reasonSummary = ($codes | ForEach-Object { "$_=$($acceptedAnalyzerReasons[$_])" }) -join " | "
            Write-Host ("- Analyzer OK: {0} | aceptadas: {1}" -f $path, ($codes -join ", "))
            Write-Host ("  {0}" -f $reasonSummary)
        }
    }
}

function Invoke-OptionalWorkbookValidation {
    param(
        [string]$Name,
        [string]$Path,
        [scriptblock]$Action
    )

    if ($SkipGeneratedWorkbooks) {
        $warnings.Add("$Name omitido por parametro -SkipGeneratedWorkbooks.")
        return
    }

    $resolved = Join-Path $repoRoot $Path

    if (-not (Test-Path -LiteralPath $resolved)) {
        $message = "$Name pendiente: no existe $Path. Ejecuta el wrapper desde UiPath Studio 23.10.4 y repite el validador."

        if ($RequireGeneratedWorkbooks) {
            $failures.Add($message)
        }
        else {
            $warnings.Add($message)
        }

        return
    }

    & $Action
}

$focalXamls = @(
    "oportunidad-generar-core.xaml",
    "oportunidad-generar-ipf.xaml",
    "oportunidad-generar-core-lib\oportunidad-generar-core-pc.xaml",
    "oportunidad-generar-core-lib\oportunidad-generar-core-at.xaml",
    "oportunidad-generar-core-lib\core-common-construir-modelo-ppo.xaml",
    "oportunidad-generar-core-lib\core-common-preparar-resources.xaml",
    "oportunidad-generar-core-lib\core-common-preparar-cost-planning.xaml",
    "test\test_generar_core_pc.xaml",
    "test\test_generar_core_datos_incompletos.xaml",
    "test\test_generar_ipf.xaml",
    "test\test_generar_ipf_datos_incompletos.xaml",
    "test\ejemplo_generar_core_ipf.xaml"
)

Push-Location -LiteralPath $repoRoot

try {
    Invoke-QualityStep -Name "XML focalizado" -Action { Test-XamlLoad -Paths $focalXamls }
    Invoke-QualityStep -Name "Negativos CORE" -Action { & (Join-Path $PSScriptRoot "validar_negativos_core.ps1") }
    Invoke-QualityStep -Name "Negativos IPF" -Action { & (Join-Path $PSScriptRoot "validar_negativos_ipf.ps1") }
    Invoke-QualityStep -Name "Integridad referencias IPF" -Action { & (Join-Path $PSScriptRoot "validar_integridad_ipf.ps1") -SkipGenerated }
    Invoke-QualityStep -Name "Workflow Analyzer focalizado" -Action { Invoke-WorkflowAnalyzer -Paths $focalXamls }

    Invoke-QualityStep -Name "CORE PC generado completo" -Action {
        Invoke-OptionalWorkbookValidation -Name "CORE PC generado completo" -Path ".local\test-output\CORE_PC_20250256445_test.xlsx" -Action {
            & (Join-Path $PSScriptRoot "validar_integridad_core_pc.ps1")
        }
    }

    Invoke-QualityStep -Name "CORE PC generado datos incompletos" -Action {
        Invoke-OptionalWorkbookValidation -Name "CORE PC generado datos incompletos" -Path ".local\test-output\CORE_PC_20250256445_datos_incompletos.xlsx" -Action {
            & (Join-Path $PSScriptRoot "validar_integridad_core_pc.ps1") -CorePath ".\.local\test-output\CORE_PC_20250256445_datos_incompletos.xlsx" -Scenario DatosIncompletos -SkipBaseline
        }
    }

    Invoke-QualityStep -Name "IPF generado completo" -Action {
        Invoke-OptionalWorkbookValidation -Name "IPF generado completo" -Path ".local\test-output\IPF_20250256445_test.xlsx" -Action {
            & (Join-Path $PSScriptRoot "validar_integridad_ipf.ps1")
        }
    }

    Invoke-QualityStep -Name "IPF generado datos incompletos" -Action {
        Invoke-OptionalWorkbookValidation -Name "IPF generado datos incompletos" -Path ".local\test-output\IPF_20250256445_datos_incompletos.xlsx" -Action {
            & (Join-Path $PSScriptRoot "validar_integridad_ipf.ps1") -IpfPath ".\.local\test-output\IPF_20250256445_datos_incompletos.xlsx" -Scenario DatosIncompletos
        }
    }
}
finally {
    Pop-Location
}

if ($warnings.Count -gt 0) {
    Write-Warning "Advertencias E06:"
    $warnings | ForEach-Object { Write-Warning $_ }
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Cierre de calidad E06 no superado. Fallos: $($failures.Count)"
}

Write-Host "Cierre de calidad E06 superado: XML, negativos CORE/IPF, referencias IPF y Workflow Analyzer focalizado. Las validaciones de Excel generado se ejecutan cuando los wrappers se han lanzado desde Studio."
