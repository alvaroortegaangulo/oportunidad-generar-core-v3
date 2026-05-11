param(
    [string]$WorkflowPath = ".\oportunidad-generar-ipf.xaml",
    [string]$WrapperCompletoPath = ".\test\test_generar_ipf.xaml",
    [string]$WrapperIncompletoPath = ".\test\test_generar_ipf_datos_incompletos.xaml",
    [string]$WrapperCoreIpfPath = ".\test\ejemplo_generar_core_ipf.xaml"
)

$ErrorActionPreference = "Stop"

function Resolve-RequiredFile {
    param(
        [string]$Path,
        [string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$Label no existe: $Path"
    }

    return (Resolve-Path -LiteralPath $Path).Path
}

function Assert-ContainsText {
    param(
        [System.Collections.Generic.List[string]]$Failures,
        [string]$Text,
        [string]$Case,
        [string]$Pattern
    )

    $decodedText = [System.Net.WebUtility]::HtmlDecode($Text)

    if (-not ($Text.Contains($Pattern) -or $decodedText.Contains($Pattern))) {
        $Failures.Add("$Case no encontrado: $Pattern")
    }
}

function Assert-NotContainsText {
    param(
        [System.Collections.Generic.List[string]]$Failures,
        [string]$Text,
        [string]$Case,
        [string]$Pattern
    )

    $decodedText = [System.Net.WebUtility]::HtmlDecode($Text)

    if ($Text.Contains($Pattern) -or $decodedText.Contains($Pattern)) {
        $Failures.Add("$Case no permitido en modulo IPF: $Pattern")
    }
}

$workflowResolved = Resolve-RequiredFile -Path $WorkflowPath -Label "Workflow IPF"
$wrapperCompletoResolved = Resolve-RequiredFile -Path $WrapperCompletoPath -Label "Wrapper IPF completo"
$wrapperIncompletoResolved = Resolve-RequiredFile -Path $WrapperIncompletoPath -Label "Wrapper IPF datos incompletos"
$wrapperCoreIpfResolved = Resolve-RequiredFile -Path $WrapperCoreIpfPath -Label "Wrapper CORE + IPF"

$workflowText = Get-Content -Raw -LiteralPath $workflowResolved
$wrapperCompletoText = Get-Content -Raw -LiteralPath $wrapperCompletoResolved
$wrapperIncompletoText = Get-Content -Raw -LiteralPath $wrapperIncompletoResolved
$wrapperCoreIpfText = Get-Content -Raw -LiteralPath $wrapperCoreIpfResolved

[xml]$null = $workflowText
[xml]$null = $wrapperCompletoText
[xml]$null = $wrapperIncompletoText
[xml]$null = $wrapperCoreIpfText

$failures = New-Object System.Collections.Generic.List[string]

$blockingCases = @(
    @{ Case = "Plantilla IPF no informada"; Pattern = "in_RutaPlantillaIPF no está informada." },
    @{ Case = "Plantilla IPF inexistente"; Pattern = "No existe la plantilla IPF indicada:" },
    @{ Case = "Ruta IPF no informada"; Pattern = "in_RutaIPF no está informada." },
    @{ Case = "Ruta IPF sin carpeta"; Pattern = "in_RutaIPF debe incluir una carpeta destino completa:" },
    @{ Case = "Carpeta destino IPF inexistente"; Pattern = "La carpeta destino del IPF no existe. Debe crearla el proceso padre:" },
    @{ Case = "Error IPF relanzado"; Pattern = "ERROR oportunidad-generar-ipf" },
    @{ Case = "Rethrow al padre"; Pattern = "Relanzar excepción al proceso padre" }
)

$functionalCases = @(
    @{ Case = "Request Type invalido"; Literal = "arg_RequestTypeValor = ValorLista("; Red = "Aplicar rojo a Request Type C3" },
    @{ Case = "Bill Type invalido"; Literal = 'ValorLista(arg_BillTypeEntrada, billTypes, "Fixed"'; Red = "Aplicar rojo a Bill Type C6" },
    @{ Case = "Currency invalida"; Literal = 'ValorLista(arg_CurrencyEntrada, currencies, "EUR"'; Red = "Aplicar rojo a Currency F16" },
    @{ Case = "Company issuer faltante"; Literal = "Company issuer no disponible"; Red = "Aplicar rojo a Company Issuer C8" },
    @{ Case = "Customer faltante"; Literal = "Cliente no disponible"; Red = "Aplicar rojo a Customer C11" },
    @{ Case = "Sold-to faltante"; Literal = "Sold-to no disponible"; Red = "Aplicar rojo a Sold-to F11" },
    @{ Case = "Requested by faltante"; Literal = "Solicitante no disponible"; Red = "Aplicar rojo a Requested by C16" },
    @{ Case = "Sales Order faltante"; Literal = "Sales Order no disponible"; Red = "Aplicar rojo a Sales Order D46" },
    @{ Case = "WBS faltante"; Literal = "WBS no disponible"; Red = "Aplicar rojo a WBS D47" },
    @{ Case = "Importe invalido"; Literal = "Importe no válido"; Red = "Aplicar rojo a importe neto F32" }
)

foreach ($case in $blockingCases) {
    Assert-ContainsText -Failures $failures -Text $workflowText -Case $case.Case -Pattern $case.Pattern
}

foreach ($case in $functionalCases) {
    Assert-ContainsText -Failures $failures -Text $workflowText -Case "$($case.Case) literal" -Pattern $case.Literal
    Assert-ContainsText -Failures $failures -Text $workflowText -Case "$($case.Case) rojo" -Pattern $case.Red
}

Assert-ContainsText -Failures $failures -Text $workflowText -Case "Limpieza direccion" -Pattern "Vaciar dirección de facturación no confirmada C18:C24"
Assert-ContainsText -Failures $failures -Text $workflowText -Case "Limpieza narrativa" -Pattern "Vaciar narrativa ejemplo C29:F38"
Assert-ContainsText -Failures $failures -Text $workflowText -Case "Limpieza comentarios" -Pattern "Vaciar comentarios ejemplo F47:G57"
Assert-ContainsText -Failures $failures -Text $workflowText -Case "Narrativa no inventada" -Pattern "arg_TextNarrativeValor = Texto(arg_TextNarrativeEntrada);"
Assert-ContainsText -Failures $failures -Text $workflowText -Case "Summary no inventado" -Pattern "arg_SummaryTextValor = Texto(arg_SummaryTextEntrada);"
Assert-ContainsText -Failures $failures -Text $workflowText -Case "Comentarios no inventados" -Pattern "arg_ComentariosValor = Texto(arg_ComentariosEntrada);"
Assert-NotContainsText -Failures $failures -Text $workflowText -Case "Hito hardcodeado" -Pattern "Hito 1"
Assert-NotContainsText -Failures $failures -Text $workflowText -Case "Auditoria externa" -Pattern "Trazabilidad_RPA"
Assert-NotContainsText -Failures $failures -Text $workflowText -Case "Markdown externo" -Pattern ".md"

Assert-ContainsText -Failures $failures -Text $wrapperCompletoText -Case "Wrapper completo invoca IPF" -Pattern 'WorkflowFileName="oportunidad-generar-ipf.xaml"'
Assert-ContainsText -Failures $failures -Text $wrapperCompletoText -Case "Wrapper completo salida controlada" -Pattern "IPF_20250256445_test.xlsx"
Assert-ContainsText -Failures $failures -Text $wrapperIncompletoText -Case "Wrapper incompleto invoca IPF" -Pattern 'WorkflowFileName="oportunidad-generar-ipf.xaml"'
Assert-ContainsText -Failures $failures -Text $wrapperIncompletoText -Case "Wrapper incompleto salida controlada" -Pattern "IPF_20250256445_datos_incompletos.xlsx"
Assert-ContainsText -Failures $failures -Text $wrapperIncompletoText -Case "Wrapper incompleto deja WBS vacio" -Pattern 'x:Key="in_SAPCodigoPEPWBS">[String.Empty]'
Assert-ContainsText -Failures $failures -Text $wrapperIncompletoText -Case "Wrapper incompleto deja narrativa vacia" -Pattern 'x:Key="in_IPFTextNarrative">[String.Empty]'

$coreInvokeIndex = $wrapperCoreIpfText.IndexOf('WorkflowFileName="oportunidad-generar-core.xaml"', [System.StringComparison]::Ordinal)
$ipfInvokeIndex = $wrapperCoreIpfText.IndexOf('WorkflowFileName="oportunidad-generar-ipf.xaml"', [System.StringComparison]::Ordinal)

if ($coreInvokeIndex -lt 0) {
    $failures.Add("Wrapper CORE + IPF no invoca oportunidad-generar-core.xaml")
}

if ($ipfInvokeIndex -lt 0) {
    $failures.Add("Wrapper CORE + IPF no invoca oportunidad-generar-ipf.xaml")
}

if ($coreInvokeIndex -ge 0 -and $ipfInvokeIndex -ge 0 -and $coreInvokeIndex -gt $ipfInvokeIndex) {
    $failures.Add("Wrapper CORE + IPF debe invocar CORE antes que IPF")
}

Assert-ContainsText -Failures $failures -Text $wrapperCoreIpfText -Case "Wrapper CORE + IPF condicional" -Pattern 'Condition="[GenerarIPF]"'
Assert-ContainsText -Failures $failures -Text $wrapperCoreIpfText -Case "Wrapper CORE + IPF salida controlada" -Pattern '".local", "handover-core-ipf-output"'

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Contrato negativo IPF no superado. Fallos: $($failures.Count)"
}

Write-Host "Contrato negativo IPF validado estaticamente: $($blockingCases.Count) bloqueantes, $($functionalCases.Count) funcionales y wrappers S13 revisados."
