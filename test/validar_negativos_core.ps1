param(
    [string]$WorkflowPath = ".\oportunidad-generar-core.xaml"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $WorkflowPath)) {
    throw "No existe el workflow a validar: $WorkflowPath"
}

$workflowResolved = (Resolve-Path -LiteralPath $WorkflowPath).Path
$workflowRoot = Split-Path -Parent $workflowResolved
$workflowFiles = New-Object System.Collections.Generic.List[string]
$workflowFiles.Add($workflowResolved)

$libPath = Join-Path $workflowRoot "oportunidad-generar-core-lib"

if (Test-Path -LiteralPath $libPath) {
    Get-ChildItem -LiteralPath $libPath -Filter "*.xaml" |
        Sort-Object FullName |
        ForEach-Object { $workflowFiles.Add($_.FullName) }
}

$invokeAuditFiles = New-Object System.Collections.Generic.List[string]
$workflowFiles | ForEach-Object { $invokeAuditFiles.Add($_) }

$ipfPath = Join-Path $workflowRoot "oportunidad-generar-ipf.xaml"

if (Test-Path -LiteralPath $ipfPath) {
    $invokeAuditFiles.Add((Resolve-Path -LiteralPath $ipfPath).Path)
}

$workflowText = ($workflowFiles | ForEach-Object {
    Get-Content -Raw -LiteralPath $_
}) -join [Environment]::NewLine
$workflowTextDecoded = [System.Net.WebUtility]::HtmlDecode($workflowText)

# XML parse first: this also proves the static checks are reading a loadable XAML file.
foreach ($workflowFile in $workflowFiles) {
    [xml]$null = Get-Content -Raw -LiteralPath $workflowFile
}

$failures = New-Object System.Collections.Generic.List[string]

function Assert-ContainsText {
    param(
        [string]$Case,
        [string]$Pattern
    )

    if (-not ($workflowText.Contains($Pattern) -or $workflowTextDecoded.Contains($Pattern))) {
        $failures.Add("$Case no encontrado en XAML: $Pattern")
    }
}

$invokeCodeThreshold = @{
    MaxChars = 7000
    MaxLines = 220
    MaxStatements = 120
}

$invokeCodeOversizeJustifications = @(
    @{
        FilePattern = "core-common-preparar-resources.xaml"
        DisplayName = "Utilidad S06 - preparar tablas Resources"
        Reason = "Justificado tras E04: bloque comun no duplicado PC/AT; concentra normalizacion de recursos, intercompany y tablas Resources."
    },
    @{
        FilePattern = "core-common-preparar-cost-planning.xaml"
        DisplayName = "Utilidad S07 - preparar tablas Cost Planning"
        Reason = "Justificado tras E04: bloque comun no duplicado PC/AT; calcula prorrateos, gastos, compras, riesgos y garantia una sola vez."
    },
    @{
        FilePattern = "oportunidad-generar-core.xaml"
        DisplayName = "Utilidad S05 - construir mapa Project Infor"
        Reason = "Justificado tras E05: centraliza el mapa de celdas Project Infor y los textos funcionales revisables en rojo."
    },
    @{
        FilePattern = "core-common-construir-modelo-ppo.xaml"
        DisplayName = "Utilidad S04B - cabecera y fechas PPO"
        Reason = "Justificado tras E04: concentra cabecera, calendario y validaciones de duración del PPO en un único punto común."
    },
    @{
        FilePattern = "core-common-construir-modelo-ppo.xaml"
        DisplayName = "Utilidad S04G - serializar modelo PPO"
        Reason = "Justificado tras E04: serializa el modelo común completo que comparten Project Infor, Resources y Cost Planning."
    },
    @{
        FilePattern = "oportunidad-generar-ipf.xaml"
        DisplayName = "Utilidad S12 - normalizar listas y valores IPF"
        Reason = "Justificado tras S12: agrupa normalización de listas IPF y preparación de valores confirmados antes de escribir la plantilla."
    }
)

function Get-InvokeCodeText {
    param(
        [System.Xml.XmlElement]$Node
    )

    $code = $Node.Code

    if ([string]::IsNullOrEmpty($code)) {
        $codeNode = $Node.SelectSingleNode('*[local-name()="InvokeCode.Code"]')

        if ($null -ne $codeNode) {
            $code = $codeNode.InnerText
        }
    }

    return [System.Net.WebUtility]::HtmlDecode([string]$code)
}

function Get-InvokeCodeJustification {
    param(
        [string]$File,
        [string]$DisplayName
    )

    foreach ($justification in $invokeCodeOversizeJustifications) {
        if ($File.EndsWith($justification.FilePattern, [System.StringComparison]::OrdinalIgnoreCase) -and
            [string]::Equals($DisplayName, $justification.DisplayName, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $justification.Reason
        }
    }

    return $null
}

function Test-InvokeCodeHeader {
    param(
        [string]$Code
    )

    return $Code -match '(?m)^\s*//\s*Objetivo:' -and
        $Code -match '(?m)^\s*//\s*Entradas:' -and
        $Code -match '(?m)^\s*//\s*Salidas:' -and
        $Code -match '(?m)^\s*//\s*Regla' -and
        $Code -match '(?m)^\s*//\s*No hace:'
}

function Test-InvokeCodeMetrics {
    $metrics = New-Object System.Collections.Generic.List[object]

    foreach ($invokeAuditFile in $invokeAuditFiles) {
        [xml]$invokeXml = Get-Content -Raw -LiteralPath $invokeAuditFile
        $nodes = $invokeXml.SelectNodes('//*[local-name()="InvokeCode"]')

        foreach ($node in $nodes) {
            $displayName = [string]$node.DisplayName
            $code = Get-InvokeCodeText -Node $node
            $lines = if ([string]::IsNullOrWhiteSpace($code)) { 0 } else { ($code -split "`r?`n").Count }
            $statements = ([regex]::Matches($code, ';')).Count
            $commentLines = ([regex]::Matches($code, '(?m)^\s*//')).Count
            $justification = Get-InvokeCodeJustification -File $invokeAuditFile -DisplayName $displayName

            $metrics.Add([pscustomobject]@{
                File = (Resolve-Path -LiteralPath $invokeAuditFile -Relative)
                DisplayName = $displayName
                Chars = $code.Length
                Lines = $lines
                Statements = $statements
                CommentLines = $commentLines
                OneLine = ($lines -le 1)
                Header = (Test-InvokeCodeHeader -Code $code)
                OversizeJustified = -not [string]::IsNullOrWhiteSpace($justification)
                Justification = $justification
            })
        }
    }

    Write-Host "Inventario Invoke Code:"

    $metrics |
        Sort-Object File, DisplayName |
        ForEach-Object {
            Write-Host ("- {0} | {1} | chars={2}; lineas={3}; sentencias={4}; comentarios={5}; unaLinea={6}; cabecera={7}; justificado={8}" -f
                $_.File,
                $_.DisplayName,
                $_.Chars,
                $_.Lines,
                $_.Statements,
                $_.CommentLines,
                $_.OneLine,
                $_.Header,
                $_.OversizeJustified)
        }

    foreach ($metric in $metrics) {
        if ($metric.Chars -eq 0 -or $metric.Lines -eq 0) {
            $failures.Add("Invoke Code vacio: $($metric.DisplayName) en $($metric.File)")
        }

        if ($metric.OneLine) {
            $failures.Add("Invoke Code de una sola linea no permitido: $($metric.DisplayName) en $($metric.File)")
        }

        if (-not $metric.Header) {
            $failures.Add("Invoke Code sin cabecera Objetivo/Entradas/Salidas/Regla/No hace: $($metric.DisplayName) en $($metric.File)")
        }

        $isOversize = $metric.Chars -gt $invokeCodeThreshold.MaxChars -or
            $metric.Lines -gt $invokeCodeThreshold.MaxLines -or
            $metric.Statements -gt $invokeCodeThreshold.MaxStatements

        if ($isOversize) {
            if ($metric.OversizeJustified) {
                Write-Warning "Invoke Code supera umbral pero esta justificado: $($metric.DisplayName) en $($metric.File). $($metric.Justification)"
            }
            else {
                $failures.Add("Invoke Code supera umbral sin justificacion: $($metric.DisplayName) en $($metric.File). Chars=$($metric.Chars)/$($invokeCodeThreshold.MaxChars), Lines=$($metric.Lines)/$($invokeCodeThreshold.MaxLines), Statements=$($metric.Statements)/$($invokeCodeThreshold.MaxStatements)")
            }
        }
    }
}

function Assert-TextContains {
    param(
        [System.Collections.Generic.List[string]]$Failures,
        [string]$Text,
        [string]$Case,
        [string]$Pattern
    )

    if (-not $Text.Contains($Pattern)) {
        $Failures.Add("$Case no encontrado: $Pattern")
    }
}

function Test-CoreIncompleteWrapper {
    $wrapperPath = Join-Path $workflowRoot "test\test_generar_core_datos_incompletos.xaml"

    if (-not (Test-Path -LiteralPath $wrapperPath)) {
        $failures.Add("Falta wrapper E06 de datos SAP/SF incompletos CORE: $wrapperPath")
        return
    }

    $wrapperText = Get-Content -Raw -LiteralPath $wrapperPath
    [xml]$null = $wrapperText

    Assert-TextContains -Failures $failures -Text $wrapperText -Case "Wrapper CORE incompleto invoca modulo" -Pattern 'WorkflowFileName="oportunidad-generar-core.xaml"'
    Assert-TextContains -Failures $failures -Text $wrapperText -Case "Wrapper CORE incompleto salida controlada" -Pattern "CORE_PC_20250256445_datos_incompletos.xlsx"
    Assert-TextContains -Failures $failures -Text $wrapperText -Case "Wrapper CORE incompleto WBS vacio" -Pattern 'x:Key="in_SAP_CodigoPEP_WBS">[String.Empty]'
    Assert-TextContains -Failures $failures -Text $wrapperText -Case "Wrapper CORE incompleto cliente SAP vacio" -Pattern 'x:Key="in_SAP_CodigoCliente">[String.Empty]'
    Assert-TextContains -Failures $failures -Text $wrapperText -Case "Wrapper CORE incompleto Sales Order vacio" -Pattern 'x:Key="in_SAP_PedidoSalesOrder">[String.Empty]'
    Assert-TextContains -Failures $failures -Text $wrapperText -Case "Wrapper CORE incompleto compras vacias" -Pattern 'x:Key="in_SAP_ProveedorCompra">[String.Empty]'
    Assert-TextContains -Failures $failures -Text $wrapperText -Case "Wrapper CORE incompleto recursos sin JSON" -Pattern 'x:Key="in_SAP_CodigosEmpleadoRecursosJson">[String.Empty]'
}

function Assert-BaselineCoreContract {
    $baselinePath = Join-Path $workflowRoot "test\baseline_core_pc.csv"

    if (-not (Test-Path -LiteralPath $baselinePath)) {
        $failures.Add("Falta baseline CORE PC: $baselinePath")
        return
    }

    $baseline = Import-Csv -LiteralPath $baselinePath -Encoding UTF8

    if ($baseline.Count -lt 58) {
        $failures.Add("El baseline CORE PC debe conservar al menos las 58 celdas criticas de S09/E06. Detectadas: $($baseline.Count)")
    }

    $duplicates = $baseline |
        Group-Object Hoja, Celda |
        Where-Object { $_.Count -gt 1 }

    foreach ($duplicate in $duplicates) {
        $failures.Add("Baseline CORE PC contiene celda duplicada: $($duplicate.Name)")
    }

    $criticalCells = @(
        @{ Area = "Project Infor"; Hoja = "Project Infor"; Celda = "C10"; Tipo = "Text"; Esperado = "ZEV-PCE0012-1001" },
        @{ Area = "Project Infor"; Hoja = "Project Infor"; Celda = "C28"; Tipo = "Number"; Esperado = "583896.55" },
        @{ Area = "Resources"; Hoja = "Resources"; Celda = "D8"; Tipo = "Text"; Esperado = "INTERCO" },
        @{ Area = "Cost Planning"; Hoja = "Cost Planning"; Celda = "B64"; Tipo = "Text"; Esperado = "Proveedor Test" },
        @{ Area = "Cost Planning"; Hoja = "Cost Planning"; Celda = "R47"; Tipo = "Number"; Esperado = "18375.757042" }
    )

    foreach ($criticalCell in $criticalCells) {
        $match = $baseline | Where-Object {
            $_.Area -eq $criticalCell.Area -and
            $_.Hoja -eq $criticalCell.Hoja -and
            $_.Celda -eq $criticalCell.Celda -and
            $_.Tipo -eq $criticalCell.Tipo -and
            $_.Esperado -eq $criticalCell.Esperado
        }

        if (-not $match) {
            $failures.Add("Baseline CORE PC no conserva la celda critica esperada: $($criticalCell.Hoja)!$($criticalCell.Celda)=$($criticalCell.Esperado)")
        }
    }
}

$blockingCases = @(
    @{ Case = "PPO no informado"; Pattern = "in_RutaPPO no está informada." },
    @{ Case = "PPO inexistente"; Pattern = "No existe el PPO indicado:" },
    @{ Case = "Ruta CORE no informada"; Pattern = "in_RutaCORE no está informada." },
    @{ Case = "Carpeta destino inexistente"; Pattern = "La carpeta destino del CORE no existe. Debe crearla el proceso padre:" },
    @{ Case = "Tipo proyecto invalido"; Pattern = "in_TipoProyecto debe ser 'PC' o 'AT'. Valor recibido:" },
    @{ Case = "Plantilla no informada"; Pattern = "No se ha informado la ruta de plantilla CORE para el tipo" },
    @{ Case = "Plantilla inexistente"; Pattern = "No existe la plantilla CORE seleccionada:" },
    @{ Case = "Hoja obligatoria ausente"; Pattern = "La plantilla CORE no contiene la hoja obligatoria:" },
    @{ Case = "Duracion PPO no valida"; Pattern = "La duración del PPO no está informada o no es válida en Hoja de datos!B29." },
    @{ Case = "Opportunity Number no leido"; Pattern = "No se ha podido leer Opportunity Number en Hoja de datos!B22." },
    @{ Case = "Capacidad plantilla no detectable"; Pattern = "No se ha podido detectar capacidad válida en plantilla CORE para" },
    @{ Case = "Capacidad Resources insuficiente"; Pattern = "perfiles útiles en Resources. Perfiles detectados:" },
    @{ Case = "Capacidad Cost Planning insuficiente"; Pattern = "perfiles útiles en Cost Planning. Perfiles detectados:" },
    @{ Case = "Capacidad gastos insuficiente"; Pattern = "líneas de gastos en Cost Planning. Gastos detectados:" },
    @{ Case = "Capacidad compras insuficiente"; Pattern = "líneas de compras en Cost Planning. Compras detectadas:" }
)

$projectInforFunctionalCases = @(
    @{ Case = "WBS faltante"; Literal = "WBS no disponible"; Cell = "C10" },
    @{ Case = "Cliente SAP faltante"; Literal = "Cliente SAP no disponible"; Cell = "C16" },
    @{ Case = "Sales Order faltante"; Literal = "Sales Order no disponible"; Cell = "C19" },
    @{ Case = "PM codigo faltante"; Literal = "PM cod no disponible"; Cell = "C22" },
    @{ Case = "PM nombre faltante"; Literal = "PM no disponible"; Cell = "D22" },
    @{ Case = "Practice leader codigo faltante"; Literal = "PL cod no disponible"; Cell = "C23" },
    @{ Case = "Practice leader nombre faltante"; Literal = "Practice leader no disponible"; Cell = "D23" },
    @{ Case = "Sales manager codigo faltante"; Literal = "Sales manager cod no disponible"; Cell = "C24" },
    @{ Case = "Sales manager nombre faltante"; Literal = "Sales manager no disponible"; Cell = "D24" },
    @{ Case = "Descripcion PPO faltante"; Literal = "Descripción no disponible en PPO"; Cell = "C11" },
    @{ Case = "Practica faltante"; Literal = "Práctica no disponible"; Cell = "C12" },
    @{ Case = "Cliente fallback faltante"; Literal = "Cliente no disponible"; Cell = "D16" },
    @{ Case = "Importe faltante"; Literal = "Importe no disponible en PPO"; Cell = "C28" },
    @{ Case = "Horas faltantes"; Literal = "Horas no disponibles en PPO"; Cell = "C29" },
    @{ Case = "Coste recursos faltante"; Literal = "Coste recursos no disponible"; Cell = "C30" },
    @{ Case = "Riesgos faltantes"; Literal = "Riesgos no disponibles en PPO"; Cell = "C31" },
    @{ Case = "Garantia faltante"; Literal = "Garantía no disponible en PPO"; Cell = "C32" },
    @{ Case = "Gastos faltantes"; Literal = "Gastos no disponibles en PPO"; Cell = "C33" },
    @{ Case = "Compras faltantes"; Literal = "Compras no disponibles en PPO"; Cell = "C34" }
)

$functionalCases = @(
    @{ Case = "Intercompany no determinable"; Literal = "Intercompany no determinable"; Red = "Aplicar rojo funcional CORE" },
    @{ Case = "Entidad recurso faltante"; Literal = "Entidad recurso no disponible"; Red = "Aplicar rojo funcional CORE" },
    @{ Case = "Codigo recurso faltante"; Literal = "empleado no disponible"; Red = "Aplicar rojo funcional CORE" },
    @{ Case = "Coste hora faltante"; Literal = "Coste no disponible"; Red = "Aplicar rojo funcional CORE" },
    @{ Case = "Tarifa AT faltante"; Literal = "Tarifa AT no disponible"; Red = "Aplicar rojo funcional CORE" },
    @{ Case = "Proveedor compra faltante"; Literal = "Proveedor no disponible"; Red = "Aplicar rojo funcional CORE" },
    @{ Case = "Pedido compra faltante"; Literal = "PO compra no disponible"; Red = "Aplicar rojo funcional CORE" },
    @{ Case = "Codigo Ariba faltante"; Literal = "Ariba no disponible"; Red = "Aplicar rojo funcional CORE" }
)

foreach ($case in $blockingCases) {
    Assert-ContainsText -Case $case.Case -Pattern $case.Pattern
}

foreach ($case in $projectInforFunctionalCases) {
    Assert-ContainsText -Case "$($case.Case) literal" -Pattern $case.Literal
    Assert-ContainsText -Case "$($case.Case) celda Project Infor" -Pattern $case.Cell
}

foreach ($case in $functionalCases) {
    Assert-ContainsText -Case "$($case.Case) literal" -Pattern $case.Literal
    Assert-ContainsText -Case "$($case.Case) rojo" -Pattern $case.Red
}

Assert-ContainsText -Case "Bloques Project Infor 05A" -Pattern "05A Fechas e identificadores"
Assert-ContainsText -Case "Bloques Project Infor 05B" -Pattern "05B Cliente y práctica"
Assert-ContainsText -Case "Bloques Project Infor 05C" -Pattern "05C Responsables"
Assert-ContainsText -Case "Bloques Project Infor 05D" -Pattern "05D Totales económicos"
Assert-ContainsText -Case "Mapa Project Infor" -Pattern "dtMapaProjectInfor"
Assert-ContainsText -Case "Errores funcionales Project Infor" -Pattern "dtErroresFuncionales"
Assert-ContainsText -Case "Rojo dinamico Project Infor" -Pattern "Aplicar rojo Project Infor desde mapa"
Assert-ContainsText -Case "Escritura dinamica Project Infor" -Pattern 'Cell(filaProjectInfor("Celda").ToString())'

Assert-ContainsText -Case "Lectura PPO por rangos" -Pattern "03 Leer PPO por rangos estables"
Assert-ContainsText -Case "Mapa lectura PPO" -Pattern "dtMapaLecturaPPO"
Assert-ContainsText -Case "Subworkflow construir modelo PPO" -Pattern "core-common-construir-modelo-ppo.xaml"
Assert-ContainsText -Case "E05 lectura estructura plantilla con Excel" -Pattern "Leer estructura de plantilla CORE"
Assert-ContainsText -Case "E05 rendimiento lectura PPO" -Pattern "Segundos="
Assert-ContainsText -Case "E05 rendimiento escritura CORE" -Pattern "Rendimiento CORE | Escritura CORE segundos="
Assert-ContainsText -Case "E05 rendimiento recalculo final" -Pattern "Rendimiento CORE | Recálculo final segundos="

if ($workflowTextDecoded.Contains("Preparar valores Project Infor desde PPO SAP y Salesforce")) {
    $failures.Add("Project Infor no debe volver a concentrarse en un unico Multiple Assign monolitico; usar bloques 05A-05D y dtMapaProjectInfor.")
}

$projectInforWriteCells = [regex]::Matches($workflowTextDecoded, '<ueab:WriteCellX[^>]+Project Infor')
if ($projectInforWriteCells.Count -gt 3) {
    $failures.Add("Project Infor debe escribirse por mapa dinamico, no mediante Write CellX fijos repetidos. Detectados: $($projectInforWriteCells.Count)")
}

if ($workflowTextDecoded.Contains("<ui:ReadCell")) {
    $failures.Add("La lectura PPO no debe volver a usar Read Cell dispersos; usar rangos estables y dtMapaLecturaPPO.")
}

if ($workflowTextDecoded.Contains("Utilidad S04 - construir modelo desde DataTables PPO")) {
    $failures.Add("La utilidad S04 monolitica no debe reaparecer; usar core-common-construir-modelo-ppo.xaml dividido.")
}

Test-InvokeCodeMetrics

[xml]$mainWorkflowXml = Get-Content -Raw -LiteralPath $workflowResolved
$s08CapacidadNode = $mainWorkflowXml.SelectSingleNode('//*[local-name()="InvokeCode" and @DisplayName="Utilidad S08 - preparar matriz intercompany y capacidad plantilla"]')
if ($null -ne $s08CapacidadNode) {
    $failures.Add("Utilidad S08 - preparar matriz intercompany y capacidad plantilla debe estar implementada con actividades UiPath, no como Invoke Code.")
    $s08CapacidadCode = Get-InvokeCodeText -Node $s08CapacidadNode
    if ($s08CapacidadCode -match 'Type\.GetTypeFromProgID|Workbooks\.Open|ReleaseComObject|Marshal\.ReleaseComObject') {
        $failures.Add("E05 exige que la capacidad de plantilla se mida desde actividades Excel/DataTables, no mediante COM en Utilidad S08 - preparar matriz intercompany y capacidad plantilla.")
    }
}

Assert-ContainsText -Case "S08 capacidad por actividades" -Pattern "Utilidad S08 - preparar matriz intercompany y capacidad plantilla"
Assert-ContainsText -Case "S08 matriz intercompany Build Data Table" -Pattern "Construir dtMatrizIntercompany"
Assert-ContainsText -Case "S08 capacidad AT por actividades" -Pattern "Medir capacidad plantilla AT"
Assert-ContainsText -Case "S08 capacidad PC por actividades" -Pattern "Medir capacidad plantilla PC"
Assert-ContainsText -Case "S08 capacidad AT dinamica" -Pattern "Recorrer reglas capacidad AT"
Assert-ContainsText -Case "S08 capacidad PC dinamica" -Pattern "Recorrer reglas capacidad PC"

$s08CierreFinalNode = $mainWorkflowXml.SelectSingleNode('//*[local-name()="InvokeCode" and @DisplayName="Utilidad S08 - aplicar rojo dinámico y recalcular CORE"]')
if ($null -ne $s08CierreFinalNode) {
    $failures.Add("Utilidad S08 - aplicar rojo dinámico y recalcular CORE debe estar implementada con actividades UiPath, no como Invoke Code.")
}

Assert-ContainsText -Case "S08 rojo funcional por rangos" -Pattern "Recorrer rangos con errores CORE"
Assert-ContainsText -Case "S08 lectura dinamica de rangos para rojo" -Pattern "Leer rango para errores CORE"
Assert-ContainsText -Case "S08 rojo con FormatRangeX" -Pattern "Aplicar rojo funcional CORE"
Assert-ContainsText -Case "S08 recalculo con actividad Excel" -Pattern "Recalcular CORE con VBA controlado"
Assert-ContainsText -Case "S08 VBA externo controlado" -Pattern "core-recalcular-final.txt"

if ($workflowTextDecoded -match 'ReleaseComObject|Marshal\.ReleaseComObject|Type\.GetTypeFromProgID|Workbooks\.Open') {
    $failures.Add("El cierre final CORE no debe conservar COM manual; usar actividades Excel y VBA controlado para recálculo.")
}

$vbaRecalculoPath = Join-Path $workflowRoot "oportunidad-generar-core-lib\core-recalcular-final.txt"
if (-not (Test-Path -LiteralPath $vbaRecalculoPath)) {
    $failures.Add("No existe el VBA controlado de recálculo final: $vbaRecalculoPath")
} else {
    $vbaRecalculoText = Get-Content -Raw -LiteralPath $vbaRecalculoPath
    if ($vbaRecalculoText -notmatch 'CalculateFullRebuild' -or
        $vbaRecalculoText -notmatch 'ForceFullCalculation' -or
        $vbaRecalculoText -notmatch 'FullCalculationOnLoad') {
        $failures.Add("El VBA controlado de recálculo final debe conservar CalculateFullRebuild, ForceFullCalculation y FullCalculationOnLoad.")
    }
}

foreach ($workflowFile in $workflowFiles) {
    $fileName = Split-Path -Leaf $workflowFile
    if ($fileName -in @("oportunidad-generar-core-pc.xaml", "oportunidad-generar-core-at.xaml")) {
        [xml]$workflowXml = Get-Content -Raw -LiteralPath $workflowFile
        $localInvokeCodeCount = $workflowXml.SelectNodes('//*[local-name()="InvokeCode"]').Count
        if ($localInvokeCodeCount -gt 0) {
            $failures.Add("$fileName no debe contener Invoke Code tras E04; debe invocar common workflows y conservar solo escrituras especificas.")
        }

        $excelScopeCount = $workflowXml.SelectNodes('//*[local-name()="ExcelProcessScopeX"]').Count
        $excelAppCount = $workflowXml.SelectNodes('//*[local-name()="ExcelApplicationCard"]').Count
        if ($excelScopeCount -ne 1 -or $excelAppCount -ne 1) {
            $failures.Add("$fileName debe mantener una unica apertura Excel consolidada tras E05. ExcelProcessScopeX=$excelScopeCount; ExcelApplicationCard=$excelAppCount")
        }
    }
}

Assert-ContainsText -Case "Subworkflow comun Resources E04" -Pattern "core-common-preparar-resources.xaml"
Assert-ContainsText -Case "Subworkflow comun Cost Planning E04" -Pattern "core-common-preparar-cost-planning.xaml"
Assert-ContainsText -Case "Resources comun invocado" -Pattern "Preparar tablas comunes de Resources"
Assert-ContainsText -Case "Cost Planning comun invocado" -Pattern "Preparar tablas comunes de Cost Planning"
Assert-ContainsText -Case "Cost Planning reutiliza Resources" -Pattern "arg_dtResourcesATFijos"

$normalizarCompaniaCount = ([regex]::Matches($workflowTextDecoded, 'string\s+NormalizarCompania\s*\(')).Count
$intercompanyActualCount = ([regex]::Matches($workflowTextDecoded, 'string\s+IntercompanyActual\s*\(')).Count
if ($normalizarCompaniaCount -ne 1 -or $intercompanyActualCount -ne 1) {
    $failures.Add("La logica tecnica de intercompany debe quedar en un unico punto comun tras E04. NormalizarCompania=$normalizarCompaniaCount; IntercompanyActual=$intercompanyActualCount")
}

Assert-ContainsText -Case "Formato rojo dinamico" -Pattern "Aplicar rojo funcional CORE"
Assert-ContainsText -Case "Rangos de pintado dinamico" -Pattern "Recorrer rangos con errores CORE"

Test-CoreIncompleteWrapper
Assert-BaselineCoreContract

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Contrato negativo CORE no superado. Fallos: $($failures.Count)"
}

Write-Host "Contrato negativo CORE validado estaticamente: $($blockingCases.Count) bloqueantes, $($projectInforFunctionalCases.Count) Project Infor, $($functionalCases.Count) funcionales dinamicos, wrapper datos incompletos y baseline CORE PC en $($workflowFiles.Count) XAML CORE"
