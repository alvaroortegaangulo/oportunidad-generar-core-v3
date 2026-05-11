param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [string]$PcWorkbookPath = '',
    [string]$AtWorkbookPath = '',
    [switch]$SkipWorkbooks
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$failures = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]

function Add-Failure {
    param([string]$Message)
    $failures.Add($Message)
}

function Resolve-ProjectPath {
    param([string]$RelativePath)
    return Join-Path $ProjectRoot $RelativePath
}

function Resolve-FirstExisting {
    param([string[]]$Candidates)

    foreach ($candidate in $Candidates) {
        $resolved = if ([IO.Path]::IsPathRooted($candidate)) { $candidate } else { Resolve-ProjectPath $candidate }
        if (Test-Path -LiteralPath $resolved) {
            return (Resolve-Path -LiteralPath $resolved).Path
        }
    }

    return $null
}

function Assert-TextContains {
    param(
        [string]$Text,
        [string]$Pattern,
        [string]$Case
    )

    if (-not $Text.Contains($Pattern)) {
        Add-Failure "$Case no encontrado: $Pattern"
    }
}

function Assert-TextNotContains {
    param(
        [string]$Text,
        [string]$Pattern,
        [string]$Case
    )

    if ($Text.Contains($Pattern)) {
        Add-Failure "$Case no debe aparecer: $Pattern"
    }
}

function Assert-NearNumber {
    param(
        [object]$Actual,
        [double]$Expected,
        [double]$Tolerance,
        [string]$Case
    )

    $number = 0.0
    if (-not [double]::TryParse([string]$Actual, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        if (-not [double]::TryParse([string]$Actual, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::GetCultureInfo('es-ES'), [ref]$number)) {
            Add-Failure "$Case no es numerico. Valor='$Actual'"
            return
        }
    }

    if ([math]::Abs($number - $Expected) -gt $Tolerance) {
        Add-Failure "$Case esperado=$Expected actual=$number tolerancia=$Tolerance"
    }
}

function Get-ColIndex {
    param([string]$Letters)

    $value = 0
    foreach ($char in $Letters.ToUpperInvariant().ToCharArray()) {
        if ($char -lt 'A' -or $char -gt 'Z') { continue }
        $value = ($value * 26) + ([int][char]$char - [int][char]'A' + 1)
    }
    return $value
}

function Get-ColLetters {
    param([int]$Index)

    $result = ''
    while ($Index -gt 0) {
        $Index--
        $result = [char](65 + ($Index % 26)) + $result
        $Index = [math]::Floor($Index / 26)
    }
    return $result
}

function Read-ZipText {
    param($Zip, [string]$EntryName)

    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) { return $null }

    $reader = [IO.StreamReader]::new($entry.Open())
    try {
        return $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
    }
}

function Read-ZipXml {
    param($Zip, [string]$EntryName)

    $text = Read-ZipText -Zip $Zip -EntryName $EntryName
    if ($null -eq $text) { return $null }
    return [xml]$text
}

function Get-SheetMap {
    param($Zip)

    $workbook = Read-ZipXml -Zip $Zip -EntryName 'xl/workbook.xml'
    $rels = Read-ZipXml -Zip $Zip -EntryName 'xl/_rels/workbook.xml.rels'
    if ($null -eq $workbook -or $null -eq $rels) { throw 'Workbook OpenXML incompleto.' }

    $wbNs = [Xml.XmlNamespaceManager]::new($workbook.NameTable)
    $wbNs.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $wbNs.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

    $relNs = [Xml.XmlNamespaceManager]::new($rels.NameTable)
    $relNs.AddNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships')
    $relMap = @{}
    foreach ($rel in $rels.SelectNodes('//rel:Relationship', $relNs)) {
        $target = $rel.Target
        if ($target -notlike 'xl/*') { $target = 'xl/' + $target.TrimStart('/') }
        $relMap[$rel.Id] = $target
    }

    $map = @{}
    foreach ($sheet in $workbook.SelectNodes('//x:sheet', $wbNs)) {
        $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $map[$sheet.name] = $relMap[$rid]
    }
    return $map
}

function Get-SharedStrings {
    param($Zip)

    $xml = Read-ZipXml -Zip $Zip -EntryName 'xl/sharedStrings.xml'
    $strings = New-Object System.Collections.Generic.List[string]
    if ($null -eq $xml) { return $strings }

    $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
    $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    foreach ($si in $xml.SelectNodes('//x:si', $ns)) {
        $parts = @()
        foreach ($t in $si.SelectNodes('.//x:t', $ns)) { $parts += $t.InnerText }
        $strings.Add(($parts -join ''))
    }
    return $strings
}

function Get-Cell {
    param([xml]$WorksheetXml, [string]$Address)

    $ns = [Xml.XmlNamespaceManager]::new($WorksheetXml.NameTable)
    $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    return $WorksheetXml.SelectSingleNode("//x:c[@r='$Address']", $ns)
}

function Get-CellValue {
    param($Cell, $SharedStrings)

    if ($null -eq $Cell) { return $null }

    $valueNode = $Cell.ChildNodes | Where-Object { $_.LocalName -eq 'v' } | Select-Object -First 1
    if ($Cell.t -eq 's' -and $null -ne $valueNode) {
        $index = [int]$valueNode.InnerText
        if ($index -lt $SharedStrings.Count) { return $SharedStrings[$index] }
    }
    if ($Cell.t -eq 'inlineStr') {
        $textNodes = $Cell.SelectNodes('.//*[local-name()="t"]')
        if ($null -ne $textNodes -and $textNodes.Count -gt 0) {
            return (($textNodes | ForEach-Object { $_.InnerText }) -join '')
        }
        return $Cell.InnerText
    }
    if ($null -ne $valueNode) { return $valueNode.InnerText }
    return $null
}

function Convert-HeaderDate {
    param($Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }

    $number = 0.0
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return [DateTime]::FromOADate($number)
    }

    $date = [DateTime]::MinValue
    foreach ($culture in @([Globalization.CultureInfo]::GetCultureInfo('es-ES'), [Globalization.CultureInfo]::InvariantCulture)) {
        if ([DateTime]::TryParse([string]$Value, $culture, [Globalization.DateTimeStyles]::None, [ref]$date)) {
            return $date
        }
    }
    return $null
}

function Get-CellNumber {
    param($Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return 0.0 }

    $number = 0.0
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::GetCultureInfo('es-ES'), [ref]$number)) {
        return $number
    }
    return [double]::NaN
}

function Assert-MonthHeaders {
    param(
        [xml]$WorksheetXml,
        $SharedStrings,
        [string]$SheetName,
        [int]$HeaderRow,
        [string]$StartColumn,
        [datetime]$ExpectedStart,
        [int]$Months
    )

    $startIndex = Get-ColIndex $StartColumn
    for ($offset = 0; $offset -lt $Months; $offset++) {
        $address = (Get-ColLetters ($startIndex + $offset)) + $HeaderRow
        $value = Get-CellValue -Cell (Get-Cell -WorksheetXml $WorksheetXml -Address $address) -SharedStrings $SharedStrings
        $actual = Convert-HeaderDate -Value $value
        $expected = $ExpectedStart.AddMonths($offset)
        if ($null -eq $actual -or $actual.Year -ne $expected.Year -or $actual.Month -ne $expected.Month) {
            Add-Failure "$SheetName $address esperado $($expected.ToString('yyyy-MM')) actual '$value'"
        }
    }
}

function Assert-WorkbookNoRef {
    param($Zip, $SheetMap)

    foreach ($sheetName in $SheetMap.Keys) {
        $text = Read-ZipText -Zip $Zip -EntryName $SheetMap[$sheetName]
        if ($text -match '#REF!') {
            Add-Failure "Workbook contiene #REF! en hoja $sheetName"
        }
    }
}

function Assert-CellText {
    param(
        [xml]$Sheet,
        $SharedStrings,
        [string]$Address,
        [string]$Expected,
        [string]$Case
    )

    $actual = [string](Get-CellValue -Cell (Get-Cell -WorksheetXml $Sheet -Address $Address) -SharedStrings $SharedStrings)
    if ($actual -ne $Expected) {
        Add-Failure "$Case $Address esperado='$Expected' actual='$actual'"
    }
}

function Assert-CellContains {
    param(
        [xml]$Sheet,
        $SharedStrings,
        [string]$Address,
        [string]$ExpectedPart,
        [string]$Case
    )

    $actual = Get-CellValue -Cell (Get-Cell -WorksheetXml $Sheet -Address $Address) -SharedStrings $SharedStrings
    if ($null -eq $actual -or -not $actual.ToString().Contains($ExpectedPart)) {
        Add-Failure "$Case $Address debe contener '$ExpectedPart'. Actual='$actual'"
    }
}

function Assert-CellNear {
    param(
        [xml]$Sheet,
        $SharedStrings,
        [string]$Address,
        [double]$Expected,
        [double]$Tolerance,
        [string]$Case
    )

    $actual = Get-CellValue -Cell (Get-Cell -WorksheetXml $Sheet -Address $Address) -SharedStrings $SharedStrings
    Assert-NearNumber -Actual $actual -Expected $Expected -Tolerance $Tolerance -Case "$Case $Address"
}

function Assert-CellGreater {
    param(
        [xml]$Sheet,
        $SharedStrings,
        [string]$Address,
        [double]$Minimum,
        [string]$Case
    )

    $actual = Get-CellValue -Cell (Get-Cell -WorksheetXml $Sheet -Address $Address) -SharedStrings $SharedStrings
    $number = Get-CellNumber -Value $actual
    if ([double]::IsNaN($number) -or $number -le $Minimum) {
        Add-Failure "$Case $Address debe ser mayor que $Minimum. Actual='$actual'"
    }
}

function Assert-RangeBlankOrZero {
    param(
        [xml]$Sheet,
        $SharedStrings,
        [string]$FromColumn,
        [string]$ToColumn,
        [int]$FromRow,
        [int]$ToRow,
        [string]$Case
    )

    for ($row = $FromRow; $row -le $ToRow; $row++) {
        for ($col = (Get-ColIndex $FromColumn); $col -le (Get-ColIndex $ToColumn); $col++) {
            $address = (Get-ColLetters $col) + $row
            $value = Get-CellValue -Cell (Get-Cell -WorksheetXml $Sheet -Address $address) -SharedStrings $SharedStrings
            if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { continue }

            $number = Get-CellNumber -Value $value
            if ([double]::IsNaN($number) -or [math]::Abs($number) -gt 0.0000001) {
                Add-Failure "$Case debe estar vacio o cero. $address='$value'"
            }
        }
    }
}

function Test-StaticContract {
    Write-Host '== Contrato estatico E06'

    $xamlFiles = @()
    $xamlFiles += Get-ChildItem -LiteralPath $ProjectRoot -Filter '*.xaml' -File
    $xamlFiles += Get-ChildItem -LiteralPath (Resolve-ProjectPath 'lib') -Filter '*.xaml' -File
    $xamlFiles += Get-ChildItem -LiteralPath (Resolve-ProjectPath 'test') -Filter '*.xaml' -File

    foreach ($file in $xamlFiles | Sort-Object FullName) {
        [xml]$null = Get-Content -Raw -LiteralPath $file.FullName
    }
    Write-Host ("- XML OK: {0} XAML" -f $xamlFiles.Count)

    foreach ($file in $xamlFiles) {
        $text = Get-Content -Raw -LiteralPath $file.FullName
        foreach ($match in [regex]::Matches($text, 'WorkflowFileName="(?<Path>[^"]+)"')) {
            $workflowPath = $match.Groups['Path'].Value
            if ($workflowPath.Contains('[')) { continue }
            $candidate = Resolve-ProjectPath $workflowPath
            if (-not (Test-Path -LiteralPath $candidate)) {
                Add-Failure "WorkflowFileName no resuelve desde $($file.Name): $workflowPath"
            }
        }
    }

    $costPlanningPath = Resolve-ProjectPath 'lib/core-common-preparar-cost-planning.xaml'
    $ppoPath = Resolve-ProjectPath 'lib/core-common-construir-modelo-ppo.xaml'
    $resourcesPath = Resolve-ProjectPath 'lib/core-common-preparar-resources.xaml'
    $corePath = Resolve-ProjectPath 'lib/oportunidad-generar-core.xaml'
    $pcPath = Resolve-ProjectPath 'lib/oportunidad-generar-core-pc.xaml'
    $atPath = Resolve-ProjectPath 'lib/oportunidad-generar-core-at.xaml'
    $cp = [System.Net.WebUtility]::HtmlDecode((Get-Content -Raw -LiteralPath $costPlanningPath))
    $ppo = [System.Net.WebUtility]::HtmlDecode((Get-Content -Raw -LiteralPath $ppoPath))
    $resources = [System.Net.WebUtility]::HtmlDecode((Get-Content -Raw -LiteralPath $resourcesPath))
    $core = [System.Net.WebUtility]::HtmlDecode((Get-Content -Raw -LiteralPath $corePath))
    $pc = [System.Net.WebUtility]::HtmlDecode((Get-Content -Raw -LiteralPath $pcPath))
    $at = [System.Net.WebUtility]::HtmlDecode((Get-Content -Raw -LiteralPath $atPath))

    Assert-TextContains -Text $ppo -Pattern 'dtHorasPorAnio = Tabla("dtHorasPorAnio", "Sigla", "Anio", "Horas", "ResourceKey")' -Case 'dtHorasPorAnio con ResourceKey'
    Assert-TextContains -Text $ppo -Pattern 'LeerLineasBloqueImportes("Compras", filaCompras, dtCompras)' -Case 'Compras desde bloque semantico'
    Assert-TextContains -Text $ppo -Pattern 'filaRiesgos = BuscarFilaPresupuesto("Riesgos", "Riesgos Mano de Obra")' -Case 'Riesgos separados de compras'
    Assert-TextContains -Text $ppo -Pattern 'filaGarantia = BuscarFilaPresupuesto("Garantia", "Garantia Mano de Obra")' -Case 'Garantia separada de compras'
    Assert-TextContains -Text $ppo -Pattern 'factorTarifaAT = pedidoAT / ventaPlanificadaAT' -Case 'AT reconcilia tarifas con pedido'
    Assert-TextNotContains -Text $ppo -Pattern 'layoutAntiguo' -Case 'Sin rama layout antiguo'
    Assert-TextNotContains -Text $ppo -Pattern 'layoutNuevo' -Case 'Sin rama layout nuevo'

    Assert-TextContains -Text $resources -Pattern 'PickCost(costesPorKey, key, anioMes, out costeMes)' -Case 'Resources aplica coste por anualidad del mes'
    Assert-TextContains -Text $core -Pattern 'Fecha reporting prerrelleno' -Case 'CORE no marca realizado sin SAP real'
    Assert-TextContains -Text $core -Pattern 'IF(D8=0,0,D11/D8)' -Case 'Cost Summary avance realizado cero'

    Assert-TextContains -Text $cp -Pattern 'ResourceKey' -Case 'Cost Planning consume ResourceKey'
    Assert-TextContains -Text $cp -Pattern 'horasPorClaveCostPlanning' -Case 'Horas agrupadas por clave robusta'
    Assert-TextContains -Text $cp -Pattern 'conteoPorAnioCostPlanning' -Case 'Prorrateo por anualidad real'
    Assert-TextContains -Text $cp -Pattern 'duracionCostPlanning' -Case 'Duracion dinamica Cost Planning'
    Assert-TextContains -Text $cp -Pattern 'indiceUltimoMesCostPlanning' -Case 'Riesgos/garantia al ultimo mes'
    Assert-TextContains -Text $cp -Pattern 'Proveedor no disponible' -Case 'Texto rojo proveedor acotado a compras'
    Assert-TextContains -Text $cp -Pattern 'PO compra no disponible' -Case 'Texto rojo PO acotado a compras'
    Assert-TextContains -Text $cp -Pattern 'Ariba no disponible' -Case 'Texto rojo Ariba acotado a compras'
    Assert-TextContains -Text $cp -Pattern 'New Object() {out_dt_CostPlanningATRiesgosGarantia, duracionCostPlanning, 2}' -Case 'Riesgos AT con duracion dinamica'
    Assert-TextContains -Text $cp -Pattern 'New Object() {out_dt_CostPlanningPCRiesgosGarantia, duracionCostPlanning, 2}' -Case 'Riesgos PC con duracion dinamica'
    Assert-TextNotContains -Text $cp -Pattern 'New Object() {out_dt_CostPlanningATRiesgosGarantia, 25' -Case 'Sin truncado AT a 25 meses'
    Assert-TextNotContains -Text $cp -Pattern 'New Object() {out_dt_CostPlanningPCRiesgosGarantia, 25' -Case 'Sin truncado PC a 25 meses'

    Assert-TextContains -Text $pc -Pattern 'columnaFinCostPlanningMeses' -Case 'PC escribe Cost Planning hasta columna calculada'
    Assert-TextContains -Text $at -Pattern 'columnaFinCostPlanningMeses' -Case 'AT escribe Cost Planning hasta columna calculada'
    Assert-TextContains -Text $at -Pattern 'Range("H67:" + columnaFinCostPlanningMeses' -Case 'AT limpia meses compras solo hasta duracion'
    Assert-TextContains -Text $pc -Pattern 'Range("G64:" + columnaFinCostPlanningMeses' -Case 'PC limpia meses compras solo hasta duracion'
}

function Test-CorePcWorkbook {
    param([string]$Path)

    Write-Host "== CORE PC E06: $Path"

    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $sheetMap = Get-SheetMap -Zip $zip
        foreach ($required in @('Project Infor', 'Resources', 'Cost Planning', 'Monthly View', 'Cost Summary')) {
            if (-not $sheetMap.ContainsKey($required)) { Add-Failure "CORE PC sin hoja obligatoria: $required" }
        }
        if ($sheetMap.ContainsKey('Trazabilidad RPA')) { Add-Failure 'CORE PC no debe crear hojas nuevas: Trazabilidad RPA' }
        Assert-WorkbookNoRef -Zip $zip -SheetMap $sheetMap

        $sharedStrings = Get-SharedStrings -Zip $zip
        $projectInfor = Read-ZipXml -Zip $zip -EntryName $sheetMap['Project Infor']
        $resources = Read-ZipXml -Zip $zip -EntryName $sheetMap['Resources']
        $costPlanning = Read-ZipXml -Zip $zip -EntryName $sheetMap['Cost Planning']
        $monthlyView = Read-ZipXml -Zip $zip -EntryName $sheetMap['Monthly View']
        $costSummary = Read-ZipXml -Zip $zip -EntryName $sheetMap['Cost Summary']

        Assert-MonthHeaders -WorksheetXml $costPlanning -SharedStrings $sharedStrings -SheetName 'Cost Planning PC' -HeaderRow 8 -StartColumn 'G' -ExpectedStart ([datetime]'2025-11-01') -Months 12
        Assert-CellText -Sheet $costPlanning -SharedStrings $sharedStrings -Address 'G7' -Expected 'Prev' -Case 'PC Cost Planning sin meses Real iniciales'
        Assert-CellText -Sheet $monthlyView -SharedStrings $sharedStrings -Address 'E4' -Expected 'Prev' -Case 'PC Monthly View sin meses Real iniciales'
        Assert-CellNear -Sheet $projectInfor -SharedStrings $sharedStrings -Address 'C28' -Expected 583896.55 -Tolerance 0.01 -Case 'PC importe pedido'
        Assert-CellNear -Sheet $projectInfor -SharedStrings $sharedStrings -Address 'C29' -Expected 13570 -Tolerance 0.01 -Case 'PC horas'
        Assert-CellNear -Sheet $projectInfor -SharedStrings $sharedStrings -Address 'C30' -Expected 367515.14 -Tolerance 0.02 -Case 'PC coste recursos'
        Assert-CellNear -Sheet $projectInfor -SharedStrings $sharedStrings -Address 'C31' -Expected 18375.76 -Tolerance 0.02 -Case 'PC riesgos'
        Assert-CellNear -Sheet $projectInfor -SharedStrings $sharedStrings -Address 'C32' -Expected 0 -Tolerance 0.01 -Case 'PC garantia'
        Assert-CellNear -Sheet $projectInfor -SharedStrings $sharedStrings -Address 'C33' -Expected 7400 -Tolerance 0.01 -Case 'PC gastos'
        Assert-CellNear -Sheet $projectInfor -SharedStrings $sharedStrings -Address 'C34' -Expected 40000 -Tolerance 0.01 -Case 'PC compras'

        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'F7' -Expected 40 -Tolerance 0.01 -Case 'PC coste hora JP'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'F8' -Expected 26 -Tolerance 0.01 -Case 'PC coste hora CSS'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'F9' -Expected 16 -Tolerance 0.01 -Case 'PC coste hora CJSS'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'F10' -Expected 38.75 -Tolerance 0.01 -Case 'PC coste hora A'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'F11' -Expected 27 -Tolerance 0.01 -Case 'PC coste hora II'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'F12' -Expected 20 -Tolerance 0.01 -Case 'PC coste hora P'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'F13' -Expected 25 -Tolerance 0.01 -Case 'PC coste hora DG'
        Assert-CellGreater -Sheet $costSummary -SharedStrings $sharedStrings -Address 'E8' -Minimum 300000 -Case 'PC proyeccion coste recursos'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'D11' -Expected 0 -Tolerance 0.01 -Case 'PC realizado venta vacio'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'E11' -Expected 0 -Tolerance 0.01 -Case 'PC realizado coste recursos vacio'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'F11' -Expected 0 -Tolerance 0.01 -Case 'PC realizado horas vacio'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'D12' -Expected 0 -Tolerance 0.0001 -Case 'PC avance venta sin realizado'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'E12' -Expected 0 -Tolerance 0.0001 -Case 'PC avance coste sin realizado'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'F12' -Expected 0 -Tolerance 0.0001 -Case 'PC avance horas sin realizado'
        Assert-CellContains -Sheet $costSummary -SharedStrings $sharedStrings -Address 'B18' -ExpectedPart 'Trazabilidad: PPO=baseline' -Case 'PC trazabilidad estructurada en comentario normalizado'
        Assert-CellContains -Sheet $costSummary -SharedStrings $sharedStrings -Address 'B18' -ExpectedPart 'ADVERTENCIA: PPO contiene anualidades de horas fuera del periodo CORE planificado.' -Case 'PC trazabilidad anualidades fuera de periodo'

        Assert-CellText -Sheet $costPlanning -SharedStrings $sharedStrings -Address 'B64' -Expected 'Proveedor Test' -Case 'PC compra proveedor'
        Assert-CellText -Sheet $costPlanning -SharedStrings $sharedStrings -Address 'C64' -Expected 'Trabajos on site' -Case 'PC compra descripcion'
        Assert-CellNear -Sheet $costPlanning -SharedStrings $sharedStrings -Address 'R47' -Expected 18375.76 -Tolerance 0.02 -Case 'PC riesgos ultimo mes'
    }
    finally {
        $zip.Dispose()
    }
}

function Test-CoreAtWorkbook {
    param([string]$Path)

    Write-Host "== CORE AT E06: $Path"

    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $sheetMap = Get-SheetMap -Zip $zip
        foreach ($required in @('Project Infor', 'Resources', 'Cost Planning', 'Monthly View', 'Cost Summary')) {
            if (-not $sheetMap.ContainsKey($required)) { Add-Failure "CORE AT sin hoja obligatoria: $required" }
        }
        if ($sheetMap.ContainsKey('Trazabilidad RPA')) { Add-Failure 'CORE AT no debe crear hojas nuevas: Trazabilidad RPA' }
        Assert-WorkbookNoRef -Zip $zip -SheetMap $sheetMap

        $sharedStrings = Get-SharedStrings -Zip $zip
        $projectInfor = Read-ZipXml -Zip $zip -EntryName $sheetMap['Project Infor']
        $resources = Read-ZipXml -Zip $zip -EntryName $sheetMap['Resources']
        $costPlanning = Read-ZipXml -Zip $zip -EntryName $sheetMap['Cost Planning']
        $monthlyView = Read-ZipXml -Zip $zip -EntryName $sheetMap['Monthly View']
        $costSummary = Read-ZipXml -Zip $zip -EntryName $sheetMap['Cost Summary']

        Assert-MonthHeaders -WorksheetXml $resources -SharedStrings $sharedStrings -SheetName 'Resources AT' -HeaderRow 6 -StartColumn 'G' -ExpectedStart ([datetime]'2026-01-01') -Months 48
        Assert-MonthHeaders -WorksheetXml $costPlanning -SharedStrings $sharedStrings -SheetName 'Cost Planning AT' -HeaderRow 8 -StartColumn 'H' -ExpectedStart ([datetime]'2026-01-01') -Months 48
        Assert-MonthHeaders -WorksheetXml $monthlyView -SharedStrings $sharedStrings -SheetName 'Monthly View AT' -HeaderRow 5 -StartColumn 'E' -ExpectedStart ([datetime]'2026-01-01') -Months 48
        Assert-CellText -Sheet $costPlanning -SharedStrings $sharedStrings -Address 'H7' -Expected 'Prev' -Case 'AT Cost Planning sin meses Real iniciales'
        Assert-CellText -Sheet $monthlyView -SharedStrings $sharedStrings -Address 'E4' -Expected 'Prev' -Case 'AT Monthly View sin meses Real iniciales'

        Assert-CellNear -Sheet $projectInfor -SharedStrings $sharedStrings -Address 'C28' -Expected 465000 -Tolerance 0.01 -Case 'AT importe pedido semantico'
        Assert-CellNear -Sheet $projectInfor -SharedStrings $sharedStrings -Address 'C33' -Expected 0 -Tolerance 0.01 -Case 'AT gastos sin datos reales'
        Assert-CellNear -Sheet $projectInfor -SharedStrings $sharedStrings -Address 'C34' -Expected 0 -Tolerance 0.01 -Case 'AT compras sin datos reales'

        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'G7' -Expected 32.84 -Tolerance 0.01 -Case 'AT coste hora Susana'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'G9' -Expected 15.17 -Tolerance 0.01 -Case 'AT coste hora Jorge'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'G11' -Expected 20 -Tolerance 0.01 -Case 'AT coste hora soporte'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'S7' -Expected 33.83 -Tolerance 0.02 -Case 'AT coste hora Susana 2027'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'AE7' -Expected 34.84 -Tolerance 0.02 -Case 'AT coste hora Susana 2028'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'AQ7' -Expected 35.89 -Tolerance 0.02 -Case 'AT coste hora Susana 2029'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'G8' -Expected 44.79 -Tolerance 0.02 -Case 'AT tarifa venta Susana reconciliada'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'G10' -Expected 20.69 -Tolerance 0.02 -Case 'AT tarifa venta Jorge reconciliada'
        Assert-CellNear -Sheet $resources -SharedStrings $sharedStrings -Address 'G12' -Expected 27.28 -Tolerance 0.02 -Case 'AT tarifa venta soporte reconciliada'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'E8' -Expected 356584.84 -Tolerance 0.05 -Case 'AT proyeccion coste recursos cuadra baseline'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'J8' -Expected 0 -Tolerance 0.01 -Case 'AT proyeccion compras sin compras PPO'
        Assert-CellText -Sheet $costSummary -SharedStrings $sharedStrings -Address 'B14' -Expected '' -Case 'AT warning venta no contradictorio'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'D11' -Expected 0 -Tolerance 0.01 -Case 'AT realizado venta vacio'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'E11' -Expected 0 -Tolerance 0.01 -Case 'AT realizado coste recursos vacio'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'F11' -Expected 0 -Tolerance 0.01 -Case 'AT realizado horas vacio'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'D12' -Expected 0 -Tolerance 0.0001 -Case 'AT avance venta sin realizado'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'E12' -Expected 0 -Tolerance 0.0001 -Case 'AT avance coste sin realizado'
        Assert-CellNear -Sheet $costSummary -SharedStrings $sharedStrings -Address 'F12' -Expected 0 -Tolerance 0.0001 -Case 'AT avance horas sin realizado'
        Assert-CellContains -Sheet $costSummary -SharedStrings $sharedStrings -Address 'B18' -ExpectedPart 'Trazabilidad: PPO=baseline' -Case 'AT trazabilidad estructurada en comentario normalizado'
        Assert-CellNear -Sheet $costPlanning -SharedStrings $sharedStrings -Address 'H44' -Expected 9687.5 -Tolerance 0.02 -Case 'AT venta facturable mensual reconciliada'

        $expectedRows = @(
            @{ Row = 10; Sigla = 'AP'; Perfil = 'Susana Matarranz'; Tipo = 'Horas Reales' },
            @{ Row = 11; Sigla = 'AP'; Perfil = 'Susana Matarranz'; Tipo = 'Horas/Jornadas Facturables' },
            @{ Row = 12; Sigla = 'PR'; Perfil = 'Jorge Puertas'; Tipo = 'Horas Reales' },
            @{ Row = 13; Sigla = 'PR'; Perfil = 'Jorge Puertas'; Tipo = 'Horas/Jornadas Facturables' },
            @{ Row = 14; Sigla = 'AP'; Perfil = 'Sopore especializado'; Tipo = 'Horas Reales' },
            @{ Row = 15; Sigla = 'AP'; Perfil = 'Sopore especializado'; Tipo = 'Horas/Jornadas Facturables' }
        )

        foreach ($row in $expectedRows) {
            Assert-CellText -Sheet $costPlanning -SharedStrings $sharedStrings -Address ("C{0}" -f $row.Row) -Expected $row.Sigla -Case 'AT sigla Cost Planning'
            Assert-CellText -Sheet $costPlanning -SharedStrings $sharedStrings -Address ("E{0}" -f $row.Row) -Expected $row.Perfil -Case 'AT perfil Cost Planning'
            Assert-CellText -Sheet $costPlanning -SharedStrings $sharedStrings -Address ("G{0}" -f $row.Row) -Expected $row.Tipo -Case 'AT tipo horas Cost Planning'
        }

        foreach ($pair in @(@(10, 11), @(12, 13), @(14, 15))) {
            for ($col = (Get-ColIndex 'H'); $col -le (Get-ColIndex 'BC'); $col++) {
                $addressReal = (Get-ColLetters $col) + $pair[0]
                $addressBillable = (Get-ColLetters $col) + $pair[1]
                $real = Get-CellValue -Cell (Get-Cell -WorksheetXml $costPlanning -Address $addressReal) -SharedStrings $sharedStrings
                $billable = Get-CellValue -Cell (Get-Cell -WorksheetXml $costPlanning -Address $addressBillable) -SharedStrings $sharedStrings
                if ([string]$real -ne [string]$billable) {
                    Add-Failure "AT horas reales/facturables no replicadas para revision JP: $addressReal='$real' $addressBillable='$billable'"
                }
            }
        }

        Assert-RangeBlankOrZero -Sheet $costPlanning -SharedStrings $sharedStrings -FromColumn 'H' -ToColumn 'BB' -FromRow 47 -ToRow 48 -Case 'AT riesgos/garantia antes del ultimo mes'
        Assert-CellNear -Sheet $costPlanning -SharedStrings $sharedStrings -Address 'BC47' -Expected 3565.85 -Tolerance 0.02 -Case 'AT riesgos ultimo mes'
        Assert-CellNear -Sheet $costPlanning -SharedStrings $sharedStrings -Address 'BC48' -Expected 0 -Tolerance 0.01 -Case 'AT garantia ultimo mes'

        Assert-RangeBlankOrZero -Sheet $costPlanning -SharedStrings $sharedStrings -FromColumn 'B' -ToColumn 'G' -FromRow 67 -ToRow 76 -Case 'AT bloque fijo compras sin compras PPO'
        Assert-RangeBlankOrZero -Sheet $costPlanning -SharedStrings $sharedStrings -FromColumn 'H' -ToColumn 'BC' -FromRow 67 -ToRow 76 -Case 'AT meses compras sin compras PPO'
    }
    finally {
        $zip.Dispose()
    }
}

Push-Location -LiteralPath $ProjectRoot
try {
    Test-StaticContract

    if (-not $SkipWorkbooks) {
        if ([string]::IsNullOrWhiteSpace($PcWorkbookPath)) {
            $PcWorkbookPath = Resolve-FirstExisting @(
                'data/output/CORE_PC_20250256445.xlsx',
                'data/output/CORE_PC_20250256445_test.xlsx'
            )
        }
        elseif (-not [IO.Path]::IsPathRooted($PcWorkbookPath)) {
            $PcWorkbookPath = Resolve-ProjectPath $PcWorkbookPath
        }

        if ([string]::IsNullOrWhiteSpace($AtWorkbookPath)) {
            $AtWorkbookPath = Resolve-FirstExisting @(
                'data/output/CORE_AT_20251160543.xlsx',
                'data/output/CORE_AT_20251160543_test.xlsx'
            )
        }
        elseif (-not [IO.Path]::IsPathRooted($AtWorkbookPath)) {
            $AtWorkbookPath = Resolve-ProjectPath $AtWorkbookPath
        }

        if ([string]::IsNullOrWhiteSpace($PcWorkbookPath) -or -not (Test-Path -LiteralPath $PcWorkbookPath)) {
            Add-Failure 'No existe output CORE PC para validar E06. Ejecuta Main.xaml o test/test_generar_core_pc.xaml.'
        }
        else {
            Test-CorePcWorkbook -Path (Resolve-Path -LiteralPath $PcWorkbookPath).Path
        }

        if ([string]::IsNullOrWhiteSpace($AtWorkbookPath) -or -not (Test-Path -LiteralPath $AtWorkbookPath)) {
            Add-Failure 'No existe output CORE AT para validar E06. Ejecuta Main.xaml o test/test_generar_core_at.xaml.'
        }
        else {
            Test-CoreAtWorkbook -Path (Resolve-Path -LiteralPath $AtWorkbookPath).Path
        }
    }
    else {
        $warnings.Add('Validacion de workbooks omitida por parametro -SkipWorkbooks.')
    }
}
finally {
    Pop-Location
}

if ($warnings.Count -gt 0) {
    Write-Warning 'Advertencias E06:'
    $warnings | ForEach-Object { Write-Warning $_ }
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Validacion E06 no superada. Fallos: $($failures.Count)"
}

Write-Host 'OK - Validacion E06 superada: ResourceKey, prorrateo anual, AT real/facturable, riesgos ultimo mes, 48 meses y ausencia de compras falsas.'
