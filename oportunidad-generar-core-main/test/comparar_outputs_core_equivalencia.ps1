param(
    [Parameter(Mandatory = $true)]
    [string]$BaselinePath,

    [Parameter(Mandatory = $true)]
    [string]$CandidatePath,

    [Parameter(Mandatory = $true)]
    [ValidateSet('PC', 'AT')]
    [string]$Kind,

    [double]$NumericTolerance = 0.000001,

    [int]$MaxDifferences = 200
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$keySheets = @('Project Infor', 'Resources', 'Cost Planning', 'Monthly View', 'Cost Summary')
$reviewRanges = @{
    'Project Infor' = @('A1:J80')
    'Resources' = @('A1:CA120')
    'Cost Planning' = @('A1:CB140')
}
$timestampNormalizedCells = @{
    'Cost Summary' = @('B18')
}

$script:DifferenceCount = 0
$script:Differences = New-Object System.Collections.Generic.List[object]

function Add-Difference {
    param(
        [string]$Sheet,
        [string]$Cell,
        [string]$Type,
        [object]$Baseline,
        [object]$Candidate,
        [string]$Detail
    )

    $script:DifferenceCount++
    if ($script:Differences.Count -lt $MaxDifferences) {
        $script:Differences.Add([pscustomobject]@{
            Hoja = $Sheet
            Celda = $Cell
            Tipo = $Type
            Baseline = Format-ReportValue -Value $Baseline
            Candidate = Format-ReportValue -Value $Candidate
            Detalle = $Detail
        })
    }
}

function Format-ReportValue {
    param([object]$Value)

    if ($null -eq $Value) { return '' }
    $text = ([string]$Value) -replace "`r", ' ' -replace "`n", ' '
    if ($text.Length -gt 160) {
        return $text.Substring(0, 157) + '...'
    }
    return $text
}

function Resolve-WorkbookPath {
    param([string]$Path)

    $candidate = if ([IO.Path]::IsPathRooted($Path)) { $Path } else { Join-Path (Get-Location) $Path }
    if (-not (Test-Path -LiteralPath $candidate)) {
        throw "Workbook no encontrado: $Path"
    }
    return (Resolve-Path -LiteralPath $candidate).Path
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

    $xml = [xml]::new()
    $xml.PreserveWhitespace = $true
    $xml.LoadXml($text)
    return $xml
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

function Split-CellAddress {
    param([string]$Address)

    if ($Address -notmatch '^([A-Z]+)([0-9]+)$') {
        throw "Direccion de celda no valida: $Address"
    }

    return [pscustomobject]@{
        ColumnLetters = $Matches[1]
        Column = Get-ColIndex -Letters $Matches[1]
        Row = [int]$Matches[2]
    }
}

function Get-RangeAddresses {
    param([string]$Range)

    $bounds = Get-RangeBounds -Range $Range
    $addresses = New-Object System.Collections.Generic.List[string]

    for ($row = $bounds.StartRow; $row -le $bounds.EndRow; $row++) {
        for ($col = $bounds.StartColumn; $col -le $bounds.EndColumn; $col++) {
            $addresses.Add((Get-ColLetters -Index $col) + $row)
        }
    }

    return $addresses
}

function Get-RangeBounds {
    param([string]$Range)

    if ($Range -notmatch '^([A-Z]+[0-9]+):([A-Z]+[0-9]+)$') {
        throw "Rango no valido: $Range"
    }

    $start = Split-CellAddress -Address $Matches[1]
    $end = Split-CellAddress -Address $Matches[2]

    return [pscustomobject]@{
        StartColumn = [math]::Min($start.Column, $end.Column)
        EndColumn = [math]::Max($start.Column, $end.Column)
        StartRow = [math]::Min($start.Row, $end.Row)
        EndRow = [math]::Max($start.Row, $end.Row)
    }
}

function Get-RangeCellCount {
    param([string]$Range)

    $bounds = Get-RangeBounds -Range $Range
    return (($bounds.EndColumn - $bounds.StartColumn + 1) * ($bounds.EndRow - $bounds.StartRow + 1))
}

function Test-AddressInsideBounds {
    param([string]$Address, $Bounds)

    $parts = Split-CellAddress -Address $Address
    return ($parts.Column -ge $Bounds.StartColumn -and
        $parts.Column -le $Bounds.EndColumn -and
        $parts.Row -ge $Bounds.StartRow -and
        $parts.Row -le $Bounds.EndRow)
}

function Get-AddressRow {
    param([string]$Address)
    return (Split-CellAddress -Address $Address).Row
}

function Get-AddressColumn {
    param([string]$Address)
    return (Split-CellAddress -Address $Address).Column
}

function Sort-CellAddresses {
    param([string[]]$Addresses)

    $items = foreach ($address in $Addresses) {
        if ([string]::IsNullOrWhiteSpace($address)) { continue }
        $parts = Split-CellAddress -Address $address
        [pscustomobject]@{
            Address = $address
            SortKey = ([int64]$parts.Row * 20000) + $parts.Column
        }
    }

    return $items | Sort-Object SortKey | Select-Object -ExpandProperty Address
}

function Get-AddressUnion {
    param([string[]]$Left, [string[]]$Right)

    $set = @{}
    foreach ($address in $Left) { $set[$address] = $true }
    foreach ($address in $Right) { $set[$address] = $true }
    return Sort-CellAddresses -Addresses @($set.Keys)
}

function Get-SheetMap {
    param($Zip)

    $workbook = Read-ZipXml -Zip $Zip -EntryName 'xl/workbook.xml'
    $rels = Read-ZipXml -Zip $Zip -EntryName 'xl/_rels/workbook.xml.rels'
    if ($null -eq $workbook -or $null -eq $rels) {
        throw 'Workbook OpenXML incompleto: faltan workbook.xml o workbook.xml.rels.'
    }

    $wbNs = [Xml.XmlNamespaceManager]::new($workbook.NameTable)
    $wbNs.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $wbNs.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

    $relNs = [Xml.XmlNamespaceManager]::new($rels.NameTable)
    $relNs.AddNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships')

    $relMap = @{}
    foreach ($rel in $rels.SelectNodes('//rel:Relationship', $relNs)) {
        $target = [string]$rel.Target
        if ($target.StartsWith('/')) {
            $target = $target.TrimStart('/')
        }
        elseif ($target -notlike 'xl/*') {
            $target = 'xl/' + $target
        }
        $relMap[[string]$rel.Id] = $target
    }

    $map = @{}
    foreach ($sheet in $workbook.SelectNodes('//x:sheet', $wbNs)) {
        $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        if (-not $relMap.ContainsKey($rid)) {
            throw "Relationship no encontrada para hoja '$($sheet.name)': $rid"
        }
        $map[[string]$sheet.name] = $relMap[$rid]
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
        foreach ($t in $si.SelectNodes('.//x:t', $ns)) {
            $parts += $t.InnerText
        }
        $strings.Add(($parts -join ''))
    }

    return $strings
}

function Test-RgbIsRed {
    param([string]$Rgb)

    if ([string]::IsNullOrWhiteSpace($Rgb)) { return $false }

    $hex = $Rgb.Trim().TrimStart('#')
    if ($hex.Length -eq 8) {
        $hex = $hex.Substring(2)
    }
    elseif ($hex.Length -gt 6) {
        $hex = $hex.Substring($hex.Length - 6)
    }

    if ($hex -notmatch '^[0-9A-Fa-f]{6}$') { return $false }

    $red = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $green = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $blue = [Convert]::ToInt32($hex.Substring(4, 2), 16)

    return ($red -ge 180 -and $green -le 80 -and $blue -le 80)
}

function Test-IndexedColorIsRed {
    param([string]$Indexed)

    if ([string]::IsNullOrWhiteSpace($Indexed)) { return $false }

    $index = 0
    if (-not [int]::TryParse($Indexed, [ref]$index)) { return $false }
    return $index -in @(3, 10)
}

function Get-RedStyleMap {
    param($Zip)

    $styles = Read-ZipXml -Zip $Zip -EntryName 'xl/styles.xml'
    $styleRed = @{}
    $styleRed[0] = $false
    if ($null -eq $styles) { return $styleRed }

    $ns = [Xml.XmlNamespaceManager]::new($styles.NameTable)
    $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    $fontRed = @{}
    $fontIndex = 0
    foreach ($font in $styles.SelectNodes('//x:fonts/x:font', $ns)) {
        $isRed = $false
        $color = $font.SelectSingleNode('x:color', $ns)
        if ($null -ne $color) {
            $isRed = (Test-RgbIsRed -Rgb $color.GetAttribute('rgb')) -or (Test-IndexedColorIsRed -Indexed $color.GetAttribute('indexed'))
        }
        $fontRed[$fontIndex] = $isRed
        $fontIndex++
    }

    $styleIndex = 0
    foreach ($xf in $styles.SelectNodes('//x:cellXfs/x:xf', $ns)) {
        $fontId = 0
        $fontIdText = $xf.GetAttribute('fontId')
        if (-not [string]::IsNullOrWhiteSpace($fontIdText)) {
            [void][int]::TryParse($fontIdText, [ref]$fontId)
        }
        $styleRed[$styleIndex] = ($fontRed.ContainsKey($fontId) -and $fontRed[$fontId])
        $styleIndex++
    }

    return $styleRed
}

function Open-CoreWorkbook {
    param([string]$Path)

    $resolvedPath = Resolve-WorkbookPath -Path $Path
    $zip = [IO.Compression.ZipFile]::OpenRead($resolvedPath)

    try {
        return [pscustomobject]@{
            Path = $resolvedPath
            Zip = $zip
            SheetMap = Get-SheetMap -Zip $zip
            SharedStrings = Get-SharedStrings -Zip $zip
            RedStyleMap = Get-RedStyleMap -Zip $zip
            SheetXmlCache = @{}
            CellMapCache = @{}
        }
    }
    catch {
        $zip.Dispose()
        throw
    }
}

function Get-SheetXml {
    param($Workbook, [string]$SheetName)

    if (-not $Workbook.SheetXmlCache.ContainsKey($SheetName)) {
        if (-not $Workbook.SheetMap.ContainsKey($SheetName)) {
            throw "Hoja no encontrada: $SheetName"
        }
        $Workbook.SheetXmlCache[$SheetName] = Read-ZipXml -Zip $Workbook.Zip -EntryName $Workbook.SheetMap[$SheetName]
    }

    return $Workbook.SheetXmlCache[$SheetName]
}

function Get-CellMap {
    param($Workbook, [string]$SheetName)

    if (-not $Workbook.CellMapCache.ContainsKey($SheetName)) {
        $sheet = Get-SheetXml -Workbook $Workbook -SheetName $SheetName
        $ns = [Xml.XmlNamespaceManager]::new($sheet.NameTable)
        $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

        $map = @{}
        foreach ($cell in $sheet.SelectNodes('//x:c', $ns)) {
            $address = $cell.GetAttribute('r')
            if (-not [string]::IsNullOrWhiteSpace($address)) {
                $map[$address] = $cell
            }
        }
        $Workbook.CellMapCache[$SheetName] = $map
    }

    return $Workbook.CellMapCache[$SheetName]
}

function Get-FirstChildByLocalName {
    param($Node, [string]$LocalName)

    if ($null -eq $Node) { return $null }
    foreach ($child in $Node.ChildNodes) {
        if ($child.LocalName -eq $LocalName) { return $child }
    }
    return $null
}

function Get-CellRawValue {
    param($Cell, $SharedStrings)

    if ($null -eq $Cell) { return '' }

    $type = $Cell.GetAttribute('t')
    $valueNode = Get-FirstChildByLocalName -Node $Cell -LocalName 'v'

    if ($type -eq 's' -and $null -ne $valueNode) {
        $index = 0
        if ([int]::TryParse($valueNode.InnerText, [ref]$index) -and $index -ge 0 -and $index -lt $SharedStrings.Count) {
            return $SharedStrings[$index]
        }
        return $valueNode.InnerText
    }

    if ($type -eq 'inlineStr') {
        $textNodes = $Cell.SelectNodes(".//*[local-name()='t']")
        if ($null -ne $textNodes -and $textNodes.Count -gt 0) {
            return (($textNodes | ForEach-Object { $_.InnerText }) -join '')
        }
        return $Cell.InnerText
    }

    if ($null -ne $valueNode) { return $valueNode.InnerText }
    return ''
}

function Normalize-CoreValue {
    param(
        [string]$SheetName,
        [string]$Address,
        [object]$Value
    )

    $text = if ($null -eq $Value) { '' } else { [string]$Value }

    if ($timestampNormalizedCells.ContainsKey($SheetName) -and $timestampNormalizedCells[$SheetName] -contains $Address) {
        $text = [regex]::Replace($text, 'Generado por RPA: \d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}(\.\d+)?', 'Generado por RPA: <TIMESTAMP_NORMALIZADO>')
    }

    return $text
}

function Get-ComparableCellValue {
    param($Workbook, [string]$SheetName, [string]$Address)

    $cellMap = Get-CellMap -Workbook $Workbook -SheetName $SheetName
    $cell = if ($cellMap.ContainsKey($Address)) { $cellMap[$Address] } else { $null }
    $value = Get-CellRawValue -Cell $cell -SharedStrings $Workbook.SharedStrings
    return Normalize-CoreValue -SheetName $SheetName -Address $Address -Value $value
}

function Get-FormulaSignature {
    param($Cell)

    if ($null -eq $Cell) { return '' }

    $formulaNode = Get-FirstChildByLocalName -Node $Cell -LocalName 'f'
    if ($null -eq $formulaNode) { return '' }

    $type = $formulaNode.GetAttribute('t')
    $text = $formulaNode.InnerText

    if ([string]::IsNullOrWhiteSpace($text) -and $type -eq 'shared') {
        return '[shared-formula]'
    }

    if ([string]::IsNullOrWhiteSpace($type)) {
        return $text
    }

    return ('type={0}|text={1}' -f $type, $text)
}

function Get-CellAddressesWithValues {
    param($Workbook, [string]$SheetName)

    $addresses = New-Object System.Collections.Generic.List[string]
    $cellMap = Get-CellMap -Workbook $Workbook -SheetName $SheetName

    foreach ($address in $cellMap.Keys) {
        $cell = $cellMap[$address]
        $hasValue = $null -ne (Get-FirstChildByLocalName -Node $cell -LocalName 'v')
        $hasInlineString = $cell.GetAttribute('t') -eq 'inlineStr'
        $hasFormula = -not [string]::IsNullOrEmpty((Get-FormulaSignature -Cell $cell))

        if ($hasValue -or $hasInlineString -or $hasFormula) {
            $addresses.Add($address)
        }
    }

    return Sort-CellAddresses -Addresses $addresses.ToArray()
}

function Get-CellAddressesWithFormulas {
    param($Workbook, [string]$SheetName)

    $addresses = New-Object System.Collections.Generic.List[string]
    $cellMap = Get-CellMap -Workbook $Workbook -SheetName $SheetName

    foreach ($address in $cellMap.Keys) {
        if (-not [string]::IsNullOrEmpty((Get-FormulaSignature -Cell $cellMap[$address]))) {
            $addresses.Add($address)
        }
    }

    return Sort-CellAddresses -Addresses $addresses.ToArray()
}

function Test-ValuesEquivalent {
    param([string]$Left, [string]$Right)

    if ($Left -eq $Right) { return $true }

    $leftNumber = 0.0
    $rightNumber = 0.0
    $leftIsNumber = [double]::TryParse($Left, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$leftNumber)
    $rightIsNumber = [double]::TryParse($Right, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$rightNumber)

    if ($leftIsNumber -and $rightIsNumber) {
        return ([math]::Abs($leftNumber - $rightNumber) -le $NumericTolerance)
    }

    return $false
}

function Test-CellFontIsRed {
    param($Workbook, [string]$SheetName, [string]$Address)

    $cellMap = Get-CellMap -Workbook $Workbook -SheetName $SheetName
    if (-not $cellMap.ContainsKey($Address)) { return $false }

    return Test-CellObjectFontIsRed -Workbook $Workbook -Cell $cellMap[$Address]
}

function Test-CellObjectFontIsRed {
    param($Workbook, $Cell)

    if ($null -eq $Cell) { return $false }

    $styleText = $Cell.GetAttribute('s')
    if ([string]::IsNullOrWhiteSpace($styleText)) { return $false }

    $styleId = 0
    if (-not [int]::TryParse($styleText, [ref]$styleId)) { return $false }

    return ($Workbook.RedStyleMap.ContainsKey($styleId) -and $Workbook.RedStyleMap[$styleId])
}

function Get-RedCellSetInRange {
    param($Workbook, [string]$SheetName, [string]$Range)

    $bounds = Get-RangeBounds -Range $Range
    $cellMap = Get-CellMap -Workbook $Workbook -SheetName $SheetName
    $set = @{}

    foreach ($address in $cellMap.Keys) {
        if ((Test-AddressInsideBounds -Address $address -Bounds $bounds) -and (Test-CellObjectFontIsRed -Workbook $Workbook -Cell $cellMap[$address])) {
            $set[$address] = $true
        }
    }

    return $set
}

function Assert-RequiredSheets {
    param($Baseline, $Candidate)

    $missing = $false
    foreach ($sheet in $keySheets) {
        if (-not $Baseline.SheetMap.ContainsKey($sheet)) {
            Add-Difference -Sheet $sheet -Cell '' -Type 'HojaObligatoria' -Baseline 'No existe' -Candidate '' -Detail 'Falta hoja clave en baseline.'
            $missing = $true
        }
        if (-not $Candidate.SheetMap.ContainsKey($sheet)) {
            Add-Difference -Sheet $sheet -Cell '' -Type 'HojaObligatoria' -Baseline '' -Candidate 'No existe' -Detail 'Falta hoja clave en candidato.'
            $missing = $true
        }
    }
    return (-not $missing)
}

function Compare-CellValues {
    param($Baseline, $Candidate)

    $compared = 0

    foreach ($sheet in $keySheets) {
        $addresses = Get-AddressUnion -Left (Get-CellAddressesWithValues -Workbook $Baseline -SheetName $sheet) -Right (Get-CellAddressesWithValues -Workbook $Candidate -SheetName $sheet)
        foreach ($address in $addresses) {
            $baselineValue = Get-ComparableCellValue -Workbook $Baseline -SheetName $sheet -Address $address
            $candidateValue = Get-ComparableCellValue -Workbook $Candidate -SheetName $sheet -Address $address
            $compared++

            if (-not (Test-ValuesEquivalent -Left $baselineValue -Right $candidateValue)) {
                Add-Difference -Sheet $sheet -Cell $address -Type 'Valor' -Baseline $baselineValue -Candidate $candidateValue -Detail "Valor visible/cached distinto. Tolerancia numerica=$NumericTolerance."
            }
        }
    }

    return $compared
}

function Compare-Formulas {
    param($Baseline, $Candidate)

    $compared = 0

    foreach ($sheet in $keySheets) {
        $addresses = Get-AddressUnion -Left (Get-CellAddressesWithFormulas -Workbook $Baseline -SheetName $sheet) -Right (Get-CellAddressesWithFormulas -Workbook $Candidate -SheetName $sheet)
        $baselineCells = Get-CellMap -Workbook $Baseline -SheetName $sheet
        $candidateCells = Get-CellMap -Workbook $Candidate -SheetName $sheet

        foreach ($address in $addresses) {
            $baselineCell = if ($baselineCells.ContainsKey($address)) { $baselineCells[$address] } else { $null }
            $candidateCell = if ($candidateCells.ContainsKey($address)) { $candidateCells[$address] } else { $null }
            $baselineFormula = Get-FormulaSignature -Cell $baselineCell
            $candidateFormula = Get-FormulaSignature -Cell $candidateCell
            $compared++

            if ($baselineFormula -ne $candidateFormula) {
                Add-Difference -Sheet $sheet -Cell $address -Type 'Formula' -Baseline $baselineFormula -Candidate $candidateFormula -Detail 'Formula ausente, fija o distinta.'
            }
        }
    }

    return $compared
}

function Assert-NoRefErrors {
    param($Workbook, [string]$WorkbookRole)

    $count = 0
    foreach ($sheet in $keySheets) {
        $cellMap = Get-CellMap -Workbook $Workbook -SheetName $sheet
        foreach ($address in $cellMap.Keys) {
            $cell = $cellMap[$address]
            $value = Normalize-CoreValue -SheetName $sheet -Address $address -Value (Get-CellRawValue -Cell $cell -SharedStrings $Workbook.SharedStrings)
            $formula = Get-FormulaSignature -Cell $cell
            if ($value.Contains('#REF!') -or $formula.Contains('#REF!')) {
                Add-Difference -Sheet $sheet -Cell $address -Type 'RefError' -Baseline $WorkbookRole -Candidate '#REF!' -Detail "El $WorkbookRole contiene #REF!."
                $count++
            }
        }
    }
    return $count
}

function Compare-RedFontRanges {
    param($Baseline, $Candidate)

    $compared = 0
    foreach ($sheet in $reviewRanges.Keys) {
        foreach ($range in $reviewRanges[$sheet]) {
            $compared += Get-RangeCellCount -Range $range
            $baselineRedCells = Get-RedCellSetInRange -Workbook $Baseline -SheetName $sheet -Range $range
            $candidateRedCells = Get-RedCellSetInRange -Workbook $Candidate -SheetName $sheet -Range $range
            $addresses = Get-AddressUnion -Left @($baselineRedCells.Keys) -Right @($candidateRedCells.Keys)

            foreach ($address in $addresses) {
                $baselineRed = $baselineRedCells.ContainsKey($address)
                $candidateRed = $candidateRedCells.ContainsKey($address)
                if ($baselineRed -ne $candidateRed) {
                    Add-Difference -Sheet $sheet -Cell $address -Type 'FuenteRoja' -Baseline $baselineRed -Candidate $candidateRed -Detail "Estado rojo/no rojo distinto dentro del rango contractual $range."
                }
            }
        }
    }

    return $compared
}

$baselineWorkbook = Open-CoreWorkbook -Path $BaselinePath
$candidateWorkbook = Open-CoreWorkbook -Path $CandidatePath

try {
    $valueCellsCompared = 0
    $formulaCellsCompared = 0
    $refErrorsFound = 0
    $redCellsCompared = 0

    if (Assert-RequiredSheets -Baseline $baselineWorkbook -Candidate $candidateWorkbook) {
        $valueCellsCompared = Compare-CellValues -Baseline $baselineWorkbook -Candidate $candidateWorkbook
        $formulaCellsCompared = Compare-Formulas -Baseline $baselineWorkbook -Candidate $candidateWorkbook
        $refErrorsFound += Assert-NoRefErrors -Workbook $baselineWorkbook -WorkbookRole 'baseline'
        $refErrorsFound += Assert-NoRefErrors -Workbook $candidateWorkbook -WorkbookRole 'candidato'
        $redCellsCompared = Compare-RedFontRanges -Baseline $baselineWorkbook -Candidate $candidateWorkbook
    }

    $summary = [pscustomobject]@{
        Kind = $Kind
        Baseline = $baselineWorkbook.Path
        Candidate = $candidateWorkbook.Path
        ValueCellsCompared = $valueCellsCompared
        FormulaCellsCompared = $formulaCellsCompared
        RefErrorsFound = $refErrorsFound
        RedCellsCompared = $redCellsCompared
        NumericTolerance = $NumericTolerance
        TimestampNormalizer = 'Cost Summary!B18'
        Differences = $script:DifferenceCount
    }

    if ($script:DifferenceCount -gt 0) {
        Write-Host "KO - Comparacion CORE $Kind no superada."
        $summary | Format-List
        Write-Host "Primeras diferencias reportadas: $($script:Differences.Count) de $script:DifferenceCount"
        $script:Differences | Format-Table -AutoSize
        throw "Comparacion CORE $Kind no superada. Diferencias: $script:DifferenceCount"
    }

    $summary | Format-List
    Write-Host "OK - Comparacion CORE $Kind superada: valores, formulas, ausencia de #REF! y fuente roja contractual equivalentes."
}
finally {
    $baselineWorkbook.Zip.Dispose()
    $candidateWorkbook.Zip.Dispose()
}
