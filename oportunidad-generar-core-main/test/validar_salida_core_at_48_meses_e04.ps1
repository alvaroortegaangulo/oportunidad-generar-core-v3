param(
    [Alias('Path')]
    [string]$WorkbookPath = (Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..')).Path 'data/output/CORE_AT_20251160543_test.xlsx'),
    [ValidateSet('AT', 'PC')]
    [string]$Kind = 'AT',
    [datetime]$ExpectedStart = [datetime]'2026-01-01',
    [int]$Months = 48
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipText {
    param($Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) { return $null }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Read-ZipXml {
    param($Zip, [string]$EntryName)
    $text = Read-ZipText -Zip $Zip -EntryName $EntryName
    if ($null -eq $text) { return $null }
    return [xml]$text
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

function Get-SheetMap {
    param($Zip)
    $workbook = Read-ZipXml -Zip $Zip -EntryName 'xl/workbook.xml'
    $rels = Read-ZipXml -Zip $Zip -EntryName 'xl/_rels/workbook.xml.rels'

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

function Test-MonthHeaders {
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
            throw "$SheetName $address expected $($expected.ToString('yyyy-MM')) but found '$value'."
        }
    }

    return [pscustomobject]@{
        Sheet = $SheetName
        Range = "${StartColumn}${HeaderRow}:$((Get-ColLetters ($startIndex + $Months - 1)))${HeaderRow}"
        First = $ExpectedStart.ToString('yyyy-MM')
        Last = $ExpectedStart.AddMonths($Months - 1).ToString('yyyy-MM')
        Months = $Months
    }
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

$zip = [IO.Compression.ZipFile]::OpenRead((Resolve-Path $WorkbookPath))
try {
    $sheetMap = Get-SheetMap -Zip $zip
    foreach ($required in @('Project Infor', 'Resources', 'Cost Planning', 'Monthly View', 'Cost Summary')) {
        if (-not $sheetMap.ContainsKey($required)) { throw "Missing required sheet: $required" }
    }

    $refErrors = 0
    $formulaCount = 0
    foreach ($sheetName in $sheetMap.Keys) {
        $text = Read-ZipText -Zip $zip -EntryName $sheetMap[$sheetName]
        if ($text -match '#REF!') { $refErrors += [regex]::Matches($text, '#REF!').Count }
        $xml = [xml]$text
        $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
        $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $formulaCount += $xml.SelectNodes('//x:f', $ns).Count
    }
    if ($refErrors -ne 0) { throw "Workbook contains #REF!: $refErrors" }
    $minimumFormulaCount = if ($Kind -eq 'AT') { 3500 } else { 3200 }
    if ($formulaCount -lt $minimumFormulaCount) { throw "Workbook has too few formulas for CORE ${Kind}: $formulaCount" }

    $sharedStrings = Get-SharedStrings -Zip $zip
    $resourcesStart = if ($Kind -eq 'AT') { 'G' } else { 'F' }
    $costPlanningStart = if ($Kind -eq 'AT') { 'H' } else { 'G' }

    $checks = @()
    $checks += Test-MonthHeaders -WorksheetXml (Read-ZipXml -Zip $zip -EntryName $sheetMap['Resources']) -SharedStrings $sharedStrings -SheetName 'Resources' -HeaderRow 6 -StartColumn $resourcesStart -ExpectedStart $ExpectedStart -Months $Months
    $checks += Test-MonthHeaders -WorksheetXml (Read-ZipXml -Zip $zip -EntryName $sheetMap['Cost Planning']) -SharedStrings $sharedStrings -SheetName 'Cost Planning' -HeaderRow 8 -StartColumn $costPlanningStart -ExpectedStart $ExpectedStart -Months $Months
    $checks += Test-MonthHeaders -WorksheetXml (Read-ZipXml -Zip $zip -EntryName $sheetMap['Monthly View']) -SharedStrings $sharedStrings -SheetName 'Monthly View' -HeaderRow 5 -StartColumn 'E' -ExpectedStart $ExpectedStart -Months $Months

    [pscustomobject]@{
        Workbook = $WorkbookPath
        FormulaCount = $formulaCount
        RefErrors = $refErrors
        HeaderChecks = ($checks | ForEach-Object { "$($_.Sheet) $($_.Range) $($_.First)..$($_.Last)" }) -join '; '
    } | Format-List

    Write-Host "OK - CORE $Kind output validated for $Months monthly headers through $($ExpectedStart.AddMonths($Months - 1).ToString('yyyy-MM')) and no #REF!."
}
finally {
    $zip.Dispose()
}
