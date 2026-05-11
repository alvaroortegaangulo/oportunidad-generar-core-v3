param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
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
    if ($null -eq $workbook -or $null -eq $rels) {
        throw 'Workbook metadata not found.'
    }

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

function Measure-MonthCapacity {
    param(
        [xml]$WorksheetXml,
        [int]$Row,
        [string]$StartColumn
    )

    $ns = [Xml.XmlNamespaceManager]::new($WorksheetXml.NameTable)
    $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    $startIndex = Get-ColIndex $StartColumn
    $cellsByColumn = @{}
    foreach ($cell in $WorksheetXml.SelectNodes("//x:sheetData/x:row[@r='$Row']/x:c", $ns)) {
        if ($cell.r -match '^([A-Z]+)\d+$') {
            $cellsByColumn[(Get-ColIndex $matches[1])] = $cell
        }
    }

    $count = 0
    for ($column = $startIndex; $column -le 300; $column++) {
        if (-not $cellsByColumn.ContainsKey($column)) { break }
        $cell = $cellsByColumn[$column]
        $hasFormula = $null -ne $cell.f
        $hasValue = $null -ne $cell.v -and -not [string]::IsNullOrWhiteSpace($cell.v.InnerText)
        if (-not ($hasFormula -or $hasValue)) { break }
        $count++
    }

    return [pscustomobject]@{
        Count = $count
        From = $StartColumn
        To = Get-ColLetters ($startIndex + $count - 1)
    }
}

function Test-CoreTemplate {
    param(
        [string]$Path,
        [string]$Kind,
        [int]$ExpectedMonths
    )

    $requiredSheets = @('Project Infor', 'Cost Overview', 'Resources', 'Cost Planning', 'Monthly View', 'Cost Summary', 'Ayuda', 'aux Billing Plan')
    $zip = [IO.Compression.ZipFile]::OpenRead((Resolve-Path $Path))
    try {
        $sheetMap = Get-SheetMap -Zip $zip
        $missing = $requiredSheets | Where-Object { -not $sheetMap.ContainsKey($_) }
        if ($missing) { throw "Missing required sheets in ${Kind}: $($missing -join ', ')" }

        $formulaCount = 0
        $refErrors = 0
        $dataValidations = 0
        $styledCells = 0
        $dimensions = @{}
        foreach ($sheetName in $sheetMap.Keys) {
            $text = Read-ZipText -Zip $zip -EntryName $sheetMap[$sheetName]
            if ($text -match '#REF!') {
                $refErrors += [regex]::Matches($text, '#REF!').Count
            }
            $xml = [xml]$text
            $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
            $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            $formulaCount += $xml.SelectNodes('//x:f', $ns).Count
            $dataValidations += $xml.SelectNodes('//x:dataValidations/x:dataValidation', $ns).Count
            $styledCells += $xml.SelectNodes('//x:c[@s]', $ns).Count
            $dimension = $xml.SelectSingleNode('//x:dimension', $ns)
            if ($dimension) { $dimensions[$sheetName] = $dimension.ref }
        }

        if ($refErrors -ne 0) { throw "Template ${Kind} contains #REF! markers: $refErrors" }
        if ($formulaCount -lt 3000) { throw "Template ${Kind} has too few formulas: $formulaCount" }
        if ($dataValidations -lt 1) { throw "Template ${Kind} has no data validations." }
        if ($styledCells -lt 10000) { throw "Template ${Kind} has too few styled cells: $styledCells" }

        $resourcesXml = Read-ZipXml -Zip $zip -EntryName $sheetMap['Resources']
        $costPlanningXml = Read-ZipXml -Zip $zip -EntryName $sheetMap['Cost Planning']
        $monthlyViewXml = Read-ZipXml -Zip $zip -EntryName $sheetMap['Monthly View']

        $resourceStart = if ($Kind -eq 'AT') { 'G' } else { 'F' }
        $costPlanningStart = if ($Kind -eq 'AT') { 'H' } else { 'G' }

        $resourcesMonths = Measure-MonthCapacity -WorksheetXml $resourcesXml -Row 6 -StartColumn $resourceStart
        $costPlanningMonths = Measure-MonthCapacity -WorksheetXml $costPlanningXml -Row 8 -StartColumn $costPlanningStart
        $monthlyViewMonths = Measure-MonthCapacity -WorksheetXml $monthlyViewXml -Row 5 -StartColumn 'E'

        foreach ($measure in @(
            @('Resources', $resourcesMonths),
            @('Cost Planning', $costPlanningMonths),
            @('Monthly View', $monthlyViewMonths)
        )) {
            if ($measure[1].Count -ne $ExpectedMonths) {
                throw "Template ${Kind} $($measure[0]) has $($measure[1].Count) months; expected $ExpectedMonths."
            }
        }

        [pscustomobject]@{
            Template = Split-Path $Path -Leaf
            Kind = $Kind
            MonthsResources = "$($resourcesMonths.Count) [$($resourcesMonths.From):$($resourcesMonths.To)]"
            MonthsCostPlanning = "$($costPlanningMonths.Count) [$($costPlanningMonths.From):$($costPlanningMonths.To)]"
            MonthsMonthlyView = "$($monthlyViewMonths.Count) [$($monthlyViewMonths.From):$($monthlyViewMonths.To)]"
            Formulas = $formulaCount
            DataValidations = $dataValidations
            StyledCells = $styledCells
            Dimensions = ($dimensions.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '; '
        }
    }
    finally {
        $zip.Dispose()
    }
}

$templatesDir = Join-Path $ProjectRoot 'data/templates'
$results = @(
    Test-CoreTemplate -Path (Join-Path $templatesDir 'CORE_PC_template.xlsx') -Kind 'PC' -ExpectedMonths 60
    Test-CoreTemplate -Path (Join-Path $templatesDir 'CORE_AT_template.xlsx') -Kind 'AT' -ExpectedMonths 60
)

$results | Format-List
Write-Host 'OK - CORE PC/AT templates validated for 60 useful months.'
