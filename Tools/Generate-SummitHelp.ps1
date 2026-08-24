param(
    [string]$StructurePath = (Join-Path $PSScriptRoot '..\Structure.xml'),
    [string]$WorkbookPath = 'Z:\Sandbox\TestFileClean.xlsb',
    [string]$OutputPath = (Join-Path $PSScriptRoot '..\Help\data\interfaces.js')
)

$ErrorActionPreference = 'Stop'

function Text-Of($node, [string]$xpath) {
    $match = $node.SelectSingleNode($xpath)
    if ($null -eq $match) { return '' }
    return [string]$match.InnerText.Trim()
}

function Add-UniqueText([System.Collections.Generic.List[string]]$list, [string]$value) {
    if (-not [string]::IsNullOrWhiteSpace($value) -and -not $list.Contains($value)) { $list.Add($value) }
}

$structure = [xml](Get-Content -Raw -LiteralPath $StructurePath)
$commentsBySheet = @{}
$guideByInterface = @{}
$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    try { $excel.AutomationSecurity = 3 } catch {}
    $workbook = $excel.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path, 0, $true)

    foreach ($sheet in $workbook.Worksheets) {
        $sheetComments = [System.Collections.Generic.List[object]]::new()
        try {
            foreach ($comment in $sheet.Comments) {
                $commentText = [string]$comment.Text()
                if (-not [string]::IsNullOrWhiteSpace($commentText)) {
                    $sheetComments.Add([ordered]@{
                        sheet = [string]$sheet.Name
                        address = [string]$comment.Parent.Address($false, $false)
                        text = ($commentText -replace '[\r\n]+', ' ').Trim()
                    })
                }
            }
        } catch {}
        $commentsBySheet[[string]$sheet.Name] = @($sheetComments)
    }

    $guide = $workbook.Worksheets.Item('User Guide')
    $used = $guide.UsedRange
    $currentInterface = ''
    for ($row = 1; $row -le $used.Rows.Count; $row++) {
        $marker = [string]$guide.Cells.Item($row, 2).Text
        $text = [string]$guide.Cells.Item($row, 3).Text
        if ($marker.Trim() -eq '>' -and -not [string]::IsNullOrWhiteSpace($text)) {
            $currentInterface = $text.Trim()
            if (-not $guideByInterface.ContainsKey($currentInterface)) { $guideByInterface[$currentInterface] = [System.Collections.Generic.List[string]]::new() }
        } elseif (-not [string]::IsNullOrWhiteSpace($currentInterface) -and -not [string]::IsNullOrWhiteSpace($text)) {
            $guideByInterface[$currentInterface].Add(($text -replace '[\r\n]+', ' ').Trim())
        }
    }
} finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    if ($workbook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
    if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$groups = [System.Collections.Generic.List[object]]::new()
foreach ($groupNode in $structure.SelectNodes('//GroupStructure')) {
    $children = [System.Collections.Generic.List[object]]::new()
    foreach ($childNode in $groupNode.SelectNodes('./ChildStructure')) {
        $worksheets = [System.Collections.Generic.List[string]]::new()
        Add-UniqueText $worksheets (Text-Of $childNode './DefaultWorksheet')
        foreach ($worksheetNode in $childNode.SelectNodes('.//Worksheet')) { Add-UniqueText $worksheets ([string]$worksheetNode.InnerText.Trim()) }

        $sections = [System.Collections.Generic.List[object]]::new()
        foreach ($sectionNode in $childNode.SelectNodes('./CSInterfaceSection')) {
            $fields = [System.Collections.Generic.List[object]]::new()
            $dataSources = [System.Collections.Generic.List[string]]::new()
            foreach ($sourceNode in $sectionNode.SelectNodes('./ISDatasource')) {
                $sourceName = Text-Of $sourceNode './ISDName'
                $namedRange = Text-Of $sourceNode './/NRDSName'
                $readOnly = Text-Of $sourceNode './RO'
                $summary = $sourceName
                if ($namedRange) { $summary += ' [' + $namedRange + ']' }
                if ($readOnly.ToUpperInvariant() -eq 'TRUE') { $summary += ' (read only)' }
                Add-UniqueText $dataSources $summary
                foreach ($fieldNode in $sourceNode.SelectNodes('.//DataFieldDefinition')) {
                    $fieldName = Text-Of $fieldNode './FieldName'
                    if (-not $fieldName) { $fieldName = $sourceName }
                    $fields.Add([ordered]@{ name = $fieldName; tip = (Text-Of $fieldNode './TipText') })
                }
            }
            $sections.Add([ordered]@{ name = (Text-Of $sectionNode './ISName'); dataSources = @($dataSources); fields = @($fields) })
        }

        $comments = [System.Collections.Generic.List[object]]::new()
        foreach ($worksheetName in $worksheets) {
            if ($commentsBySheet.ContainsKey($worksheetName)) { foreach ($comment in $commentsBySheet[$worksheetName]) { $comments.Add($comment) } }
        }
        $childName = Text-Of $childNode './CSName'
        $guideText = ''
        foreach ($guideName in $guideByInterface.Keys) {
            if ($guideName.Equals($childName, [StringComparison]::OrdinalIgnoreCase)) { $guideText = ($guideByInterface[$guideName] -join ' '); break }
        }
        $children.Add([ordered]@{
            id = [int](Text-Of $childNode './CSID')
            name = $childName
            defaultWorksheet = (Text-Of $childNode './DefaultWorksheet')
            worksheets = @($worksheets)
            userGuide = $guideText
            sections = @($sections)
            comments = @($comments)
        })
    }
    $groups.Add([ordered]@{ id = [int](Text-Of $groupNode './GSID'); name = (Text-Of $groupNode './GSName'); children = @($children) })
}

$payload = [ordered]@{
    generatedUtc = [DateTime]::UtcNow.ToString('o')
    structureName = Text-Of $structure.DocumentElement './Name'
    groups = @($groups)
}
$json = $payload | ConvertTo-Json -Depth 12 -Compress
$outputDirectory = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory | Out-Null }
[IO.File]::WriteAllText($OutputPath, "/* Generated by Tools/Generate-SummitHelp.ps1. Do not hand-edit; use overrides.js. */`r`nwindow.SummitHelpData = $json;`r`n", [Text.UTF8Encoding]::new($false))
Write-Host "Generated Summit help data for $($groups.Count) groups and $(($groups | ForEach-Object { $_.children.Count } | Measure-Object -Sum).Sum) interfaces: $OutputPath"
