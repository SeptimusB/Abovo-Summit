param([Parameter(Mandatory=$true)][string[]]$Files,[Parameter(Mandatory=$true)][string]$Report,[switch]$ProbeStyles,[switch]$ProbeArrays)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$root = [IO.Path]::GetFullPath((Join-Path $repo 'obj/AsposeTrial')) + [IO.Path]::DirectorySeparatorChar
function Assert-PrivatePath([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path)
    if (-not $full.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)) { throw 'Disposable trial paths only.' }
    if ((Test-Path -LiteralPath $full) -and ((Get-Item -LiteralPath $full).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Reparse point prohibited.' }
    $parent = [IO.DirectoryInfo]::new([IO.Path]::GetDirectoryName($full))
    while ($null -ne $parent) {
        if ($parent.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Reparse point prohibited.' }
        $parent = $parent.Parent
    }
    return $full
}
$Report = Assert-PrivatePath $Report
if (Test-Path -LiteralPath $Report) { throw 'Report must be new.' }
foreach ($file in $Files) { $null = Assert-PrivatePath $file }
function Release-Com($item) {
    if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
}
$excel=$null; $seed=$null; $book=$null; $results=@()
try {
    # Own instance only. No GetActiveObject, process termination or source saves.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false
    $excel.AutomationSecurity=3; $excel.AskToUpdateLinks=$false
    $seed=$excel.Workbooks.Add(); $excel.Calculation=-4135; $excel.CalculateBeforeSave=$false
    foreach ($file in $Files) {
        $before=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
        try {
            # Default xlNormalLoad, no repair mode requested. Macros/events/links disabled.
            $book=$excel.Workbooks.Open($file,0,$true)
            $item=[pscustomobject]@{
                file=$file; excelVersion=$excel.Version; excelBuild=$excel.Build
                opened=$true; readOnly=$book.ReadOnly; openMode='xlNormalLoad (default; no recovery requested)'
                worksheets=$book.Worksheets.Count; names=$book.Names.Count
                hasVba=$book.HasVBProject; customXmlParts=$book.CustomXMLParts.Count
                calculation=$excel.Calculation; macrosDisabled=($excel.AutomationSecurity -eq 3)
            }
            if ($ProbeStyles) {
                $styles=@()
                foreach ($probe in @(@('Development BP Assumptions','Q66'),@('Development BP Assumptions','Q98'),@('Stock Condition Inputs','P8'),@('Rent Assumptions','D22'),@('OW - Live Stress Reporting','AL23'))) {
                    $sheet=$book.Worksheets.Item($probe[0]); $cell=$sheet.Range($probe[1])
                    try {
                        $styles+=[pscustomobject]@{sheet=$probe[0];cell=$probe[1];pattern=$cell.Interior.Pattern;colour=$cell.Interior.Color;patternColour=$cell.Interior.PatternColor;fontBold=$cell.Font.Bold;fontColour=$cell.Font.Color;locked=$cell.Locked;displayPattern=$cell.DisplayFormat.Interior.Pattern;conditionalRules=$cell.FormatConditions.Count}
                    } finally { Release-Com $cell; Release-Com $sheet }
                }
                $item | Add-Member NoteProperty styleProbes $styles
            }
            $results+=$item
            if ($ProbeArrays) {
                $arrays=@()
                foreach ($probe in @(@('Management Costs Assumptions','D61'),@('Dvpt BP Rev and Exp Assumptions','D51'),@('Leaseholder Units','C13'))) {
                    $sheet=$book.Worksheets.Item($probe[0]); $cell=$sheet.Range($probe[1])
                    try {
                        $arrays+=[pscustomobject]@{sheet=$probe[0];cell=$probe[1];hasArray=$cell.HasArray;hasSpill=$cell.HasSpill;formula=$cell.Formula;formula2=$cell.Formula2}
                    } finally { Release-Com $cell; Release-Com $sheet }
                }
                $item | Add-Member NoteProperty arrayProbes $arrays
            }
            $book.Close($false); Release-Com $book; $book=$null
        } finally {
            if ($book) { $book.Close($false); Release-Com $book; $book=$null }
            if ((Get-FileHash -LiteralPath $file).Hash -ne $before) { throw 'Private input hash changed.' }
        }
    }
} finally {
    if ($book) { $book.Close($false); Release-Com $book }
    if ($seed) { $seed.Close($false); Release-Com $seed }
    if ($excel) { $excel.Quit(); Release-Com $excel }
    $results | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Report -Encoding UTF8
}
$results | Format-List
Write-Output 'No calculation, VBA execution or SaveCopyAs performed. This checks normal Excel opening only, not financial acceptance.'
