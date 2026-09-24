# Internal packaging/verification helper. Not included in the support ZIP.
# Requires the caller's process-local SG_SUPPORT_LICENSE for the full control run.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression.FileSystem
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$source = Join-Path $repo 'Tools/SpreadsheetGearSupportRepro'
$output = Join-Path $repo 'Library/Support'
$name = 'SpreadsheetGear_Support_Package_v01_2026-09-24.zip'
$final = Join-Path $output $name
if (Test-Path -LiteralPath $final) { throw 'Versioned support ZIP already exists; do not overwrite it.' }
if ([string]::IsNullOrWhiteSpace($env:SG_SUPPORT_LICENSE)) { throw 'Supply a process-local licence to validate all probes. It is never packaged.' }
$run = Join-Path $repo ('obj/SpreadsheetGearSupport/package-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $run,$output -Force | Out-Null
$files = @('Repro.csproj','SpreadsheetGearSupportRepro.sln','NuGet.Config','packages.lock.json',
    'Program.cs','ExcelProbe.cs','PackageAudit.cs','Run-Repro.ps1','README.md','Support-email.txt',
    'Samples/CustomXmlAndDynamicArrays.xlsm','Evidence/observations.json','Evidence/gear-save-only.xlsm',
    'Evidence/gear-serial-insert.xlsx','Evidence/excel-serial-insert.xlsx','Evidence/excel-grouped-insert.xlsx')
$script:assertions = 0
function Assert-Support([bool]$condition,[string]$description) {
    if (-not $condition) { throw ('Verification failed: ' + $description) }
    $script:assertions++
}
function Write-SupportZip([string]$path,[string[]]$entries) {
    $zip = [IO.Compression.ZipFile]::Open($path,[IO.Compression.ZipArchiveMode]::Create)
    try { foreach ($file in $entries) {
        [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip,(Join-Path $source $file),$file,[IO.Compression.CompressionLevel]::Optimal) | Out-Null
    }} finally { $zip.Dispose() }
}
function Check-Report([string]$path,[bool]$full) {
    $report = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    Assert-Support ($report.spreadsheetGearVersion -eq '9.3.85.102') 'Pinned engine version'
    Assert-Support ($report.processBits -eq 64) '64-bit reproduction'
    Assert-Support $report.inputFileUnchanged 'Synthetic input unchanged'
    Assert-Support ($report.inputPackage.customParts.Count -eq 3) 'Input has 3 custom XML parts'
    Assert-Support ($report.gearSavedPackage.customParts.Count -eq 0) 'Saved custom XML loss reproduced'
    Assert-Support $report.inputPackage.hasMetadata 'Input has dynamic metadata'
    Assert-Support (-not $report.gearSavedPackage.hasMetadata) 'Dynamic metadata loss reproduced'
    Assert-Support ($report.inputPackage.cellsWithMetadata.Count -eq 1) 'Input cell metadata association'
    Assert-Support ($report.gearSavedPackage.cellsWithMetadata.Count -eq 0) 'Output cell association removed'
    Assert-Support ($report.inputPackage.arrayCells[0].formula -eq $report.gearSavedPackage.arrayCells[0].formula) 'Formula text preserved'
    Assert-Support ($report.gearSavedPackage.arrayCells[0].reference -eq 'B2:B4') 'Fixed array extent retained'
    Assert-Support (-not $report.inputPackage.hasVba -and -not $report.gearSavedPackage.hasVba) 'No VBA payloads'
    Assert-Support ($report.inputPackage.externalRelationships.Count -eq 0 -and $report.gearSavedPackage.externalRelationships.Count -eq 0) 'No external relationships'
    if ($full) {
        Assert-Support ($report.gearSerialInsert.after -eq '=SUM(First:Last!B1)' -and $report.gearSerialInsert.value -eq 0) 'Gear serial result'
        Assert-Support ($report.excelSerialInsert.after -eq '=SUM(First:Last!B1)' -and $report.excelSerialInsert.value -eq 0) 'Excel serial control'
        Assert-Support ($report.excelGroupedInsert.after -eq '=SUM(First:Last!D1)' -and $report.excelGroupedInsert.value -eq 12) 'Excel grouped control'
        foreach ($probe in @($report.gearSerialInsert,$report.excelSerialInsert,$report.excelGroupedInsert)) {
            Assert-Support ($probe.beforeValue -eq 12 -and $probe.firstD1 -eq 5 -and $probe.lastD1 -eq 7) 'Original values moved to D1'
        }
        Assert-Support $report.originalArrayInExcel.hasSpillAfter 'Original expands as a dynamic array'
        Assert-Support (($report.originalArrayInExcel.valuesAfter -join ',') -eq '1,2,3,4,5') 'Original spill expands to five cells'
        Assert-Support $report.gearSavedArrayInExcel.hasArrayAfter 'Saved output is a legacy array'
        Assert-Support (-not $report.gearSavedArrayInExcel.hasSpillAfter) 'Saved output no longer spills'
        Assert-Support (($report.gearSavedArrayInExcel.valuesAfter -join ',') -eq '1,2,3,,') 'Saved output does not expand'
    } else { Assert-Support ($report.gearSerialInsert -is [string] -and $report.gearSerialInsert.StartsWith('Skipped:')) 'Unlicensed insertion is explicitly skipped' }
}
function Check-Text([string]$text,[string]$label) {
    # Check exact secrets without printing them, plus known private identifiers.
    Assert-Support (-not $text.Contains($env:SG_SUPPORT_LICENSE)) ('No supplied licence in ' + $label)
    Assert-Support ($text -notmatch '(?i)(Signature=|SpreadsheetGear\.License, Type=|jmwor|silvervent|Stori|\bAGL\b|(?<![A-Za-z])[A-Z]:[\\/])') ('Privacy scan: ' + $label)
}
function Check-Workbook([string]$path) {
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try { foreach ($part in $zip.Entries) {
        Assert-Support ($part.FullName -notmatch '(?i)(vbaProject|externalLinks|\.bin$)') 'No VBA, binary payload or external workbook link'
        if ($part.FullName -match '\.(xml|rels)$') {
            $reader = [IO.StreamReader]::new($part.Open())
            try { $content = $reader.ReadToEnd() } finally { $reader.Dispose() }
            Check-Text $content $part.FullName
            Assert-Support ($content -notmatch 'TargetMode="External"') 'No external package relationship'
        }
    }} finally { $zip.Dispose() }
}

# First ZIP contains the exact source and sample files which will be delivered.
$candidate = Join-Path $run 'candidate.zip'
Write-SupportZip $candidate $files
$extracted = Join-Path $run 'extracted'
[IO.Compression.ZipFile]::ExtractToDirectory($candidate,$extracted)
foreach ($config in @('Debug','Release')) {
    & dotnet build (Join-Path $extracted 'SpreadsheetGearSupportRepro.sln') -c $config --packages (Join-Path $repo 'packages/engine-trial') -p:RestoreLockedMode=true --nologo
    Assert-Support ($LASTEXITCODE -eq 0) ("Clean extracted $config build")
}
$fullOutput = Join-Path $run 'licensed-controls'
& (Join-Path $extracted 'bin/Release/net48/Repro.exe') --excel --output $fullOutput
Assert-Support ($LASTEXITCODE -eq 0) 'Clean extracted licensed Excel control run'
Check-Report (Join-Path $fullOutput 'observations.json') $true
Check-Report (Join-Path $source 'Evidence/observations.json') $true
$secret = $env:SG_SUPPORT_LICENSE
try {
    [Environment]::SetEnvironmentVariable('SG_SUPPORT_LICENSE',$null,'Process')
    foreach ($config in @('Debug','Release')) {
        $freeOutput = Join-Path $run ('free-' + $config)
        & (Join-Path $extracted "bin/$config/net48/Repro.exe") --output $freeOutput
        Assert-Support ($LASTEXITCODE -eq 0) ("Clean extracted unlicensed $config run")
        Check-Report (Join-Path $freeOutput 'observations.json') $false
    }
} finally { $env:SG_SUPPORT_LICENSE = $secret; $secret = $null }
foreach ($file in $files) {
    $original = Join-Path $source $file
    Assert-Support ((Get-FileHash -LiteralPath $original).Hash -eq (Get-FileHash -LiteralPath (Join-Path $extracted $file)).Hash) 'Exact source/evidence extraction'
    if ([IO.Path]::GetExtension($file) -in '.xlsx','.xlsm') { Check-Workbook $original }
    else { Check-Text (Get-Content -LiteralPath $original -Raw) $file }
}
$verification = [ordered]@{
    packageVersion = 'v01'; verifiedUtc = [DateTime]::UtcNow.ToString('o');
    status = 'Reproductions verified; this is not an engine compatibility pass';
    scope = 'Clean extracted solution; all source/sample/evidence files identical to final package';
    debugBuild = 'passed'; releaseBuild = 'passed';
    unlicensedDebugAndReleaseRuns = 'passed (whole-column test explicitly skipped)';
    licensedExcelControls = 'passed (all three observed behaviors reproduced)';
    privacyScan = 'passed, including nested workbook XML; no licence, customer information, VBA, binary workbook payloads, external links or local absolute paths';
    assertionsPassed = $script:assertions;
    originalSampleSha256 = (Get-FileHash -LiteralPath (Join-Path $source 'Samples/CustomXmlAndDynamicArrays.xlsm')).Hash
}
$verification | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $source 'Evidence/verification.json') -Encoding UTF8
$files += 'Evidence/verification.json'
$manifest = foreach ($file in $files) { [ordered]@{file=$file;sha256=(Get-FileHash -LiteralPath (Join-Path $source $file)).Hash} }
ConvertTo-Json -InputObject @($manifest) -Depth 4 | Set-Content -LiteralPath (Join-Path $source 'Manifest.json') -Encoding UTF8
$files += 'Manifest.json'
Write-SupportZip $final $files
$finalExtract = Join-Path $run 'final-package'
[IO.Compression.ZipFile]::ExtractToDirectory($final,$finalExtract)
foreach ($file in $files) {
    Assert-Support ((Get-FileHash -LiteralPath (Join-Path $source $file)).Hash -eq (Get-FileHash -LiteralPath (Join-Path $finalExtract $file)).Hash) 'Final ZIP entry hash'
}
Copy-Item -LiteralPath (Join-Path $source 'Support-email.txt') -Destination (Join-Path $output 'SpreadsheetGear_Support_Email_v01_2026-09-24.txt')
[pscustomobject]@{Zip=$final;Bytes=(Get-Item -LiteralPath $final).Length;Sha256=(Get-FileHash -LiteralPath $final).Hash;Assertions=$script:assertions;CleanTestDirectory=$run}
