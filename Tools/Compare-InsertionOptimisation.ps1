param([Parameter(Mandatory=$true)][string]$Baseline,[Parameter(Mandatory=$true)][string]$Candidate)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot;$bin=Join-Path $repo 'bin/Release'
$out=Join-Path $repo ('obj/InsertionComparison/'+[guid]::NewGuid().ToString('N'));[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.Core','System.Drawing','System.Data','System.Web.Extensions')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Docs.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe'
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'StructuralResultComparison.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $refs
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
$hashes=@{};foreach($file in @($Baseline,$Candidate)){$hashes[$file]=(Get-FileHash -LiteralPath $file).Hash}
Write-Output ('OUTPUT='+$out)
try{
 & $runner $bin $out $Baseline $Candidate;if($LASTEXITCODE -ne 0){throw 'Read-only comparison failed'}
 $report=Get-Content (Join-Path $out 'comparison.json') -Raw | ConvertFrom-Json
 if(!$report.complete -or !$report.sameSheetOrder -or $report.formulaDifferences -or $report.constantDifferences -or $report.savedOutputDifferences -or $report.arrayTypeDifferences -or $report.nameDifferences.Count){throw 'Comparison found differences; inspect report'}
 if(@($report.sheets | Where-Object {$_.leftProtected -ne $_.rightProtected}).Count){throw 'Protection flags differ'}
 Write-Output ('PASS: '+$report.formulaCells+' formula cells, names, inputs, selected saved outputs and worksheet protection agree; TDB array classification excluded')
}finally{foreach($file in $hashes.Keys){if((Get-FileHash -LiteralPath $file).Hash -ne $hashes[$file]){throw 'Source changed'}}}
