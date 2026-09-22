param([Parameter(Mandatory=$true)][string]$Source,[Parameter(Mandatory=$true)][string]$Reference,[ValidateSet('inspect','prepare','repair')][string]$Mode='inspect',[switch]$RestoreDashboard)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo 'bin/Release'
$out=Join-Path $repo ('obj/ClientFundingRepair/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.Core','System.Drawing','System.Data','System.Web.Extensions','System.Windows.Forms','Microsoft.CSharp')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Docs.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
$refs+=@('DevExpress.XtraSpreadsheet.v25.2.dll','DevExpress.Utils.v25.2.dll','DevExpress.XtraEditors.v25.2.dll','DevExpress.Data.Desktop.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe'
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'ClientFundingRepair.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $refs
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $out
$hash=(Get-FileHash -LiteralPath $Source).Hash
if($Mode -ne 'inspect' -and $hash -ne 'CB3F42A2D8CD5BB073BF571AD4EA5E84495B8B698A6552A5C6714C1482E35AA8') {throw 'This one-off repair is authorised only for the inspected client-file hash'}
try {
 Write-Output ('OUTPUT='+$out)
 $arguments=@($bin,$Source,$Reference,$out,(Join-Path $PSScriptRoot 'ClientFundingRepairSheets.txt'))
 if($Mode -ne 'inspect'){$arguments+=$Mode}
 if($RestoreDashboard){if($Mode -ne 'repair'){throw 'Dashboard restoration requires repair mode'};$arguments+='restore-dashboard'}
 & $runner @arguments
 if($LASTEXITCODE -ne 0){throw 'Funding repair inventory failed'}
}finally{if((Get-FileHash -LiteralPath $Source).Hash -ne $hash){throw 'Source changed'}}
