param([Parameter(Mandatory=$true)][string]$Source,[Parameter(Mandatory=$true)][string]$Candidate,[Parameter(Mandatory=$true)][string]$Manifest)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot;$bin=Join-Path $repo 'bin/Release'
$out=Join-Path $repo ('obj/ClientFundingValidation/'+[guid]::NewGuid().ToString('N'));[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.Core','System.Drawing','System.Data','System.Web.Extensions')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Docs.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe'
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'ValidateClientFundingRepair.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $refs
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
Write-Output ('OUTPUT='+$out)
& $runner $bin $Source $Candidate $Manifest $out
if($LASTEXITCODE -ne 0){throw 'Validation failed'}
