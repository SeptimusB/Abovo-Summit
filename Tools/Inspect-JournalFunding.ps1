param([string]$Workbook=(Join-Path (Split-Path -Parent $PSScriptRoot) 'Library/Demo BP v26_0001.xlsb'),[switch]$Insert)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo 'bin/Release'
$out=Join-Path $repo ('obj/JournalFundingInspection/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.Windows.Forms','System.Drawing','System.Core','System.Data','Microsoft.CSharp')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll')|ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe'
$refs+=Join-Path $bin 'DevExpress.Docs.v25.2.dll'
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'JournalFundingInspection.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $refs
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $out
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
$hash=(Get-FileHash -LiteralPath $Workbook).Hash
Write-Output ('Output: '+$out)
try { & $runner $bin $Workbook $out ([string]$Insert); if($LASTEXITCODE -ne 0){throw 'Inspection failed'} }
finally { if((Get-FileHash -LiteralPath $Workbook).Hash -ne $hash){throw 'Source changed'} }
