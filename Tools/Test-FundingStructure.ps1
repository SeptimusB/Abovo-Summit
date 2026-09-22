param([string]$Configuration='Release',[string]$Workbook=(Join-Path (Split-Path -Parent $PSScriptRoot) 'Library/Blank BP v26_0001.xlsb'))
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/FundingStructureTests/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.Windows.Forms','System.Drawing','System.Core','System.Data','Microsoft.CSharp')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Docs.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe'
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'FundingStructureFixture.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $refs
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $out
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
$hash=(Get-FileHash -LiteralPath $Workbook).Hash
Write-Output ('Output: '+$out)
try {& $runner $bin $Workbook $out; if($LASTEXITCODE -ne 0){throw 'Funding structure fixture failed'}}
finally {if((Get-FileHash -LiteralPath $Workbook).Hash -ne $hash){throw 'Source changed'};Write-Output 'Source hash unchanged'}
