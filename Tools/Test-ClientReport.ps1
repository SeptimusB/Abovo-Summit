param([string]$Configuration='Release', [string]$Workbook=(Join-Path (Split-Path -Parent $PSScriptRoot) 'Library/Demo BP v26_0001.xlsb'), [switch]$Inspect, [string]$Fixture='ClientReportFixture.cs', [string]$ReviewCases='')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/ClientReportTests/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.dll','System.Core.dll','System.Data.dll','System.Drawing.dll','System.Xml.dll','System.Xml.Linq.dll','System.Windows.Forms.dll','Microsoft.CSharp.dll')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Docs.v25.2.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.XtraEditors.v25.2.dll','DevExpress.Utils.v25.2.dll','DevExpress.XtraGrid.v25.2.dll','DevExpress.XtraVerticalGrid.v25.2.dll','DevExpress.Data.Desktop.v25.2.dll') | ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe';$compiler=New-Object System.CodeDom.Compiler.CompilerParameters
$compiler.CompilerOptions='/platform:x86';$compiler.GenerateExecutable=$true;$compiler.OutputAssembly=$runner
foreach($ref in $refs){[void]$compiler.ReferencedAssemblies.Add($ref)}
$provider=New-Object Microsoft.CSharp.CSharpCodeProvider
try{$compiled=$provider.CompileAssemblyFromSource($compiler,(Get-Content (Join-Path $PSScriptRoot $Fixture) -Raw));if($compiled.Errors.HasErrors){throw ($compiled.Errors | Out-String)}}finally{$provider.Dispose()}
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $out
$hash=(Get-FileHash -LiteralPath $Workbook).Hash
Write-Output ('OUTPUT='+$out)
try{& $runner $bin $out $Workbook ([string]$Inspect) $ReviewCases | Tee-Object -FilePath (Join-Path $out 'test.log');if($LASTEXITCODE -ne 0){throw 'Client report fixture failed'}}finally{if((Get-FileHash -LiteralPath $Workbook).Hash -ne $hash){throw 'Source workbook changed'};Write-Output 'Source hash unchanged'}
