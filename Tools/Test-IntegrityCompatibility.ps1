param([string]$Configuration='Release',[string]$Workbook,[switch]$UI)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/IntegrityCompatibility/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.dll','System.Core.dll','System.Data.dll','System.Drawing.dll','System.Xml.dll','System.Windows.Forms.dll','Microsoft.CSharp.dll','System.Configuration.dll')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Docs.v25.2.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll') | ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe';$compiler=New-Object System.CodeDom.Compiler.CompilerParameters
$compiler.CompilerOptions='/platform:x86';$compiler.GenerateExecutable=$true;$compiler.OutputAssembly=$runner
foreach($ref in $refs){[void]$compiler.ReferencedAssemblies.Add($ref)}
$provider=New-Object Microsoft.CSharp.CSharpCodeProvider
try{$compiled=$provider.CompileAssemblyFromSource($compiler,(Get-Content (Join-Path $PSScriptRoot 'IntegrityCompatibilityFixture.cs') -Raw));if($compiled.Errors.HasErrors){throw ($compiled.Errors | Out-String)}}finally{$provider.Dispose()}
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $out
Write-Output ('OUTPUT='+$out)
if($Workbook){$hash=(Get-FileHash -LiteralPath $Workbook).Hash}
try{& $runner $bin $out $(if($Workbook){$Workbook}else{'--synthetic'}) $(if($UI){'--ui'}else{'--native'});if($LASTEXITCODE -ne 0){throw 'Integrity compatibility fixture failed'}}finally{if($Workbook -and (Get-FileHash -LiteralPath $Workbook).Hash -ne $hash){throw 'Source changed'}}
