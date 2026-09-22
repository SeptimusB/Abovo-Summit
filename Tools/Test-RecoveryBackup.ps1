param([string]$Configuration='Release',[string]$Workbook,[switch]$Probe,[string]$CompareWith,[switch]$VerifyDigests,[string]$CompareSheet,[switch]$CompareValues)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/RecoveryTests/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.dll','System.Core.dll','System.Data.dll','System.Drawing.dll','System.Xml.dll','System.Xml.Linq.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll','System.Windows.Forms.dll','Microsoft.CSharp.dll')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Docs.v25.2.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll') | ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe';$compiler=New-Object System.CodeDom.Compiler.CompilerParameters
$compiler.CompilerOptions='/platform:x86';$compiler.GenerateExecutable=$true;$compiler.OutputAssembly=$runner
foreach($ref in $refs){[void]$compiler.ReferencedAssemblies.Add($ref)}
$provider=New-Object Microsoft.CSharp.CSharpCodeProvider
try{$compiled=$provider.CompileAssemblyFromSource($compiler,(Get-Content (Join-Path $PSScriptRoot 'RecoveryBackupFixture.cs') -Raw));if($compiled.Errors.HasErrors){throw ($compiled.Errors | Out-String)}}finally{$provider.Dispose()}
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $out
Write-Output ('OUTPUT='+$out)
if($Workbook){$hash=(Get-FileHash -LiteralPath $Workbook).Hash}
try{if($CompareValues){& $runner $bin $out $Workbook '--values' $CompareWith}elseif($CompareSheet){& $runner $bin $out $Workbook '--sheet-compare' $CompareWith $CompareSheet}elseif($VerifyDigests){& $runner $bin $out $Workbook '--verify-digests' $CompareWith}elseif($CompareWith){& $runner $bin $out $Workbook '--compare' $CompareWith}elseif($Probe){& $runner $bin $out $Workbook '--probe'}elseif($Workbook){& $runner $bin $out $Workbook}else{& $runner $bin $out};if($LASTEXITCODE -ne 0){throw 'Recovery fixture failed'}}finally{if($Workbook -and (Get-FileHash -LiteralPath $Workbook).Hash -ne $hash){throw 'Test source hash changed'}}
