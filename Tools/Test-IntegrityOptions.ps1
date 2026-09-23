param([string]$Configuration='Release')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/IntegrityOptionsTests/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$compiler=New-Object System.CodeDom.Compiler.CompilerParameters
$runner=Join-Path $out 'Runner.exe'
$compiler.CompilerOptions='/platform:x86';$compiler.GenerateExecutable=$true;$compiler.OutputAssembly=$runner
foreach($ref in @('System.dll','System.Core.dll','System.Drawing.dll','System.Windows.Forms.dll','Microsoft.CSharp.dll')){[void]$compiler.ReferencedAssemblies.Add($ref)}
$provider=New-Object Microsoft.CSharp.CSharpCodeProvider
try{$result=$provider.CompileAssemblyFromSource($compiler,(Get-Content (Join-Path $PSScriptRoot 'IntegrityOptionsFixture.cs') -Raw));if($result.Errors.HasErrors){throw ($result.Errors|Out-String)}}finally{$provider.Dispose()}
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
Write-Output ('OUTPUT='+$out)
& $runner $bin $out
if($LASTEXITCODE -ne 0){throw 'Integrity Options fixture failed'}
