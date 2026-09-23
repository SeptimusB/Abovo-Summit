param([Parameter(Mandatory=$true)][string]$Workbook,[string]$Configuration='Release',[switch]$RawExport,[ValidateSet('x86','x64')][string]$Architecture='x86')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/RecoveryPackaging/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.dll','System.Core.dll','System.Xml.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Docs.v25.2.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll') | ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe'
$compiler=New-Object System.CodeDom.Compiler.CompilerParameters
$compiler.CompilerOptions='/platform:'+$Architecture;$compiler.GenerateExecutable=$true;$compiler.OutputAssembly=$runner
foreach($ref in $refs){[void]$compiler.ReferencedAssemblies.Add($ref)}
$provider=New-Object Microsoft.CSharp.CSharpCodeProvider
try{$compiled=$provider.CompileAssemblyFromSource($compiler,(Get-Content (Join-Path $PSScriptRoot 'RecoveryPackagingProfile.cs') -Raw));if($compiled.Errors.HasErrors){throw ($compiled.Errors | Out-String)}}finally{$provider.Dispose()}
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
Write-Output ('OUTPUT='+$out)
$hash=(Get-FileHash -LiteralPath $Workbook).Hash
try{& $runner $bin $out $Workbook $(if($RawExport){'--raw'}else{'--export'}) | Tee-Object -FilePath (Join-Path $out 'profile.log');if($LASTEXITCODE -ne 0){throw 'Recovery packaging profile failed'}}finally{if((Get-FileHash -LiteralPath $Workbook).Hash -ne $hash){throw 'Source changed'}}
