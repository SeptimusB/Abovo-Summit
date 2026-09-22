param([string]$Configuration='Release',[string]$Workbook,[ValidateSet('anycpu','x86','x64')][string]$Architecture='anycpu',[switch]$StateTracking,[switch]$IncrementalProbe,[ValidateSet('recursive','chain','chain-no-service','recursive-buffer')][string]$SaveStageProbe)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot;$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/SavePreparationTests/'+[guid]::NewGuid().ToString('N'));[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.Core','System.Drawing','System.Data','System.Windows.Forms','Microsoft.CSharp')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Docs.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe'
$compiler=New-Object System.CodeDom.Compiler.CompilerParameters
$compiler.CompilerOptions='/platform:'+$Architecture
foreach($reference in $refs){[void]$compiler.ReferencedAssemblies.Add($(if($reference.EndsWith('.dll')){$reference}else{$reference+'.dll'}))}
[void]$compiler.ReferencedAssemblies.Add('System.dll')
$compiler.GenerateExecutable=$true;$compiler.GenerateInMemory=$false;$compiler.OutputAssembly=$runner
$provider=New-Object Microsoft.CSharp.CSharpCodeProvider
try{$compiled=$provider.CompileAssemblyFromSource($compiler,(Get-Content (Join-Path $PSScriptRoot 'SavePreparationFixture.cs') -Raw));if($compiled.Errors.HasErrors){throw ($compiled.Errors | Out-String)}}finally{$provider.Dispose()}
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $out
Write-Output ('OUTPUT='+$out)
if($Workbook){$hash=(Get-FileHash -LiteralPath $Workbook).Hash}
try {if($StateTracking){& $runner $bin $out '--state'}elseif($SaveStageProbe){if(!$Workbook){throw 'Probe requires a workbook'};& $runner $bin $out $Workbook '--save-stage' $SaveStageProbe}elseif($IncrementalProbe){if(!$Workbook){throw 'Probe requires a workbook'};& $runner $bin $out $Workbook '--incremental'}elseif($Workbook){& $runner $bin $out $Workbook}else{& $runner $bin $out};if($LASTEXITCODE -ne 0){throw 'Save preparation fixture failed'}}finally{if($Workbook -and (Get-FileHash -LiteralPath $Workbook).Hash -ne $hash){throw 'Original changed'}}
