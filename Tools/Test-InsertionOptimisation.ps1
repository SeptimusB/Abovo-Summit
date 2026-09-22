param([string]$Configuration='Release',[Parameter(Mandatory=$true)][string]$Workbook,
 [ValidateSet('FUNDING_RECORDS','DEVELOPMENT_IDENTIFIED_RECORDS','DEVELOPMENT_MULTIYEAR_RECORDS')][string]$Rule='FUNDING_RECORDS',
 [int]$Count=10,[ValidateSet('capture','insert','scan','recursive','tdb-bounded','tdb-rows','tdb-union','sync-failure','no-history','automatic')][string]$Mode='capture',
 [ValidateSet('x86','x64')][string]$Architecture='x86',[string]$ApplicationAssembly)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot;$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/InsertionOptimisation/'+[guid]::NewGuid().ToString('N'));[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.dll','System.Core.dll','System.Drawing.dll','System.Data.dll','System.Windows.Forms.dll','Microsoft.CSharp.dll')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Docs.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe';$compiler=New-Object System.CodeDom.Compiler.CompilerParameters
$compiler.CompilerOptions='/platform:'+$Architecture
foreach($reference in $refs){[void]$compiler.ReferencedAssemblies.Add($reference)}
$compiler.GenerateExecutable=$true;$compiler.GenerateInMemory=$false;$compiler.OutputAssembly=$runner
$provider=New-Object Microsoft.CSharp.CSharpCodeProvider
try{$compiled=$provider.CompileAssemblyFromSource($compiler,(Get-Content (Join-Path $PSScriptRoot 'InsertionOptimisationFixture.cs') -Raw));if($compiled.Errors.HasErrors){throw ($compiled.Errors | Out-String)}}finally{$provider.Dispose()}
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $out
$hash=(Get-FileHash -LiteralPath $Workbook).Hash
Write-Output ('OUTPUT='+$out)
try{if($ApplicationAssembly){& $runner $bin $Workbook $out $Rule $Count $Mode $ApplicationAssembly}else{& $runner $bin $Workbook $out $Rule $Count $Mode};if($LASTEXITCODE -ne 0){throw 'Insertion optimisation fixture failed'}}
finally{if((Get-FileHash -LiteralPath $Workbook).Hash -ne $hash){throw 'Source changed'};Write-Output ('Source unchanged: '+$hash)}
