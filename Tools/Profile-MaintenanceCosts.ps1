param([ValidateSet('Debug','Release')][string]$Configuration='Release',[string]$Workbook)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot;$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/MaintenanceCostProfile/'+[guid]::NewGuid().ToString('N'));[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.dll','System.Core.dll','System.Data.dll','System.Drawing.dll','System.Windows.Forms.dll','System.Xml.dll')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Docs.v25.2.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll')|ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe';$compiler=New-Object System.CodeDom.Compiler.CompilerParameters
$compiler.CompilerOptions='/platform:x86 /optimize+';$compiler.GenerateExecutable=$true;$compiler.OutputAssembly=$runner
foreach($ref in $refs){[void]$compiler.ReferencedAssemblies.Add($ref)}
$provider=New-Object Microsoft.CSharp.CSharpCodeProvider
try{$c=$provider.CompileAssemblyFromSource($compiler,(Get-Content -LiteralPath (Join-Path $PSScriptRoot 'MaintenanceCostProfile.cs') -Raw));if($c.Errors.HasErrors){throw ($c.Errors|Out-String)}}finally{$provider.Dispose()}
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
Write-Output ('OUTPUT='+$out)
if($Workbook){$source=(Resolve-Path -LiteralPath $Workbook).Path;$hash=(Get-FileHash -LiteralPath $source).Hash}
try{
 if($Workbook){& $runner $bin $source | Tee-Object -FilePath (Join-Path $out 'profile.log')}else{& $runner $bin | Tee-Object -FilePath (Join-Path $out 'profile.log')}
 if($LASTEXITCODE -ne 0){throw 'Maintenance cost profile failed'}
}finally{if($Workbook -and (Get-FileHash -LiteralPath $source).Hash -ne $hash){throw 'Source workbook changed'}}
