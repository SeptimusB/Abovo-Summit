param([string]$InputDirectory='C:/Sandbox/Insert Comp')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot;$bin=Join-Path $repo 'bin/Release'
$out=Join-Path $repo ('obj/AglCalculation/'+[guid]::NewGuid().ToString('N'));[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.Core','System.Drawing')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Docs.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe'
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'AglCalculationProbe.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $refs
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
$prefix='AGL - BP 2627 updated Mar26 v25 v2 v26_0005'
$suffixes=@('.xlsb',' 10 funding records in summit.xlsb',' 10 funding cols added in excel.xlsb')
Write-Output ('OUTPUT='+$out)
for($i=0;$i -lt $suffixes.Count;$i++){
 $file=Join-Path $InputDirectory ($prefix+$suffixes[$i]);$hash=(Get-FileHash -LiteralPath $file).Hash
 try{& $runner $bin $file (Join-Path $out ($i.ToString()+'.tsv')) | Tee-Object -FilePath (Join-Path $out ($i.ToString()+'.log'));if($LASTEXITCODE -ne 0){throw 'Calculation probe failed'}}finally{if((Get-FileHash -LiteralPath $file).Hash -ne $hash){throw 'Original changed'}}
}
Write-Output 'PASS: Sources unchanged; no macros executed; calculated in memory only'
