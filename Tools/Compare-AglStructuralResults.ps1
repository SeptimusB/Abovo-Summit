param([string]$InputDirectory='C:/Sandbox/Insert Comp')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot;$bin=Join-Path $repo 'bin/Release'
$prefix='AGL - BP 2627 updated Mar26 v25 v2 v26_0005'
$files=@('.xlsb',' 10 funding records in summit.xlsb',' 10 funding cols added in excel.xlsb',' 10 development lines added in summit.xlsb',' 10 development cols added in excel.xlsb') | ForEach-Object {Join-Path $InputDirectory ($prefix+$_)}
$hashes=@{};foreach($file in $files){$hashes[$file]=(Get-FileHash -LiteralPath $file).Hash}
$out=Join-Path $repo ('obj/AglStructuralComparison/'+[guid]::NewGuid().ToString('N'));[void](New-Item -ItemType Directory -Path $out)
$refs=@('System.Core','System.Drawing','System.Data','System.Web.Extensions')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Docs.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe'
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'StructuralResultComparison.cs') -Raw) -OutputAssembly $runner -OutputType ConsoleApplication -ReferencedAssemblies $refs
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
Write-Output ('Output: '+$out)
try {& $runner $bin $out @files;if($LASTEXITCODE -ne 0){throw 'Read-only comparison failed'}}
finally {foreach($file in $files){if((Get-FileHash -LiteralPath $file).Hash -ne $hashes[$file]){throw ('Source changed: '+$file)}};Write-Output 'PASS: All five source hashes unchanged'}
