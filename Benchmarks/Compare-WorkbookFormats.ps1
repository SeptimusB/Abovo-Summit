<# Native Excel/DevExpress IO probe. Originals are only copied and hashed.
   Excel VBA, events, links and calculation are disabled. No production integration. #>
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[ValidateRange(1,5)][int]$Runs=3)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition -eq 'Core'){throw 'Run with Windows PowerShell 5.1 (powershell.exe), required for the .NET Framework compiler.'}
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo 'bin/Release'
$source=(Resolve-Path -LiteralPath $WorkbookPath).Path
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
Add-Type -AssemblyName System.IO.Compression.FileSystem
$package=[IO.Compression.ZipFile]::OpenRead($source)
try {if(@($package.Entries | Where-Object {$_.FullName -match '^xl/(intl)?macrosheets/'}).Count -gt 0){throw 'Excel 4.0 macro sheets are outside this disabled-VBA benchmark.'}} finally {$package.Dispose()}
$out=Join-Path $repo ('obj/WorkbookFormatBenchmarks/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$copy=Join-Path $out 'input.xlsb';Copy-Item -LiteralPath $source -Destination $copy
Write-Output ('OUTPUT='+$out)
function Release-Com($item){if($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($item)}}
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class FormatExcelWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint id); }'
$existing=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {$_.Id})
$excel=$null;$seed=$null;$book=$null;$books=$null;$owned=$false;$rows=@()
try {
    $clock=[Diagnostics.Stopwatch]::StartNew();$excel=New-Object -ComObject Excel.Application
    $startupMs=$clock.ElapsedMilliseconds
    [uint32]$excelProcess=0;[void][FormatExcelWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$excelProcess)
    if($existing -contains $excelProcess){throw 'Excel activation attached to a pre-existing user process; refusing to change it.'}
    $owned=$true
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false;$excel.ScreenUpdating=$false
    $books=$excel.Workbooks
    if($books.Count -ne 0){throw 'New Excel instance loaded unexpected startup workbooks; refusing benchmark.'}
    $seed=$books.Add();$excel.Calculation=-4135;$excel.CalculateBeforeSave=$false
    $info=[ordered]@{Source=$source;SourceSha256=$sourceHash;PrivateInputSha256=(Get-FileHash $copy).Hash;ExcelVersion=$excel.Version;ExcelBuild=$excel.Build;ExcelProcess=$excelProcess;ExcelStartupMs=$startupMs;PowerShellBits=[IntPtr]::Size*8;Runs=$Runs;MacrosExecuted=$false;ExplicitCalculation=$false;ProductionChanged=$false}
    $info | ConvertTo-Json | Set-Content (Join-Path $out 'environment.json')
    foreach($ext in @('xlsb','xlsm')) {
        $book=$books.Open($copy,0,$true)
        $book.CheckCompatibility=$false
        $book.SaveAs((Join-Path $out ('baseline.'+$ext)), $(if($ext -eq 'xlsb'){50}else{52}))
        $book.Close($false);Release-Com $book;$book=$null
    }
    for($run=1;$run -le $Runs;$run++) {
        $formats=if($run%2 -eq 1){@('xlsb','xlsm')}else{@('xlsm','xlsb')}
        foreach($ext in $formats) {
            $inputPath=Join-Path $out ('baseline.'+$ext)
            $outputPath=Join-Path $out ('excel-'+$run+'.'+$ext)
            $clock.Restart();$book=$books.Open($inputPath,0,$true);$openMs=$clock.ElapsedMilliseconds
            $clock.Restart();$book.SaveCopyAs($outputPath);$saveMs=$clock.ElapsedMilliseconds
            $clock.Restart();$book.Close($false);$closeMs=$clock.ElapsedMilliseconds;Release-Com $book;$book=$null
            $rows += [pscustomobject]@{Engine='Excel';Format=$ext;Run=$run;OpenMs=$openMs;SaveMs=$saveMs;CloseMs=$closeMs;Bytes=(Get-Item $outputPath).Length;Output=$outputPath}
            $rows | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $out 'excel.json')
            Write-Output ('Excel '+$ext+' run='+$run+' open='+$openMs+' save='+$saveMs)
        }
    }
} finally {
    if($book){$book.Close($false);Release-Com $book}
    if($seed){$seed.Close($false);Release-Com $seed}
    Release-Com $books
    if($excel){if($owned){$excel.Quit()};Release-Com $excel}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
$refs=@('System','System.Core','System.Drawing','System.Data','System.Web.Extensions','System.Windows.Forms')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Docs.v25.2.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll') | ForEach-Object {Join-Path $bin $_}
$runner=Join-Path $out 'Runner.exe';$compiler=New-Object System.CodeDom.Compiler.CompilerParameters
$compiler.CompilerOptions='/platform:x86';$compiler.GenerateExecutable=$true;$compiler.GenerateInMemory=$false;$compiler.OutputAssembly=$runner
foreach($ref in $refs){[void]$compiler.ReferencedAssemblies.Add($(if($ref.EndsWith('.dll')){$ref}else{$ref+'.dll'}))}
$provider=New-Object Microsoft.CSharp.CSharpCodeProvider
try{$compiled=$provider.CompileAssemblyFromSource($compiler,(Get-Content (Join-Path $repo 'Tools/WorkbookFormatBenchmark.cs') -Raw));if($compiled.Errors.HasErrors){throw ($compiled.Errors | Out-String)}}finally{$provider.Dispose()}
Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($runner+'.config')
& $runner $bin $out $Runs
if($LASTEXITCODE -ne 0){throw 'DevExpress format benchmark failed'}
if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash){throw 'Source hash changed during benchmark; investigate concurrent edits. The benchmark never opens the source in either engine.'}
Write-Output ('PASS: Original hash unchanged. Native IO only; no macro/UDF/financial acceptance. Results: '+$out)
