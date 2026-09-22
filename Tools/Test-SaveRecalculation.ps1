param([Parameter(Mandatory=$true)][string]$SavedWorkbook)
# Diagnostic only: VBA/UDF-dependent models may report NAME errors under the
# mandatory macro-disabled policy, including unchanged source masters. A failure
# here must be investigated, not bypassed by automatically enabling macros.
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$out=Join-Path $repo ('obj/SaveRecalculation/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$roundtrip=Join-Path $out 'excel-recalculated.xlsb'
$hash=(Get-FileHash -LiteralPath $SavedWorkbook).Hash
function Release-Com($item){if($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($item)}}
$excel=$null;$seed=$null;$book=$null;$cached=@{}
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false
    $seed=$excel.Workbooks.Add();$excel.Calculation=-4135;$excel.CalculateBeforeSave=$false
    $book=$excel.Workbooks.Open($SavedWorkbook,0,$true)
    $sheets=@('Detailed Comp Inc - Trad View','Financial Position - Trad View','Cashflow detailed','Check Sheet')
    foreach($name in $sheets){$ws=$book.Worksheets.Item($name);$range=$ws.UsedRange;$cached[$name]=$range.Value2;Release-Com $range;Release-Com $ws}
    $clock=[Diagnostics.Stopwatch]::StartNew();$excel.CalculateFullRebuild()
    Write-Output ('EXCEL_FULL_REBUILD_MS='+$clock.ElapsedMilliseconds)
    foreach($name in $sheets){
        $ws=$book.Worksheets.Item($name);$range=$ws.UsedRange;$rebuilt=$range.Value2;Release-Com $range;Release-Com $ws
        $before=$cached[$name];$differences=0
        if($before.GetLength(0) -ne $rebuilt.GetLength(0) -or $before.GetLength(1) -ne $rebuilt.GetLength(1)){throw ('Output dimensions differ: '+$name)}
        for($r=1;$r -le $before.GetLength(0);$r++){for($c=1;$c -le $before.GetLength(1);$c++){
            $a=$before[$r,$c];$b=$rebuilt[$r,$c]
            if($a -is [double] -and $b -is [double]){if([Math]::Abs($a-$b) -le 1e-7){continue}}elseif([object]::Equals($a,$b)){continue}
            $differences++;if($differences -le 8){Write-Output ('DIFFERENCE '+$name+' R'+$r+'C'+$c+' saved='+$a+' Excel='+$b)}
        }}
        Write-Output ($name+' differences='+$differences)
        if($differences -ne 0){throw ('Saved caches differ from full Excel rebuild: '+$name)}
    }
    $book.SaveCopyAs($roundtrip)
    Write-Output ('PASS: Saved output caches agree with independent Excel rebuild; ROUNDTRIP='+$roundtrip)
} finally {
    if($book){$book.Close($false);Release-Com $book}
    if($seed){$seed.Close($false);Release-Com $seed}
    if($excel){$excel.Quit();Release-Com $excel}
    if((Get-FileHash -LiteralPath $SavedWorkbook).Hash -ne $hash){throw 'Supplied workbook changed'}
}
