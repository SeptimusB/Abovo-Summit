<# Isolated, UNSAVED microbenchmarks. These are not complete valid structural
   commands: source-only probes omit name/mirror post-actions; mirror-only probes
   deliberately omit source expansion. No result is saved or offered to clients. #>
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[ValidateRange(1,100)][int]$Count=10)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$source=(Resolve-Path -LiteralPath $WorkbookPath).Path
$hash=(Get-FileHash -LiteralPath $source).Hash
Add-Type -AssemblyName System.IO.Compression.FileSystem
$package=[IO.Compression.ZipFile]::OpenRead($source)
try {if(@($package.Entries | Where-Object {$_.FullName -match '^xl/(intl)?macrosheets/'}).Count -gt 0){throw 'Excel 4.0 macro sheets are outside this disabled-VBA probe.'}} finally {$package.Dispose()}
$out=Join-Path $repo ('obj/ExcelStructuralCore/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$inputPath=Join-Path $out 'private-input.xlsb';Copy-Item -LiteralPath $source -Destination $inputPath
Write-Output ('OUTPUT='+$out)
function Release-Com($item){if($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($item)}}
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class StructuralExcelWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint id); }'
$existing=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {$_.Id})
$excel=$null;$books=$null;$book=$null;$seed=$null;$owned=$false;$results=@()
try {
    $excel=New-Object -ComObject Excel.Application
    [uint32]$excelProcess=0;[void][StructuralExcelWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$excelProcess)
    if($existing -contains $excelProcess){throw 'Refusing to use a pre-existing Excel process'}
    $owned=$true;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false;$excel.ScreenUpdating=$false
    $books=$excel.Workbooks;if($books.Count -ne 0){throw 'Unexpected startup workbook'}
    $seed=$books.Add();$excel.Calculation=-4135;$excel.CalculateBeforeSave=$false
    foreach($operation in @('DevelopmentMirrors','FundingColumns','DevelopmentColumns')) {
        $book=$books.Open($inputPath,0,$true)
        $watch=[Diagnostics.Stopwatch]::StartNew();$shiftMs=0L;$copyMs=0L;$completed=0;$errorText=$null;$widths=@();$beforeWidths=@{}
        try {
            if($operation -eq 'DevelopmentMirrors') {
                $ws=$book.Worksheets.Item('Transactional DB')
                if($ws.ProtectContents){throw 'Transactional DB is protected; this probe does not request or bypass passwords.'}
                $used=$ws.UsedRange;$usedColumns=$used.Columns;$right=$used.Column+$usedColumns.Count-1;Release-Com $usedColumns;Release-Com $used
                $items=@()
                foreach($letter in [char[]]'ABCDEFGHIJKLMN') {
                    $name='TransCopy_DevptSingle_'+$letter;$dn=$book.Names.Item($name);$range=$dn.RefersToRange
                    $items += [pscustomobject]@{Name=$name;Row=$range.Row};Release-Com $range;Release-Com $dn
                }
                foreach($item in ($items | Sort-Object Row -Descending)) {
                    $dn=$book.Names.Item($item.Name);$range=$dn.RefersToRange
                    $rows=$range.Rows;$columns=$range.Columns
                    $top=$range.Row;$left=$range.Column;$bottom=$top+$rows.Count-1;$width=$columns.Count
                    Release-Com $rows;Release-Com $columns;Release-Com $range
                    $cells=$ws.Cells;$a=$cells.Item($bottom,1);$b=$cells.Item($bottom+$Count-1,$right);$area=$ws.Range($a,$b)
                    $part=[Diagnostics.Stopwatch]::StartNew();[void]$area.Insert(-4121);$shiftMs+=$part.ElapsedMilliseconds
                    Release-Com $area;Release-Com $a;Release-Com $b
                    $a=$cells.Item($bottom-1,$left);$b=$cells.Item($bottom-1,$left+$width-1);$template=$ws.Range($a,$b);Release-Com $a;Release-Com $b
                    $a=$cells.Item($bottom,$left);$b=$cells.Item($bottom+$Count-1,$left+$width-1);$target=$ws.Range($a,$b);Release-Com $a;Release-Com $b
                    $part.Restart();$template.Copy($target);$copyMs+=$part.ElapsedMilliseconds;Release-Com $template;Release-Com $target
                    $a=$cells.Item($top,$left);$b=$cells.Item($bottom+$Count,$left+$width-1);$expanded=$ws.Range($a,$b)
                    $dn.RefersTo="='Transactional DB'!"+$expanded.Address();Release-Com $expanded;Release-Com $a;Release-Com $b;Release-Com $cells;Release-Com $dn
                    $completed++
                }
                Release-Com $ws
            } else {
                # Read the reviewed native target group; do not duplicate a new
                # independent list of workbook sheet names in the benchmark.
                $ruleText=Get-Content (Join-Path $repo 'Services/WorkbookStructureRuleManager.vb') -Raw
                if($operation -eq 'FundingColumns') {
                    $block=[regex]::Match($ruleText,'(?s)AddColumnTargets\(FundingRule,(.*?)\)\s*\x27Funding_Columns').Groups[1].Value
                    $sheets=@([regex]::Matches($block,'"([^"]+)"') | ForEach-Object {$_.Groups[1].Value})
                    if($sheets.Count -ne 32){throw 'Reviewed Funding group could not be resolved'}
                    $anchor='LoanDescRev1';$offset=-1;$batches=@($Count)
                } else {
                    $sheets=@('Development BP Assumptions','Development Stock','Development Capital','Development Revenue','Development Expenditure','Dvpt NonCash','Dvpt Component Depn')
                    $anchor='LastIDColNum';$offset=0;$batches=if($Count -gt 1){@(1,($Count-1))}else{@(1)}
                }
                foreach($sheet in $sheets){$ws=$book.Worksheets.Item($sheet);if($ws.ProtectContents){throw ('Protected sheet: '+$sheet+'; no passwords requested or bypassed')};$used=$ws.UsedRange;$cols=$used.Columns;$beforeWidths[$sheet]=$cols.Count;Release-Com $cols;Release-Com $used;Release-Com $ws}
                $dn=$book.Names.Item($anchor);$range=$dn.RefersToRange;$at=$range.Column+$offset;Release-Com $range;Release-Com $dn
                foreach($batch in $batches) {
                    $first=$true;foreach($sheet in $sheets){$ws=$book.Worksheets.Item($sheet);$ws.Select($first);$first=$false;Release-Com $ws}
                    $ws=$book.Worksheets.Item($sheets[0]);$ws.Activate();$cols=$ws.Columns;$a=$cols.Item($at);$b=$cols.Item($at+$batch-1);$area=$ws.Range($a,$b)
                    $area.Select();$selection=$excel.Selection;$part=[Diagnostics.Stopwatch]::StartNew();[void]$selection.Insert(-4161);$shiftMs+=$part.ElapsedMilliseconds
                    Release-Com $selection;Release-Com $area;Release-Com $a;Release-Com $b;Release-Com $cols
                    $ws.Select();Release-Com $ws
                    foreach($sheet in $sheets) {
                        $ws=$book.Worksheets.Item($sheet);$cols=$ws.Columns
                        $template=$cols.Item($(if($operation -eq 'FundingColumns'){$at+$batch}else{$at-1}))
                        $a=$cols.Item($at);$b=$cols.Item($at+$batch-1);$target=$ws.Range($a,$b)
                        $part.Restart();$template.Copy($target);$copyMs+=$part.ElapsedMilliseconds
                        Release-Com $target;Release-Com $a;Release-Com $b;Release-Com $template;Release-Com $cols;Release-Com $ws
                    }
                    $at+=$batch
                }
                $completed=$sheets.Count
                foreach($sheet in $sheets){$ws=$book.Worksheets.Item($sheet);$used=$ws.UsedRange;$cols=$used.Columns;$widths += [pscustomobject]@{Sheet=$sheet;Before=$beforeWidths[$sheet];After=$cols.Count;Delta=($cols.Count-$beforeWidths[$sheet])};Release-Com $cols;Release-Com $used;Release-Com $ws}
            }
        } catch {$errorText=$_.Exception.Message}
        $results += [pscustomobject]@{Operation=$operation;Count=$Count;TotalMs=$watch.ElapsedMilliseconds;ShiftMs=$shiftMs;CopyMs=$copyMs;TargetsCompleted=$completed;Error=$errorText;SourceSha256=$hash;Saved=$false;FullCommandValidated=$false;UsedWidthChanges=$widths}
        $results | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $out 'results.json')
        Write-Output ($results[-1] | ConvertTo-Json -Depth 4 -Compress)
        $book.Close($false);Release-Com $book;$book=$null
    }
}finally{
    if($book){$book.Close($false);Release-Com $book};if($seed){$seed.Close($false);Release-Com $seed};Release-Com $books
    if($excel){if($owned){$excel.Quit()};Release-Com $excel}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if((Get-FileHash -LiteralPath $source).Hash -ne $hash){throw 'Source hash changed during probe; source is never opened or saved by this script'}
}
Write-Output 'PASS: Isolated core probes finished; all changes discarded, no VBA or calculation executed. Not financial or complete-command validation.'
