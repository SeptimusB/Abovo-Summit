param([Parameter(Mandatory=$true)][string]$SavedWorkbook,[Parameter(Mandatory=$true)][string]$ReferenceWorkbook)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$out=Join-Path $repo ('obj/SaveExcelRoundtrip/'+[guid]::NewGuid().ToString('N'));[void](New-Item -ItemType Directory -Path $out)
$roundtrip=Join-Path $out 'excel-roundtrip.xlsb'
$hashes=@{};foreach($file in @($SavedWorkbook,$ReferenceWorkbook)){$hashes[$file]=(Get-FileHash -LiteralPath $file).Hash}
function Release-Com($item){if($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($item)}}
$excel=$null;$seed=$null;$book=$null;$data=@()
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false
 $seed=$excel.Workbooks.Add();$excel.Calculation=-4135;$excel.CalculateBeforeSave=$false
 foreach($file in @($ReferenceWorkbook,$SavedWorkbook)){
  $book=$excel.Workbooks.Open($file,0,$true);$record=@{}
  foreach($sheet in @('Detailed Comp Inc - Trad View','Financial Position - Trad View','Cashflow detailed','Check Sheet')){
   $ws=$book.Worksheets.Item($sheet);$range=$ws.UsedRange;$record[$sheet]=$range.Value2;Release-Com $range;Release-Com $ws
  }
  $ws=$book.Worksheets.Item('Multivariable Dashboard');$cell=$ws.Range('B41');Write-Output ([IO.Path]::GetFileName($file)+' B41='+$cell.Formula);Release-Com $cell;Release-Com $ws
  $data+=,$record
  if($file -eq $SavedWorkbook){$book.SaveCopyAs($roundtrip)}
  $book.Close($false);Release-Com $book;$book=$null
 }
 foreach($sheet in $data[0].Keys){
  $a=$data[0][$sheet];$b=$data[1][$sheet];$differences=0
  if($a.GetLength(0) -ne $b.GetLength(0) -or $a.GetLength(1) -ne $b.GetLength(1)){throw ('Output dimensions differ: '+$sheet)}
  for($r=1;$r -le $a.GetLength(0);$r++){for($c=1;$c -le $a.GetLength(1);$c++){
   $x=$a[$r,$c];$y=$b[$r,$c]
   if($x -is [double] -and $y -is [double]){if([Math]::Abs($x-$y) -le 1e-7){continue}}elseif([object]::Equals($x,$y)){continue}
   $differences++;if($differences -le 4){Write-Output ('DIFFERENCE '+$sheet+' R'+$r+'C'+$c+' reference='+$x+' saved='+$y)}
  }}
  Write-Output ($sheet+' differences='+$differences)
  if($differences -ne 0){throw ('Saved outputs differ: '+$sheet)}
 }
 Write-Output ('PASS: Read-only Excel output comparison; macro-free copy saved at '+$roundtrip)
}finally{
 if($book){$book.Close($false);Release-Com $book}
 if($seed){$seed.Close($false);Release-Com $seed}
 if($excel){$excel.Quit();Release-Com $excel}
 foreach($file in $hashes.Keys){if((Get-FileHash -LiteralPath $file).Hash -ne $hashes[$file]){throw ('Source modified: '+$file)}}
}
