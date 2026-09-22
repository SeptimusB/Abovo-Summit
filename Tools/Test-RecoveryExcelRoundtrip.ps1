param([Parameter(Mandatory=$true)][string]$Recovery,[Parameter(Mandatory=$true)][string]$Normal)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$out=Join-Path $repo ('obj/RecoveryExcelRoundtrip/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$hashes=@{};foreach($file in @($Recovery,$Normal)){$hashes[$file]=(Get-FileHash -LiteralPath $file).Hash}
function Release-Com($item){if($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($item)}}
$excel=$null;$seed=$null;$books=$null;$book=$null
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false
 $books=$excel.Workbooks;$seed=$books.Add();$excel.Calculation=-4135;$excel.CalculateBeforeSave=$false
 foreach($file in @($Recovery,$Normal)){
  $book=$books.Open($file,0,$true)
  $sheets=$book.Worksheets;$sheet=$sheets.Item('Global Assumptions')
  if($book.RepairMode){throw 'Excel entered repair mode'}
  Write-Output ('PASS: Excel opened without repair: '+[IO.Path]::GetFileName($file)+', sheets='+$sheets.Count)
  Release-Com $sheet;Release-Com $sheets
  $parts=$book.CustomXMLParts;$found=0
  for($i=1;$i -le $parts.Count;$i++){$part=$parts.Item($i);try{if($part.NamespaceURI -eq 'urn:abovo:summit:recovery-history:1'){$found++}}finally{Release-Com $part}}
  Release-Com $parts;if($found -ne 1){throw ('Expected one recovery history part; found '+$found)}
  $copy=Join-Path $out ('excel-'+[IO.Path]::GetFileName($file));$book.SaveCopyAs($copy)
  Write-Output ('COPY='+$copy)
  $book.Close($false);Release-Com $book;$book=$null
 }
}finally{
 if($book){$book.Close($false);Release-Com $book}
 if($seed){$seed.Close($false);Release-Com $seed}
 Release-Com $books
 if($excel){$excel.Quit();Release-Com $excel}
 foreach($file in $hashes.Keys){if((Get-FileHash -LiteralPath $file).Hash -ne $hashes[$file]){throw 'Test input modified'}}
}
Write-Output ('PASS: read-only Excel load and SaveCopyAs. Macros, events, link updates and calculation were disabled. OUTPUT='+$out)
