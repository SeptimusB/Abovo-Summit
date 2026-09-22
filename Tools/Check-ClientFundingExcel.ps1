param([Parameter(Mandatory=$true)][string]$Source,[Parameter(Mandatory=$true)][string]$Candidate)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot;$out=Join-Path $repo ('obj/ClientFundingExcel/'+[guid]::NewGuid().ToString('N'));[void](New-Item -ItemType Directory -Path $out)
$hashes=@{};foreach($file in @($Source,$Candidate)){$hashes[$file]=(Get-FileHash -LiteralPath $file).Hash}
function Release-Com($x){if($null -ne $x -and [Runtime.InteropServices.Marshal]::IsComObject($x)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)}}
$excel=$null;$book=$null;$seed=$null
try {
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false
 $seed=$excel.Workbooks.Add();$excel.Calculation=-4135;$excel.CalculateBeforeSave=$false
 Write-Output ('OUTPUT='+$out)
 foreach($file in @($Source,$Candidate)){
  $book=$excel.Workbooks.Open($file,0,$true)
  Write-Output ('FILE='+$file)
  foreach($sn in @('Funding Assumptions','Hidden - Loan Interest','Transactional DB')){$ws=$book.Worksheets.Item($sn);Write-Output ($sn+' protected='+$ws.ProtectContents);Release-Com $ws}
  $watch=[Diagnostics.Stopwatch]::StartNew();$excel.CalculateFullRebuild();Write-Output ('Excel full rebuild ms='+$watch.ElapsedMilliseconds)
  $cycle=$excel.CircularReference;if($cycle){Write-Output ('CIRCULAR='+$cycle.Parent.Name+'!'+$cycle.Address());Release-Com $cycle;if($file -eq $Candidate){throw 'Repaired workbook has a circular reference'}}else{Write-Output 'No circular-reference address returned (macros disabled).'}
  $ws=$book.Worksheets.Item('Check Sheet')
  $checks=@();foreach($cell in $ws.Range('A6:H63').Cells){if($cell.Column -in 2,4,5){$checks+=@{Cell=$cell.Address($false,$false);Value=$cell.Text;Formula=$cell.Formula}};Release-Com $cell}
  $checks | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $out ($(if($file -eq $Source){'original'}else{'repaired'})+'-checks.json')) -Encoding UTF8
  foreach($address in @('A1','B37','E37','B39','E39','B63')){$cell=$ws.Range($address);Write-Output ($address+'='+$cell.Text);Release-Com $cell};Release-Com $ws
  if($file -eq $Candidate){
   $ok=@($checks|Where-Object {$_.Value -eq 'OK'}).Count
   if($ok -ne 43 -or @($checks|Where-Object {$_.Cell -eq 'B63' -and $_.Value -eq '0'}).Count -ne 1){throw 'Expected 43 OK check statuses and B63=0'}
   $ws=$book.Worksheets.Item('Multivariable Dashboard');$cell=$ws.Range('B41')
   if(-not $cell.HasFormula -or $cell.Formula -eq '=#VALUE!' -or $cell.Text -like '#*'){throw 'Dashboard description formula failed Excel recalculation'}
   Write-Output ('PASS: 43 check statuses OK; dashboard B41='+$cell.Text);Release-Com $cell;Release-Com $ws
   $book.SaveCopyAs((Join-Path $out 'excel-recalculated.xlsb'))
  }
  $book.Close($false);Release-Com $book;$book=$null
 }
}finally {if($book){$book.Close($false);Release-Com $book};if($seed){$seed.Close($false);Release-Com $seed};if($excel){$excel.Quit();Release-Com $excel};foreach($file in $hashes.Keys){if((Get-FileHash -LiteralPath $file).Hash -ne $hashes[$file]){throw ('Source modified: '+$file)}}}
