param([string]$Workbook=(Join-Path (Split-Path -Parent $PSScriptRoot) 'Library/Demo BP v26_0001.xlsb'))
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$out=Join-Path $repo ('obj/ClientInputExcel/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$copy=Join-Path $out 'private-readonly.xlsb'
$hash=(Get-FileHash -LiteralPath $Workbook).Hash
Copy-Item -LiteralPath $Workbook -Destination $copy
function Release-Com($obj){if($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj)}}
$excel=$null;$seed=$null;$book=$null
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false
 $seed=$excel.Workbooks.Add();$excel.Calculation=-4135
 $book=$excel.Workbooks.Open($copy,0,$true)
 $targets=@{
  'Service Charge Assumptions'=@('E19')
  'Management Costs Assumptions'=@('D19','D20','D25','D35')
  'Repairs & Maint. Assumptions'=@('C50','D58','D55')
  'Stock Condition Inputs'=@('D9')
  'Development BP Assumptions'=@('O12','O13','O169','O170')
  'Dvpt BP Rev and Exp Assumptions'=@('D8','D16','D25','D34')
  'Economic Assumptions'=@('D9','D31','D103','D207','D237')
  'Funding Assumptions'=@('E121','E131','E137','B230','C230','D230','B231','C231','D231','B237','C237','D237','B238','C238','D238')
  'Covenant Assumptions'=@('A8','B8','C8','D8','E8','F8','G8','H8')
  'Housing Asset Assumptions'=@('B14','D14','B59','D59','B61','D61','B63','D63','B69','D69','B84','D84','B96','D96')
 }
 $results=foreach($sheetName in $targets.Keys){
  $ws=$null
  try{$ws=$book.Worksheets.Item($sheetName)}catch{Write-Output "SHEET_MISSING=$sheetName";continue}
  try {foreach($address in $targets[$sheetName]){
   $cell=$ws.Range($address)
   try{
    $validation=$null;try{$validation=[pscustomobject]@{Type=$cell.Validation.Type;Operator=$cell.Validation.Operator;Formula1=$cell.Validation.Formula1;Formula2=$cell.Validation.Formula2}}catch{}
    $rules=@();for($i=1;$i -le $cell.FormatConditions.Count;$i++){
     $rule=$cell.FormatConditions.Item($i)
     try{$rules+= [pscustomobject]@{Type=$rule.Type;Formula1=$rule.Formula1;AppliesTo=$rule.AppliesTo.Address();Interior=$rule.Interior.Color;Font=$rule.Font.Color}}finally{Release-Com $rule}
    }
    [pscustomobject]@{Sheet=$sheetName;Cell=$address;Value=$cell.Value2;Text=$cell.Text;NumberFormat=$cell.NumberFormat;Locked=$cell.Locked;Formula=$cell.HasFormula;FormulaText=$(if($cell.HasFormula){$cell.Formula}else{$null});DirectFill=$cell.Interior.Color;DisplayedFill=$cell.DisplayFormat.Interior.Color;DirectFont=$cell.Font.Color;DisplayedFont=$cell.DisplayFormat.Font.Color;Validation=$validation;Rules=$rules}
   }finally{Release-Com $cell}
  }}finally{Release-Com $ws}
 }
 $results | ConvertTo-Json -Depth 8 | Out-File (Join-Path $out 'excel-observations.json') -Encoding utf8
 Write-Output ('OUTPUT='+$out)
 Write-Output ('CELLS='+$results.Count)
}finally{
 if($book){$book.Close($false);Release-Com $book};if($seed){$seed.Close($false);Release-Com $seed};if($excel){$excel.Quit();Release-Com $excel}
 if((Get-FileHash -LiteralPath $Workbook).Hash -ne $hash){throw 'Source workbook changed'}
 Write-Output 'Source hash unchanged; no macros/events/link updates/recalculation/save requested'
}
