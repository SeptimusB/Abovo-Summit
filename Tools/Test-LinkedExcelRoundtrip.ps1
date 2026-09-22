param([Parameter(Mandatory=$true)][string]$TestOutput,[Parameter(Mandatory=$true)][string]$Rule,[int]$Count=3,[switch]$CompareOnly)
$ErrorActionPreference='Stop'
$TestOutput=(Resolve-Path -LiteralPath $TestOutput).Path
$inputFile=Join-Path $TestOutput 'excel-reference-input.xlsb';$expanded=Join-Path $TestOutput 'expanded.xlsb'
$expected=Join-Path $TestOutput 'excel-grouped-reference.xlsb';$roundtrip=Join-Path $TestOutput 'excel-roundtrip.xlsb'
$rows=$false;$rangeName=$null;$anchor=$null;$linked=@();$sheets=@()
# Independent worksheet/event contracts transcribed from the inspected master.
switch($Rule){
 'SIMPLE_REP_SERVCHG_02' {$rows=$true;$rangeName='IR_ServChg_01';$anchor='LastUnitSCColumn';$sheets=@('Service Charge Assumptions');$linked=@('Unit Service Charges','Service Charge Numbers','Service Charge Income','Service Charge Voids','Service Charge Bad Debts')}
 'SIMPLE_IR_SPEC_INC_ASS1' {$rows=$true;$rangeName='IR_Spec_Inc_Ass1';$anchor='LastSIDriversCol';$sheets=@('Specific Income Assumptions');$linked=@('Specific Income Drivers','Specific Income Workings','Specific Income Voids','Specific Income Bad Debts')}
 'OTHER_INCOME_RECORDS' {$rows=$true;$rangeName='IR_Oth_Inc_Ass5';$anchor='LastOIWorkingsCol';$sheets=@('Other Income Assumptions');$linked=@('Other Income Workings')}
 'JOINT_VENTURE_RECORDS' {$rangeName='IC_JointVenture_01';$anchor='LastJointVentureCol';$sheets=@('Joint Venture Assumptions','JV Opening balances','JV Investments','JV Share of Profits','JV Payments Repayments','JV Closing Balances','JV Interest Received')}
 {$_ -in 'INTERCO_LOAN_RECORDS','INTERCO_INVESTMENT_RECORDS'} {
  $rangeName=if($Rule -eq 'INTERCO_LOAN_RECORDS'){'IC_IntercoFunding_01'}else{'IC_IntercoFunding_02'}
  $anchor=if($Rule -eq 'INTERCO_LOAN_RECORDS'){'LastIntercoLoanAssCol'}else{'LastIntercoInvestAssCol'}
  $sheets=@('Interco Funding Assumptions','Hidden - InterCo Int Rates','Hidden - InterCo Opening Bal','Hidden - InterCo Increases','Hidden - InterCo Decreases','Hidden - InterCo Closing Bal','Hidden - InterCo Interest','InterCo Opening Balances','InterCo Increases','InterCo Decreases','InterCo Closing Balances','InterCo Interest')
 }
 'SPECIFIC_INCOME_CATEGORIES' {$rangeName='Rep_SInc_01';$anchor='LastSIAssumpCol';$sheets=@('Specific Income Assumptions')}
 'OTHER_INCOME_CATEGORIES' {$rangeName='Rep_OInc_04';$anchor='LastOIAssumpCol';$sheets=@('Other Income Assumptions')}
 default {throw 'Unsupported independent contract'}
}
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}}
function Clear-EditableConstants($range){
 # Intentional native safety policy: retain formula-backed editable cells.
 try{$constants=$range.SpecialCells(2)}catch{return}
 try{foreach($cell in $constants.Cells){try{if(!$cell.Locked){[void]$cell.ClearContents()}}finally{Release-Com $cell}}}finally{Release-Com $constants}
}
if(!$CompareOnly){
 if((Test-Path $expected) -or (Test-Path $roundtrip)){throw 'Never overwrite reference results'}
 $excel=$null;$seed=$null;$book=$null
 try{
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false
  $seed=$excel.Workbooks.Add();$excel.Calculation=-4135
  $book=$excel.Workbooks.Open($inputFile,0,$true)
  if($rows){
   $nr=$book.Names.Item($rangeName).RefersToRange;$ws=$nr.Worksheet;$ws.Select()
   $firstRow=$nr.Row;$n=$nr.Rows.Count;$at=$firstRow+$n;$lastCol=$ws.UsedRange.Column+$ws.UsedRange.Columns.Count-1
   # VBA Row_Insertion: first row, copy second-last existing row, clear inputs;
   # then insert the remaining rows and copy the new blank first row.
   [void]$ws.Rows.Item($at).Insert(-4121)
   $ws.Rows.Item($at-2).Copy($ws.Rows.Item($at))
   $new=$ws.Range($ws.Cells.Item($at,1),$ws.Cells.Item($at,$lastCol));Clear-EditableConstants $new;Release-Com $new
   if($Count -gt 1){
    $new=$ws.Range($ws.Rows.Item($at+1),$ws.Rows.Item($at+$Count-1));[void]$new.Insert(-4121);Release-Com $new
    $new=$ws.Range($ws.Rows.Item($at+1),$ws.Rows.Item($at+$Count-1));$ws.Rows.Item($at).Copy($new);Release-Com $new
   }
   $resized=$ws.Range($ws.Cells.Item($firstRow,$nr.Column),$ws.Cells.Item($firstRow+$n+$Count-1,$nr.Column+$nr.Columns.Count-1))
   $book.Names.Item($rangeName).RefersTo='='+$resized.Address($true,$true,1,$true)
   Release-Com $resized;Release-Com $nr;Release-Com $ws
  }
  $group=@(if($rows){$linked}else{$sheets})
  $at=$book.Names.Item($anchor).RefersToRange.Column-1
  $first=$true;foreach($name in $group){$ws=$book.Worksheets.Item($name);[void]$ws.Select($first);$first=$false;Release-Com $ws}
  $ws=$book.Worksheets.Item($group[0]);$ws.Activate()
  $new=$ws.Range($ws.Columns.Item($at),$ws.Columns.Item($at+$Count-1));$new.Select();[void]$excel.Selection.Insert(-4161);Release-Com $new
  $template=if($rows){$at-1}else{$at+$Count};$ws.Columns.Item($template).Select();$excel.Selection.Copy()
  $new=$ws.Range($ws.Columns.Item($at),$ws.Columns.Item($at+$Count));$new.Select();$excel.ActiveSheet.Paste()
  $excel.CutCopyMode=[Enum]::ToObject([Microsoft.Office.Interop.Excel.XlCutCopyMode],0);$ws.Select();Release-Com $new
  if(!$rows){
   $bottom=$ws.UsedRange.Row+$ws.UsedRange.Rows.Count-1
   $new=$ws.Range($ws.Cells.Item(1,$at+1),$ws.Cells.Item($bottom,$at+$Count));Clear-EditableConstants $new;Release-Com $new
  }
  Release-Com $ws
  $book.SaveCopyAs($expected);$book.Close($false);Release-Com $book;$book=$null
  $book=$excel.Workbooks.Open($expanded,0,$true);$excel.CalculateFullRebuild()
  $cycle=$excel.CircularReference;if($null -ne $cycle){throw ('Excel circular-reference address: '+$cycle.Parent.Name+'!'+$cycle.Address())}
  Write-Output 'INFO: Macro-disabled Excel calculation reports no circular-reference address; not a financial/VBA execution sign-off.'
  $book.SaveCopyAs($roundtrip);$book.Close($false);Release-Com $book;$book=$null
 }finally{
  if($book){$book.Close($false);Release-Com $book};if($seed){$seed.Close($false);Release-Com $seed};if($excel){$excel.Quit();Release-Com $excel}
 }
}
& (Join-Path $PSScriptRoot 'Test-DevelopmentExcelRoundtrip.ps1') -TestOutput $TestOutput -CompareOnly -ComparisonSheets ($sheets+$linked) -Count $Count
Write-Output ('PASS: '+$Rule+' independent formula/array comparison and Excel round-trip')
