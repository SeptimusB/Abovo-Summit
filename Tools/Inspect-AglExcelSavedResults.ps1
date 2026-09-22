param([string]$InputDirectory='C:/Sandbox/Insert Comp')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$out=Join-Path $repo ('obj/AglExcelSavedResults/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$prefix='AGL - BP 2627 updated Mar26 v25 v2 v26_0005'
$suffixes=@('.xlsb',' 10 funding records in summit.xlsb',' 10 funding cols added in excel.xlsb',' 10 development lines added in summit.xlsb',' 10 development cols added in excel.xlsb')
$files=@($suffixes | ForEach-Object {Join-Path $InputDirectory ($prefix+$_)})
$hashes=@{};foreach($file in $files){$hashes[$file]=(Get-FileHash -LiteralPath $file).Hash}
$sheets=@('Summary Comp Inc - Trad View','Detailed Comp Inc - Trad View','Financial Position - Trad View','Financial Position - Alt View','Cashflow detailed','Check Sheet')
Add-Type -ReferencedAssemblies System.Core,System.Web.Extensions -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Script.Serialization;
public static class SavedResults {
 public static object Diff(object left,object right) {
  var a=left as object[,];var b=right as object[,];var examples=new List<object>();
  if(a==null||b==null)throw new Exception("Expected a rectangular Excel range");
  int rows=Math.Max(a.GetLength(0),b.GetLength(0)),cols=Math.Max(a.GetLength(1),b.GetLength(1)),diff=0;
  double max=0;int numerical=0;
  for(int i=1;i<=rows;i++)for(int j=1;j<=cols;j++){
   object x=i<=a.GetLength(0)&&j<=a.GetLength(1)?a[i,j]:null,y=i<=b.GetLength(0)&&j<=b.GetLength(1)?b[i,j]:null;
   if(Equals(x,y))continue;
   if(x is double && y is double){double delta=Math.Abs((double)x-(double)y);if(delta<=1e-7)continue;max=Math.Max(max,delta);numerical++;}
   diff++;if(examples.Count<12)examples.Add(new {row=i,column=j,left=x,right=y});
  }
  return new {rows,columns=cols,differences=diff,numericalDifferences=numerical,maxAbsoluteNumericDifference=max,examples};
 }
 public static void SaveJson(string path,string value){File.WriteAllText(path,value);}
}
'@
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}}
$excel=$null;$seed=$null;$book=$null;$data=@();$formulaData=@();$protection=@()
Write-Output ('Output: '+$out)
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false
 $seed=$excel.Workbooks.Add();$excel.Calculation=-4135;$excel.CalculateBeforeSave=$false
 foreach($file in $files){
  Write-Output ('Inspecting '+[IO.Path]::GetFileName($file))
  $book=$excel.Workbooks.Open($file,0,$true)
  $bookData=@{};$formulaBook=@{};$unprotected=@()
  foreach($name in $sheets){$ws=$book.Worksheets.Item($name);$range=$ws.UsedRange;$bookData[$name]=[pscustomobject]@{Address=$range.Address();FirstRow=$range.Row;FirstColumn=$range.Column;Values=$range.Value2};Release-Com $range;Release-Com $ws}
  foreach($name in @('Development Expenditure','Multivariable Dashboard','Transactional DB')){$ws=$book.Worksheets.Item($name);$range=$ws.UsedRange;$formulaBook[$name]=[pscustomobject]@{Address=$range.Address();FirstRow=$range.Row;FirstColumn=$range.Column;Values=$range.Formula};Release-Com $range;Release-Com $ws}
  for($i=1;$i -le $book.Worksheets.Count;$i++){$ws=$book.Worksheets.Item($i);if(!$ws.ProtectContents){$unprotected+=$ws.Name};Release-Com $ws}
  $protection+=,@{File=$file;Unprotected=$unprotected};$data+=,$bookData;$formulaData+=,$formulaBook
  $book.Close($false);Release-Com $book;$book=$null
 }
 $comparisons=@()
 foreach($pair in @(@(1,2),@(3,4),@(0,1),@(0,2),@(0,3),@(0,4))){
  $results=@();foreach($name in $sheets){$a=$data[$pair[0]][$name];$b=$data[$pair[1]][$name];if($a.FirstRow -ne $b.FirstRow -or $a.FirstColumn -ne $b.FirstColumn){throw 'Output origins differ'};$results+=,@{Sheet=$name;FirstRow=$a.FirstRow;FirstColumn=$a.FirstColumn;Result=[SavedResults]::Diff($a.Values,$b.Values)}}
  $comparisons+=,@{Left=$files[$pair[0]];Right=$files[$pair[1]];Sheets=$results}
 }
 $formulaComparisons=@();foreach($pair in @(@(1,2),@(3,4),@(0,1),@(0,2))){
  $results=@();foreach($name in @('Development Expenditure','Multivariable Dashboard','Transactional DB')){$a=$formulaData[$pair[0]][$name];$b=$formulaData[$pair[1]][$name];if($a.FirstRow -ne $b.FirstRow -or $a.FirstColumn -ne $b.FirstColumn){throw 'Formula origins differ'};$results+=,@{Sheet=$name;FirstRow=$a.FirstRow;FirstColumn=$a.FirstColumn;Result=[SavedResults]::Diff($a.Values,$b.Values)}}
  $formulaComparisons+=,@{Left=$files[$pair[0]];Right=$files[$pair[1]];Sheets=$results}
 }
 [SavedResults]::SaveJson((Join-Path $out 'excel-saved-results.json'),(@{Protection=$protection;OutputComparisons=$comparisons;FormulaComparisons=$formulaComparisons} | ConvertTo-Json -Depth 20))
 Write-Output 'PASS: Read-only Excel inspection completed, macros/events/links/calculation disabled; no workbook saved'
} finally {
 if($book){$book.Close($false);Release-Com $book}
 if($seed){$seed.Close($false);Release-Com $seed}
 if($excel){$excel.Quit();Release-Com $excel}
 foreach($file in $files){if((Get-FileHash -LiteralPath $file).Hash -ne $hashes[$file]){throw ('Source changed: '+$file)}}
 Write-Output 'PASS: All five source hashes unchanged'
}
