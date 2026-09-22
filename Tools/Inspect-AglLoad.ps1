param([string]$InputDirectory='C:/Sandbox/Insert Comp',[switch]$SaveProbe)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot;$bin=Join-Path $repo 'bin/Release'
$refs=@('System.Core','System.Drawing','System.Data')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Docs.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
foreach($ref in $refs | Where-Object {Test-Path -LiteralPath $_}){[void][Reflection.Assembly]::LoadFrom($ref)}
Add-Type -ReferencedAssemblies $refs -TypeDefinition @'
using System;
using System.IO;
using System.Linq;
using DevExpress.Spreadsheet;
public static class AglLoadProbe {
 public static void Inspect(string path,string privateCopy){
  using(var w=new Workbook()){
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;
   using(var stream=File.OpenRead(path)){w.LoadDocument(stream,DocumentFormat.Xlsb);}
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;
   Console.WriteLine(Path.GetFileName(path));
   var cell=w.Worksheets["Multivariable Dashboard"].Cells["B41"];
   Console.WriteLine("Dashboard B41 formula immediately after load: "+cell.FormulaInvariant);
   foreach(string name in new[]{"Detailed Comp Inc - Trad View","Financial Position - Trad View"}){
    var s=w.Worksheets[name];int row=name.StartsWith("Detailed")?72:25;
    Console.WriteLine(name+" row label: "+String.Join(" | ",Enumerable.Range(0,9).Select(i=>s.Cells[row,i].Value.ToString())));
    var c=s.Cells[row,9];Console.WriteLine(c.GetReferenceA1()+" = "+c.FormulaInvariant+"; saved value="+c.Value.ToString());
   }
   foreach(var item in new[]{new[]{"Cashflow detailed","BO9"},new[]{"Cashflow detailed","BR9"}}){
    var c=w.Worksheets[item[0]].Cells[item[1]];Console.WriteLine(item[0]+"!"+item[1]+": "+c.FormulaInvariant+"; saved value="+c.Value.ToString());
   }
   if(!String.IsNullOrEmpty(privateCopy)){
    if(File.Exists(privateCopy))throw new Exception("Private evidence file already exists");
    w.SaveDocument(privateCopy,DocumentFormat.Xlsb);
    using(var check=new Workbook()){
     check.Options.CalculationMode=WorkbookCalculationMode.Manual;
     using(var stream=File.OpenRead(privateCopy)){check.LoadDocument(stream,DocumentFormat.Xlsb);}
     Console.WriteLine("Dashboard B41 after no-edit private save/reopen: "+check.Worksheets["Multivariable Dashboard"].Cells["B41"].FormulaInvariant);
    }
   }
  }
 }
}
'@
$prefix='AGL - BP 2627 updated Mar26 v25 v2 v26_0005'
$privateCopy=$null
if($SaveProbe){$out=Join-Path $repo ('obj/AglNoEditRoundtrip/'+[guid]::NewGuid().ToString('N'));[void](New-Item -ItemType Directory -Path $out);$privateCopy=Join-Path $out 'native-no-edit.xlsb';Write-Output ('Private output: '+$privateCopy)}
$suffixes=if($SaveProbe){@('.xlsb')}else{@('.xlsb',' 10 funding records in summit.xlsb',' 10 funding cols added in excel.xlsb')}
foreach($suffix in $suffixes){
 $file=Join-Path $InputDirectory ($prefix+$suffix);$hash=(Get-FileHash -LiteralPath $file).Hash
 try{[AglLoadProbe]::Inspect($file,$privateCopy)}finally{if((Get-FileHash -LiteralPath $file).Hash -ne $hash){throw 'Source changed'}}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
Write-Output 'PASS: No macros or explicit calculation; original hashes unchanged; private save only if explicitly requested'
