using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;

static class PresentationTests
{
    static int assertions;
    static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("PRESENTATION PASS "+(++assertions)+" "+text);}
    internal static async Task Run(string source)
    {
        var before=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).OrderBy(x=>x).ToArray();
        var root=Path.Combine(Path.GetDirectoryName(source),"presentation-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);
        foreach(bool date1904 in new[]{false,true})
        {
            var directory=Path.Combine(root,date1904?"1904":"1900");Directory.CreateDirectory(directory);
            var path=SaveCandidateTests.Fixture(directory,".xlsx");
            using(var wb=new Workbook())
            {
                wb.LoadDocument(path);wb.DocumentSettings.Calculation.Use1904DateSystem=date1904;
                var sheet=wb.Worksheets[0];sheet.Cells["A4"].Value=CellValue.FromDateTime(new DateTime(2026,9,30),date1904);sheet.Cells["A4"].NumberFormat="dd-mmm-yyyy";
                sheet.Cells["B4"].Value=-0.125d;sheet.Cells["B4"].NumberFormat="0.00%";
                sheet.Cells["C4"].Value="=literal";sheet.Cells["D4"].Formula="=1/0";
                sheet.Cells["B1"].Font.Bold=true;sheet.Cells["B1"].Font.Italic=true;sheet.Cells["B1"].Font.Size=14;
                sheet.Cells["B1"].Alignment.Horizontal=SpreadsheetHorizontalAlignment.Right;
                sheet.Cells["B1"].Alignment.WrapText=true;
                wb.CalculateFull();wb.SaveDocument(path,DocumentFormat.Xlsx);
            }
            foreach(var engine in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired})
            {
                var s=await WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(engine,false,false,120000,true));
                try
                {
                    var area=new WorkbookReadArea("Data",0,0,4,4);var input=new WorkbookReadArea("Data",0,0,1,1);
                    var result=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{area},includePresentation:true);
                    var p=result.Presentation[0];
                    Check(p.CellAt(3,0).Text=="30-Sep-2026",s.EngineName+" date formatting respects "+(date1904?1904:1900)+" date system: "+p.CellAt(3,0).Text);
                    Check(p.CellAt(3,1).Text=="-12.50%"&&p.CellAt(3,2).Text=="=literal"&&p.CellAt(3,3).Text=="#DIV/0!","percent, literal text and typed error presentation");
                    var font=p.CellAt(0,1).Appearance;
                    Check(font.FontSize==14&&font.FontStyle.HasFlag(FontStyle.Bold)&&font.FontStyle.HasFlag(FontStyle.Italic)&&font.Horizontal==SpreadsheetHorizontalAlignment.Right&&font.WrapText,"workbook font/alignment/wrap retained");
                    Check(p.CellAt(0,0).Appearance.SolidFill,"base input fill retained before condition");
                    var again=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{area},includePresentation:true);
                    Check(!s.IsCurrent(result)&&s.IsCurrent(again)&&again.CalculationGeneration>result.CalculationGeneration,"new calculation invalidates earlier appearance even at same input revision");
                    var snapshot=await s.CaptureCellAsync(s.Revision,input);
                    var edit=await s.ApplyValuesAsync(new[]{new WorkbookValueChange(snapshot,310d,WorkbookValuePermission.UnlockedCell)},new[]{area},includePresentation:true);
                    var effective=edit.Results.Presentation[0].CellAt(0,0);
                    Check(!effective.Appearance.SolidFill&&effective.Appearance.Foreground.ToArgb()==Color.DarkRed.ToArgb()&&effective.Text=="310.00",s.EngineName+" effective conditional fill/font and formatted amount published together");
                    Check(Equals(edit.Results.Blocks[0].ValueAt(0,1),620d)&&s.IsCurrent(edit.Results),"dependent value belongs to same accepted snapshot");
                    Check(!s.IsCurrent(again),"edit invalidates previous display");
                    Console.WriteLine("PRESENTATION_NATIVE engine="+s.EngineName+" calculationMs="+edit.Results.CalculationMilliseconds+" transferMs="+edit.Results.TransferMilliseconds);
                }
                finally{await s.CloseAsync();}
            }
        }
        var timer=Stopwatch.StartNew();int[] after;
        do{await Task.Delay(100);after=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).OrderBy(x=>x).ToArray();}while(timer.ElapsedMilliseconds<10000&&!before.SequenceEqual(after));
        Check(before.SequenceEqual(after),"owned Excel exits and user instances remain");
        Console.WriteLine("PRESENTATION ASSERTIONS="+assertions);
    }
}
