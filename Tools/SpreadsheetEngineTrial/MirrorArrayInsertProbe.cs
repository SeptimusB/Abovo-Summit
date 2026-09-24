using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DX = DevExpress.Spreadsheet;

internal static partial class EngineBenchmark
{
    internal static int RunMirrorArrayInsertProbe(string directory)
    {
        directory=IoTrial.InTrial(directory);Check(!Directory.Exists(directory),"New insert fixture directory required.");
        Directory.CreateDirectory(directory);var cases=new List<object>();
        foreach(bool reload in new[]{false,true})
        foreach(bool observe in new[]{false,true})
        foreach(string mode in new[]{"raw","cells","per-insert","batch"})
        {
            bool guard=mode=="per-insert"||mode=="batch";
            string prefix=Path.Combine(directory,(reload?"reloaded":"memory")+(observe?"-observed":"-unobserved")+"-"+mode);
            using(var book=new DX.Workbook())
            {
                book.Options.CalculationMode=DX.WorkbookCalculationMode.Manual;
                book.Options.CalculationEngineType=DX.CalculationEngineType.Recursive;
                var sheet=book.Worksheets[0];sheet.Name="InsertProbe";
                for(int row=1;row<13;row++)
                {
                    sheet.Cells[row,0].Value=1;
                    if(row>5&&row<8)continue;
                    for(int col=2;col<8;col++)sheet.Cells[row,col].DynamicArrayFormulaInvariant="=-INDEX($K$1:$K$2,$A"+(row+1)+")";
                }
                sheet.Cells["K1"].Value=10;sheet.Cells["K2"].Value=20;book.CalculateFullRebuild();
                if(reload)
                {
                    using(var f=new FileStream(prefix+"-input.xlsm",FileMode.CreateNew,FileAccess.Write))book.SaveDocument(f,DX.DocumentFormat.Xlsm);
                    Check(book.LoadDocument(prefix+"-input.xlsm"),"Synthetic reload failed.");
                    book.Options.CalculationMode=DX.WorkbookCalculationMode.Manual;sheet=book.Worksheets[0];
                }
                Func<object> state=()=>new {
                    collection=sheet.DynamicArrayFormulas.Cast<DX.DynamicArrayFormula>().Select(a=>a.Range.GetReferenceA1()).ToArray(),
                    cells=Enumerable.Range(5,18).Select(row=>new {address=sheet.Cells[row,2].GetReferenceA1(),formula=sheet.Cells[row,2].FormulaInvariant,dynamic=sheet.Cells[row,2].HasDynamicArrayFormula}).ToArray()};
                var before=observe?state():null;
                MirrorDynamicArrayBatch batch=mode=="batch"?new MirrorDynamicArrayBatch(sheet):null;
                if(batch!=null)batch.InsertRows(6,10);
                else if(mode=="cells")sheet.InsertCells(sheet.Range["A7:K16"],DX.InsertCellsMode.ShiftCellsDown);
                else if(guard)InsertMirrorRowsPreservingDynamicArrays(sheet,6,10);
                else sheet.Rows.Insert(6,10);
                var inserted=observe?state():null;
                if(batch!=null){batch.CopyRow(sheet.Range["A6:H6"],sheet.Range["A7:H16"]);batch.Complete();}
                else sheet.Range["A7:H16"].CopyFrom(sheet.Range["A6:H6"],DX.PasteSpecial.All);
                var copied=state();book.CalculateFullRebuild();var calculated=state();
                if(guard)AssertMirrorInsertFixture(book);
                using(var f=new FileStream(prefix+"-copied.xlsm",FileMode.CreateNew,FileAccess.Write))book.SaveDocument(f,DX.DocumentFormat.Xlsm);
                var expectedDeclarations=new Dictionary<string,string>();
                foreach(int row in Enumerable.Range(1,15).Concat(Enumerable.Range(18,5)))for(int col=2;col<8;col++)expectedDeclarations.Add(ArrayReference(row,col),ArrayReference(row,col));
                bool exportRejected=false;
                try{AssertTrialArrayDeclarations(prefix+"-copied.xlsm",sheet.Name,expectedDeclarations);}
                catch(InvalidOperationException ex){exportRejected=ex.Message.StartsWith("Candidate rejected: dynamic-array declaration differs");}
                if(mode!="cells")Check(exportRejected!=guard,"Export guard did not distinguish the incomplete and corrected candidate.");
                if(guard)
                {
                    using(var reopened=new DX.Workbook())
                    {
                        Check(reopened.LoadDocument(prefix+"-copied.xlsm"),"Guarded fixture reload failed.");
                        reopened.CalculateFullRebuild();AssertMirrorInsertFixture(reopened);
                        reopened.Worksheets[0].Cells["A19"].Value=2;reopened.CalculateFullRebuild();
                        Check(reopened.Worksheets[0].Cells["C19"].Value.NumericValue==-20,"Shifted formula did not respond to later input.");
                    }
                    using(var gear=NewMirrorProjection(false))
                    {
                        gear.Open(prefix+"-copied.xlsm");gear.Calculate(2);
                        Check(SameValue(gear.Get("InsertProbe",18,2,1,1)[0,0],-10.0),"Gear read-back differs.");
                    }
                }
                cases.Add(new {reload,observe,guard,mode,before,inserted,copied,calculated,exportRejected});
            }
        }
        using(var book=new DX.Workbook())
        {
            var sheet=book.Worksheets[0];sheet.Cells["B3"].DynamicArrayFormulaInvariant="={1;2}";book.CalculateFullRebuild();
            bool rejected=false;try{InsertMirrorRowsPreservingDynamicArrays(sheet,1,10);}
            catch(InvalidOperationException ex){rejected=ex.Message.Contains("multi-cell dynamic spills");}
            Check(rejected&&sheet.Cells["B3"].HasDynamicArrayFormula&&sheet.Cells["B3"].Value.NumericValue==1&&sheet.Cells["B4"].Value.NumericValue==2,
                "Unsupported multi-cell spill was not rejected before mutation.");
            rejected=false;try{new MirrorDynamicArrayBatch(sheet);}
            catch(InvalidOperationException ex){rejected=ex.Message.Contains("single-cell dynamic arrays");}
            Check(rejected&&sheet.Cells["B3"].HasDynamicArrayFormula&&sheet.Cells["B4"].Value.NumericValue==2,
                "Unsupported batch spill was not rejected before mutation.");
        }
        IoTrial.WriteJson(Path.Combine(directory,"report.json"),new {cases,guardedNativeReloadAndGearValuesPassed=true,incompleteExportsRejected=true,multiCellSpillRejectedBeforeMutation=true,scope="Synthetic row-insert/copy diagnosis only."});
        Console.WriteLine("ARRAY INSERT PROBE "+directory);return 0;
    }

    static void AssertMirrorInsertFixture(DX.Workbook book)
    {
        var sheet=book.Worksheets[0];var expected=new HashSet<string>();
        foreach(int row in Enumerable.Range(1,15).Concat(Enumerable.Range(18,5)))for(int col=2;col<8;col++)
        {
            var cell=sheet.Cells[row,col];expected.Add(cell.GetReferenceA1());
            Check(cell.HasDynamicArrayFormula&&cell.Value.NumericValue==-10,"Guarded formula kind or value differs at "+cell.GetReferenceA1());
        }
        var actual=sheet.DynamicArrayFormulas.Cast<DX.DynamicArrayFormula>().Select(a=>a.Range.GetReferenceA1());
        Check(expected.SetEquals(actual)&&sheet.DynamicArrayFormulas.Count==120,"Guarded array geometry differs.");
    }

    // Narrow prototype for single-cell dynamic arrays only. Keep the expressions
    // present while native insertion translates references, then restore their
    // original kind at the mapped coordinates before any calculation or copy.
    static void InsertMirrorRowsPreservingDynamicArrays(DX.Worksheet sheet,int at,int count)
    {
        Check(at>=0&&count>0,"Positive row insertion required.");
        var arrays=sheet.DynamicArrayFormulas.Cast<DX.DynamicArrayFormula>()
            .Where(a=>a.Range.BottomRowIndex>=at).Select(a=>new {row=a.Range.TopRowIndex,col=a.Range.LeftColumnIndex,
                rows=a.Range.RowCount,cols=a.Range.ColumnCount}).ToArray();
        Check(arrays.All(a=>a.rows==1&&a.cols==1),"Row-insert adapter does not support multi-cell dynamic spills.");
        foreach(var a in arrays)
        {
            var cell=sheet.Cells[a.row,a.col];string formula=cell.FormulaInvariant;
            Check(cell.HasFormula,"Array manifest does not match its formula cell.");
            sheet.DynamicArrayFormulas.Remove(cell.GetDynamicArrayFormulaRange());
            cell.FormulaInvariant=formula;
        }
        sheet.Rows.Insert(at,count);
        foreach(var a in arrays)
        {
            var cell=sheet.Cells[a.row+count,a.col];Check(cell.HasFormula,"Moved dynamic-array expression was lost.");
            cell.DynamicArrayFormulaInvariant=cell.FormulaInvariant;
        }
    }
}
