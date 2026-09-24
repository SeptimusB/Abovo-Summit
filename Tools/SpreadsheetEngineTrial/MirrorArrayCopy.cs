using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DX = DevExpress.Spreadsheet;

internal static partial class EngineBenchmark
{
    // Isolated Funding mirror port only. This post-copy check is NOT sufficient
    // after an unguarded row insertion: the 25.2.4 array registry can still point
    // at the old rows, making HasDynamicArrayFormula misleading. Guard insertion
    // separately, then preserve the source kind with native translated formulas.
    // Multi-cell spill templates are deliberately outside this proven operation.
    static int CopyMirrorRowPreservingDynamicArrays(DX.Worksheet sheet,DX.CellRange source,DX.CellRange target)
    {
        Check(source.Worksheet==sheet&&target.Worksheet==sheet&&source.RowCount==1&&source.ColumnCount==target.ColumnCount,
            "Array-aware mirror copy requires one same-width source row on the same sheet.");
        Check(target.TopRowIndex>source.BottomRowIndex,"Mirror copy cannot overlap its source.");
        var arrays=sheet.DynamicArrayFormulas.Cast<DX.DynamicArrayFormula>().Where(a=>
            a.Range.TopRowIndex<=source.TopRowIndex&&a.Range.BottomRowIndex>=source.TopRowIndex&&
            a.Range.LeftColumnIndex<=source.RightColumnIndex&&a.Range.RightColumnIndex>=source.LeftColumnIndex).ToArray();
        foreach(var array in arrays)Check(array.Range.RowCount==1&&array.Range.ColumnCount==1,
            "A multi-cell dynamic-array template requires a separate structural adapter.");
        var columns=arrays.Select(a=>a.Range.LeftColumnIndex-source.LeftColumnIndex).ToArray();
        target.CopyFrom(source,DX.PasteSpecial.All);
        int restored=0;
        foreach(int column in columns)for(int row=target.TopRowIndex;row<=target.BottomRowIndex;row++)
        {
            var cell=sheet.Cells[row,target.LeftColumnIndex+column];
            Check(cell.HasFormula,"Copied array lost its expression.");
            if(!cell.HasDynamicArrayFormula)
            {
                string translated=cell.FormulaInvariant;
                cell.DynamicArrayFormulaInvariant=translated;
                restored++;
            }
            var range=cell.GetDynamicArrayFormulaRange();
            Check(range.RowCount==1&&range.ColumnCount==1,"Copied dynamic array expanded unexpectedly.");
        }
        return restored;
    }

    internal static int RunMirrorArrayCopyProbe(string directory)
    {
        directory=IoTrial.InTrial(directory);Check(!Directory.Exists(directory),"New fixture directory required.");
        Directory.CreateDirectory(directory);string file=Path.Combine(directory,"array-copy.xlsm");
        int rawArrays,restored;
        using(var book=new DX.Workbook())
        {
            book.Options.CalculationMode=DX.WorkbookCalculationMode.Manual;
            book.Options.CalculationEngineType=DX.CalculationEngineType.Recursive;
            var sheet=book.Worksheets[0];sheet.Name="ArrayCopy";
            sheet.Cells["E1"].Value=10;sheet.Cells["E2"].Value=20;
            for(int row=1;row<24;row++)sheet.Cells[row,0].Value=1;
            sheet.Cells["B2"].DynamicArrayFormulaInvariant="=-INDEX($E$1:$E$2,$A2)";
            sheet.Range["B3:B12"].CopyFrom(sheet.Range["B2"],DX.PasteSpecial.All);
            rawArrays=Enumerable.Range(2,10).Count(r=>sheet.Cells[r,1].HasDynamicArrayFormula);
            restored=CopyMirrorRowPreservingDynamicArrays(sheet,sheet.Range["B2"],sheet.Range["B14:B23"]);
            // Unknown multi-cell sources must fail before copying anything.
            sheet.Cells["G2"].DynamicArrayFormulaInvariant="={1;2}";book.CalculateFullRebuild();
            bool rejected=false;try{CopyMirrorRowPreservingDynamicArrays(sheet,sheet.Range["G2"],sheet.Range["G6:G7"]);}
            catch(InvalidOperationException e){rejected=e.Message.Contains("multi-cell dynamic-array template");}
            Check(rejected&&!sheet.Cells["G6"].HasFormula&&sheet.Cells["G6"].Value.IsEmpty,"Unsupported spill copy was not rejected before writing.");
            using(var stream=new FileStream(file,FileMode.CreateNew,FileAccess.Write))book.SaveDocument(stream,DX.DocumentFormat.Xlsm);
        }
        using(var book=new DX.Workbook())
        {
            book.Options.CalculationMode=DX.WorkbookCalculationMode.Manual;Check(book.LoadDocument(file),"Fixture reload failed.");
            book.CalculateFullRebuild();var sheet=book.Worksheets[0];
            for(int row=13;row<23;row++)Check(sheet.Cells[row,1].HasDynamicArrayFormula&&sheet.Cells[row,1].Value.NumericValue==-10,"Native array copy/reload differs.");
        }
        using(var gear=NewMirrorProjection(false))
        {
            gear.Open(file);gear.Calculate(2);var values=gear.Get("ArrayCopy",13,1,10,1);
            for(int row=0;row<10;row++)Check(SameValue(values[row,0],-10.0),"Gear array-copy result differs.");
        }
        IoTrial.WriteJson(Path.Combine(directory,"report.json"),new {passed=true,rawArraysOfTen=rawArrays,restored, nativeReloadAndGearValuesPassed=true,multicellSpillRejectedBeforeWrite=true,
            scope="Synthetic single-cell array row-copy fixture; no production or original workbook changes."});
        Console.WriteLine("ARRAY COPY PASS raw="+rawArrays+" restored="+restored);return 0;
    }
}
