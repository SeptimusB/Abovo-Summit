using System;
using System.Collections.Generic;
using System.Linq;
using DX = DevExpress.Spreadsheet;

internal static partial class EngineBenchmark
{
    // Disposable worker only: remove registrations once, retain every formula for
    // native reference translation, track only the reviewed whole-row geometry,
    // then restore dynamic semantics BEFORE calculation, validation or saving.
    // An exception invalidates the worker; a partially staged book must not save.
    sealed class MirrorDynamicArrayBatch
    {
        readonly DX.Worksheet sheet;
        HashSet<Tuple<int,int>> anchors;
        bool complete;

        internal MirrorDynamicArrayBatch(DX.Worksheet sheet)
        {
            this.sheet=sheet;
            var saved=sheet.DynamicArrayFormulas.Cast<DX.DynamicArrayFormula>().Select(a=>new {
                row=a.Range.TopRowIndex,col=a.Range.LeftColumnIndex,rows=a.Range.RowCount,cols=a.Range.ColumnCount,
                formula=sheet.Cells[a.Range.TopRowIndex,a.Range.LeftColumnIndex].FormulaInvariant}).ToArray();
            Check(saved.All(a=>a.rows==1&&a.cols==1&&!String.IsNullOrEmpty(a.formula)),
                "Batch insert adapter requires consistent single-cell dynamic arrays.");
            anchors=new HashSet<Tuple<int,int>>(saved.Select(a=>Tuple.Create(a.row,a.col)));
            Check(anchors.Count==saved.Length,"Duplicate native array anchors.");
            sheet.DynamicArrayFormulas.Clear();
            foreach(var a in saved)sheet.Cells[a.row,a.col].FormulaInvariant=a.formula;
            Check(sheet.DynamicArrayFormulas.Count==0,"Array registrations did not clear.");
        }

        internal void InsertRows(int at,int count)
        {
            Check(!complete&&at>=0&&count>0,"Active batch and positive insertion required.");
            sheet.Rows.Insert(at,count);
            anchors=new HashSet<Tuple<int,int>>(anchors.Select(a=>Tuple.Create(a.Item1>=at?a.Item1+count:a.Item1,a.Item2)));
        }

        internal void CopyRow(DX.CellRange source,DX.CellRange target)
        {
            Check(!complete&&source.Worksheet==sheet&&target.Worksheet==sheet&&source.RowCount==1&&source.ColumnCount==target.ColumnCount&&target.TopRowIndex>source.BottomRowIndex,
                "Batch copy requires a same-width single-row non-overlapping template.");
            Check(!anchors.Any(a=>a.Item1>=target.TopRowIndex&&a.Item1<=target.BottomRowIndex&&a.Item2>=target.LeftColumnIndex&&a.Item2<=target.RightColumnIndex),
                "Batch target contains existing dynamic arrays.");
            var columns=anchors.Where(a=>a.Item1==source.TopRowIndex&&a.Item2>=source.LeftColumnIndex&&a.Item2<=source.RightColumnIndex)
                .Select(a=>a.Item2-source.LeftColumnIndex).ToArray();
            target.CopyFrom(source,DX.PasteSpecial.All);
            foreach(int column in columns)for(int row=target.TopRowIndex;row<=target.BottomRowIndex;row++)anchors.Add(Tuple.Create(row,target.LeftColumnIndex+column));
        }

        internal int Complete()
        {
            Check(!complete&&sheet.DynamicArrayFormulas.Count==0,"Unexpected calculation/array mutation inside structural batch.");
            foreach(var a in anchors)
            {
                var cell=sheet.Cells[a.Item1,a.Item2];Check(cell.HasFormula,"Moved dynamic-array expression was lost.");
                cell.DynamicArrayFormulaInvariant=cell.FormulaInvariant;
            }
            var actual=sheet.DynamicArrayFormulas.Cast<DX.DynamicArrayFormula>().Select(a=>Tuple.Create(a.Range.TopRowIndex,a.Range.LeftColumnIndex));
            Check(anchors.Count==sheet.DynamicArrayFormulas.Count&&anchors.SetEquals(actual),"Restored array geometry differs from the structural map.");
            complete=true;return anchors.Count;
        }
    }
}
