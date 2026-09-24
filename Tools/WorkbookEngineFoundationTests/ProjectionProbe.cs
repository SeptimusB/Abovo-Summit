using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;
using DevExpress.XtraSpreadsheet.Services;

// Deliberately isolated feasibility probe, not a live calculation adapter.
// Values deliberately differ from the formulas' local answers so an accidental
// DevExpress recalculation cannot masquerade as accepted external results.
static class ProjectionProbe
{
    sealed class ExternalValues : ICustomCalculationService
    {
        internal readonly Dictionary<CellKey, CellValue> Values = new Dictionary<CellKey, CellValue>();
        internal int Hits, Missing;
        public bool OnBeginCalculation() => true;
        public void OnBeginCellCalculation(CellCalculationArgs args)
        {
            CellValue value;
            if (Values.TryGetValue(args.CellKey, out value)) { args.Value = value; Hits++; }
            else { args.Value = CellValue.ErrorValueNotAvailable; Missing++; }
            args.Handled = true;
        }
        public bool OnBeginCircularReferencesCalculation() => false;
        public void OnEndCalculation() { }
        public void OnEndCellCalculation(CellKey key, CellValue before, CellValue after) { }
        public void OnEndCircularReferencesCalculation(IList<CellKey> cells) { }
        public bool ShouldMarkupCalculateAlwaysCells() => true;
    }

    static int assertions;
    static void Check(bool condition, string text)
    {
        if (!condition) throw new InvalidOperationException(text);
        Console.WriteLine("PROJECTION PASS " + (++assertions) + " " + text);
    }

    internal static void Run()
    {
        using (var workbook = new Workbook())
        {
            workbook.Options.CalculationMode = WorkbookCalculationMode.Manual;
            Console.WriteLine("PROJECTION ENGINE_DEFAULT=" + workbook.Options.CalculationEngineType);
            workbook.Options.CalculationEngineType = CalculationEngineType.ChainBased;
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Value = 10;
            sheet.Cells["B1"].FormulaInvariant = "=A1*2";
            sheet.Cells["B2"].FormulaInvariant = "=B1+1";
            sheet.Cells["C1"].FormulaInvariant = "=1/0";
            sheet.Cells["D1"].Value = "Literal input";
            sheet.Cells["E1"].DynamicArrayFormulaInvariant = "={1;2}";
            sheet.Cells["B1"].Fill.PatternType = PatternType.Solid;
            sheet.Cells["B1"].Fill.BackgroundColor = Color.Lavender;
            sheet.Cells["B1"].Font.Color = Color.DarkBlue;
            sheet.Cells["B1"].NumberFormat = "#,##0.00";
            sheet.Cells["B1"].Protection.Locked = true;
            sheet.Cells["B2"].Fill.PatternType = PatternType.Gray125;
            workbook.CalculateFull();
            Console.WriteLine("PROJECTION BASELINE=" + sheet.Cells["B1"].Value);
            Check(sheet.Cells["B1"].Value.NumericValue == 20, "known local formula answer before projection");
            Check(sheet.Cells["E2"].Value.NumericValue == 2 && sheet.Cells["E1"].GetDynamicArrayFormulaRange().GetReferenceA1() == "E1:E2",
                "dynamic formula genuinely spills before the callback is installed");
            var beforeFormula = sheet.Cells["B1"].FormulaInvariant;
            var service = new ExternalValues();
            // In this 25.2.4 probe callback SheetId is zero, while Sheet.Id is
            // one. No inferred sheet-ID mapping is used in production.
            service.Values[new CellKey(sheet.Index, 1, 0)] = 800d;
            service.Values[new CellKey(sheet.Index, 1, 1)] = 801d;
            workbook.AddService(typeof(ICustomCalculationService), service);
            workbook.CalculateFull();
            Console.WriteLine("PROJECTION VALUES B1=" + sheet.Cells["B1"].Value + " B2=" + sheet.Cells["B2"].Value + " hits=" + service.Hits + " missing=" + service.Missing);
            Check(sheet.Cells["B1"].Value.NumericValue == 800 && sheet.Cells["B2"].Value.NumericValue == 801,
                "external results replace cached values without local formula evaluation");
            Check(sheet.Cells["C1"].Value == CellValue.ErrorValueNotAvailable && service.Missing > 0,
                "missing external value is unavailable, not a local financial answer");
            Check(sheet.Cells["B1"].FormulaInvariant == beforeFormula && sheet.Cells["D1"].Value.TextValue == "Literal input",
                "formulas and literal inputs preserved");
            Check(sheet.Cells["B1"].Fill.PatternType == PatternType.Solid && sheet.Cells["B2"].Fill.PatternType == PatternType.Gray125 &&
                sheet.Cells["B1"].Fill.BackgroundColor.ToArgb() == Color.Lavender.ToArgb() && sheet.Cells["B1"].Font.Color.ToArgb() == Color.DarkBlue.ToArgb() &&
                sheet.Cells["B1"].NumberFormat == "#,##0.00" && sheet.Cells["B1"].Protection.Locked,
                "workbook-owned pattern, colours, number format and lock unchanged");

            var source = sheet.Range["B1:B2"].GetDataSource(new RangeDataSourceOptions {
                UseFirstRowAsHeader = false, PreserveFormulas = false, EditingOptions = DataSourceEditingOptions.ReadOnly });
            var list = (IList)source;
            var properties = ((ITypedList)source).GetItemProperties(null);
            Check(Convert.ToDouble(properties[0].GetValue(list[0])) == 800, "existing RangeDataSource reads externally supplied value");
            service.Values[new CellKey(sheet.Index, 1, 0)] = 900d;
            workbook.CalculateFull();
            Check(Convert.ToDouble(properties[0].GetValue(list[0])) == 900, "existing datasource sees the next cached value");

            // Check actual native dynamic-array behavior rather than assuming
            // the scalar callback can represent a spill rectangle.
            service.Values[new CellKey(sheet.Index, 4, 0)] = 700d;
            service.Values[new CellKey(sheet.Index, 4, 1)] = 701d;
            workbook.CalculateFull();
            Console.WriteLine("PROJECTION DYNAMIC anchor=" + sheet.Cells["E1"].Value + " spill=" + sheet.Cells["E2"].Value +
                " hasArray=" + sheet.Cells["E1"].HasDynamicArrayFormula + " range=" + sheet.Cells["E1"].GetDynamicArrayFormulaRange()?.GetReferenceA1());
            Check(sheet.Cells["E1"].Value.NumericValue == 700 && sheet.Cells["E2"].Value.NumericValue == 2,
                "reproduced stale existing spill: callback updates the anchor but not the spilled value");
            sheet.Cells["G1"].DynamicArrayFormulaInvariant = "={3;4}";
            service.Values[new CellKey(sheet.Index, 6, 0)] = 1000d;
            service.Values[new CellKey(sheet.Index, 6, 1)] = 1001d;
            workbook.CalculateFull();
            Check(sheet.Cells["G1"].Value.NumericValue == 1000 && sheet.Cells["G2"].Value.IsEmpty,
                "new spill does not populate through scalar callbacks; general projection is ineligible");
            Console.WriteLine("PROJECTION CALLBACKS hits=" + service.Hits + " missing=" + service.Missing);
        }
        Console.WriteLine("PROJECTION ASSERTIONS=" + assertions + "; scalar feasibility only, not live integration");
    }
}
