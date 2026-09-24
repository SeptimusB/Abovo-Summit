using System;
using System.IO;
using System.Globalization;
using Gear = SpreadsheetGear;

internal static partial class EngineBenchmark
{
    // Synthetic, no customer data, no workbook writes and no macro execution.
    // Separates Excel's grouped operation from serial insertion semantics.
    internal static int FundingThreeDProbe(string output)
    {
        output=IoTrial.InTrial(output);Check(!File.Exists(output),"New report required.");
        SpreadsheetGearTrial.Activate();var set=Gear.Factory.GetWorkbookSet(CultureInfo.InvariantCulture);
        set.Calculation=Gear.Calculation.Manual;set.BackgroundCalculation=false;set.CalculationOnDemand=false;set.EnableWebService=false;
        object gear;
        var w=set.Workbooks.Add();
        try
        {
            var a=w.Worksheets[0];a.Name="First";var b=w.Worksheets.Add();b.Name="Last";var r=w.Worksheets.Add();r.Name="Result";
            a.Cells["B1"].Value=5;b.Cells["B1"].Value=7;r.Cells["A1"].Formula="=SUM(First:Last!B1)";set.CalculateFullRebuild();
            string before=r.Cells["A1"].Formula;a.Cells["B:C"].Insert();b.Cells["B:C"].Insert();set.CalculateFullRebuild();
            gear=new {before,after=r.Cells["A1"].Formula,value=r.Cells["A1"].Value};
        }finally{w.Close();}
        object serial,grouped;using(var e=new ExcelEngine(false)){serial=e.ThreeDProbe(false);grouped=e.ThreeDProbe(true);}
        IoTrial.WriteJson(output,new {gearSerial=gear,excelSerial=serial,excelGrouped=grouped,scope="Synthetic insert only; no copying, no VBA, no saved workbooks"});return 0;
    }
    sealed partial class ExcelEngine
    {
        internal object ThreeDProbe(bool grouped)
        {
            dynamic sheets=null,a=null,b=null,r=null,c=null,range=null,selection=null;
            try
            {
                book=books.Add();sheets=book.Worksheets;a=sheets.Item[1];a.Name="First";b=sheets.Add(After:a);b.Name="Last";r=sheets.Add(After:b);r.Name="Result";
                c=a.Range["B1"];c.Value2=5;Release((object)c);c=null;c=b.Range["B1"];c.Value2=7;Release((object)c);c=null;
                c=r.Range["A1"];c.Formula="=SUM(First:Last!B1)";string before=(string)c.Formula;
                if(grouped){a.Select(true);b.Select(false);a.Activate();range=a.Range["B:C"];range.Select();selection=app.Selection;selection.Insert(-4161);a.Select();}
                else{range=a.Range["B:C"];range.Insert(-4161);Release((object)range);range=b.Range["B:C"];range.Insert(-4161);}
                app.CalculateFullRebuild();return new {before,after=(string)c.Formula,value=(object)c.Value2};
            }
            finally{Release((object)selection);Release((object)range);Release((object)c);Release((object)r);Release((object)b);Release((object)a);Release((object)sheets);if(book!=null){book.Close(false);Release((object)book);book=null;}}
        }
    }
}
