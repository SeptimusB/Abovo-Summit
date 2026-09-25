using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Abovo;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
using DevExpress.XtraCharts;

static class NativeDashboardTests
{
    const BindingFlags F=BindingFlags.Instance|BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic;
    static int checks;
    static void Check(bool value,string text){if(!value)throw new Exception(text);Console.WriteLine("DASHBOARD_NATIVE PASS "+(++checks)+" "+text);Console.Out.Flush();}
    static IEnumerable<Control> Children(Control control){foreach(Control child in control.Controls){yield return child;foreach(var nested in Children(child))yield return nested;}}
    static Control View(string name)=>(Control)Activator.CreateInstance(typeof(ModelChangeManagerV2).Assembly.GetType(name),new object[]{0});
    static object Field(object owner,string name)=>owner.GetType().GetField(name,F).GetValue(owner);
    static Form Host(Control view){var form=new Form{Opacity=0,ShowInTaskbar=false,ClientSize=new Size(1500,1000)};form.Controls.Add(view);view.Dock=DockStyle.Fill;form.Show();Application.DoEvents();return form;}
    static string Create(string source)
    {
        string folder=Path.Combine(Path.GetDirectoryName(source),"native-dashboard-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);string path=Path.Combine(folder,"private.xlsx");
        using(var wb=new Workbook()){
            var dashboard=wb.Worksheets[0];dashboard.Name="Funding Dashboard";dashboard.Cells["A2"].Formula="=\"Selected \"&V6";
            foreach(string address in new[]{"D6","I6","R6","V6"}){var cell=dashboard.Cells[address];cell.Protection.Locked=false;cell.Fill.BackgroundColor=Color.LightBlue;}
            dashboard.Cells["D6"].Value="Funder A";dashboard.Cells["I6"].Value="Facility A";dashboard.Cells["R6"].Value="Covenant A";dashboard.Cells["V6"].Value=1d;
            dashboard.DataValidations.Add(dashboard.Range["D6"],DataValidationType.List,"Funder A,Funder B");
            dashboard.DataValidations.Add(dashboard.Range["I6"],DataValidationType.List,"Facility A,Facility B");
            dashboard.DataValidations.Add(dashboard.Range["R6"],DataValidationType.List,"Covenant A,Covenant B");
            dashboard.DataValidations.Add(dashboard.Range["V6"],DataValidationType.List,"1,2,3");
            var charts=wb.Worksheets.Add("OW - Charts Source Data");charts.Cells["AE10"].Formula="='Funding Dashboard'!V6*10";
            charts.Cells["A53"].Value="Income";charts.Cells["D53"].Formula="='Funding Dashboard'!V6*20";
            wb.Worksheets.Add("OW - Live Covenant Calculation");wb.Worksheets.Add("Covenants");var summary=wb.Worksheets.Add("BP Dashboard");summary.Cells["H9"].Formula="='Funding Dashboard'!A2";
            var stock=wb.Worksheets.Add("FFRW - Stock Numbers");stock.Cells["J11"].Formula="='Funding Dashboard'!V6*30";
            wb.CalculateFullRebuild();wb.SaveDocument(path,DocumentFormat.Xlsx);
        }
        return path;
    }
    internal static async Task Run(string source)
    {
        string path=Create(source);byte[] original=File.ReadAllBytes(path);var owners=new List<WorkbookCalculationSession>();
        foreach(var preference in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired}){
            var session=await WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(preference,false,false,120000,true));owners.Add(session);
            using(var model=new EngineChangeManagerTests.Model(path))try{
                var result=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Rebuild,new[]{new WorkbookReadArea("Funding Dashboard",5,21,1,1)});
                await model.Do(()=>{
                    typeof(ModelChangeManagerV2).GetMethod("BindEngineEditingTrial",F).Invoke(model.Manager,new object[]{session,result});
                    using(var view=View("FundingDashboard"))using(var form=Host(view)){
                        Console.WriteLine("DASHBOARD_NATIVE ENGINE "+session.EngineName);
                        var selector=Children(view).OfType<ComboBoxEdit>().Single(e=>(string)Field(e.Tag,"CellAddress")=="V6");
                        selector.Focus();Application.DoEvents();selector.EditValue="2";Application.DoEvents();
                        Check(model.Value.WB.Worksheets["Funding Dashboard"].Cells["V6"].ModelValue().NumericValue==2d,"selector posts to selected owner");
                        Check(model.Value.WB.Worksheets["Funding Dashboard"].Cells["V6"].Value.NumericValue==1d,"selector never writes the display workbook");
                        Check(Children(view).OfType<LabelControl>().Any(l=>l.Text=="Selected 2"),"heading refreshes from native calculation");
                        var series=Children(view).OfType<ChartControl>().SelectMany(c=>c.Series.Cast<Series>()).Single(s=>s.Name=="Loans Drawn");
                        Check(series.Points.Count==1&&series.Points[0].Values[0]==20d,"chart reads the recalculated native value");
                        Check(model.Manager.Undo().BSuccess,"selector Undo succeeds");
                        view.GetType().GetMethod("RefreshData").Invoke(view,null);Application.DoEvents();
                        Check(Children(view).OfType<LabelControl>().Any(l=>l.Text=="Selected 1"),"dashboard refreshes after Undo");
                        Check(model.Manager.Redo().BSuccess,"selector Redo succeeds");
                    }
                    using(var view=View("BPDashboard"))using(var form=Host(view)){
                        var stock=(ChartControl)Field(view,"StockChart");var first=((IEnumerable)stock.DataSource).Cast<object>().First();
                        Check((int)first.GetType().GetProperty("Numbers").GetValue(first,null)==60,"BP chart reads current native stock results");
                        var browser=(WebBrowser)Field(view,"GearingBrowserMsg");
                        var wait=System.Diagnostics.Stopwatch.StartNew();while(browser.ReadyState!=WebBrowserReadyState.Complete&&wait.ElapsedMilliseconds<5000){Application.DoEvents();System.Threading.Thread.Sleep(10);}
                        Check(browser.DocumentText.Contains("Selected 2"),"BP summary HTML reads current native text");
                    }
                    return true;
                });
            }finally{await session.CloseAsync();}
        }
        Check(File.ReadAllBytes(path).SequenceEqual(original),"generated original unchanged");await NativeProcessChecks.RequireOwnedProcesses(owners);Console.WriteLine("DASHBOARD_NATIVE ASSERTIONS="+checks);
    }
}
