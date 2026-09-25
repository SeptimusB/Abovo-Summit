using System;
using System.Collections.Generic;
using System.Data;
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
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraVerticalGrid;
using DevExpress.XtraVerticalGrid.Rows;

static class NativeFfrTests
{
    const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
    static int checks;
    static void Check(bool value,string text){if(!value)throw new Exception(text);Console.WriteLine("FFR_NATIVE PASS "+(++checks)+" "+text);Console.Out.Flush();}
    static object Field(object owner,string name){for(var type=owner.GetType();type!=null;type=type.BaseType){var f=type.GetField(name,F);if(f!=null)return f.GetValue(owner);}throw new Exception(name);}
    static object Call(object owner,string name,params object[] args){for(var type=owner.GetType();type!=null;type=type.BaseType){var m=type.GetMethod(name,F);if(m!=null)return m.Invoke(owner,args);}throw new Exception(name);}
    static void Pump(){Application.DoEvents();}
    static Control View(string name,params object[] args)=>(Control)Activator.CreateInstance(typeof(ModelChangeManagerV2).Assembly.GetType(name),args);
    static Form Host(Control view){var host=new Form{Opacity=0,ShowInTaskbar=false,ClientSize=new Size(1500,1000)};host.Controls.Add(view);view.Dock=DockStyle.Fill;host.Show();Pump();return host;}
    static void Input(Cell cell,object value,bool solid=true){cell.Value=CellValue.FromObject(value);cell.Protection.Locked=false;if(solid)cell.Fill.BackgroundColor=Color.LightBlue;}
    static string Create(string source)
    {
        string folder=Path.Combine(Path.GetDirectoryName(source),"native-ffr-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);string path=Path.Combine(folder,"private.xlsx");
        using(var wb=new Workbook()){
            var data=wb.Worksheets[0];data.Name="Data";Input(data.Cells["A1"],10d);data.Cells["B1"].Formula="=A1*2";data.Cells["C1"].Value="Locked";data.Cells["C1"].Fill.BackgroundColor=Color.LightBlue;
            var front=wb.Worksheets.Add("Front Sheet");front.Cells["A1"].Value="Synthetic FFR";front.Cells["A2"].Value="Front";Input(front.Cells["B5"],"RP1");Input(front.Cells["B6"],46113d);front.Cells["B6"].NumberFormat="dd/mm/yyyy";
            Input(front.Cells["B7"],"Yes");Input(front.Cells["B36"],"No");Input(front.Cells["B10"],"Yes");Input(front.Cells["C10"],"Entity");
            var workings=wb.Worksheets.Add("FFR Workings");workings.Cells["A5"].Value="Workings";workings.Cells["B10"].Value="Input";workings.Cells["B18"].Value="Calculated";Input(workings.Cells["E10"],10d);workings.Cells["E18"].Formula="=E10*2";
            var inputs=wb.Worksheets.Add("FFR Inputs Adj Stmt");inputs.Cells["B5"].Value="Actual";inputs.Cells["B6"].Value="Calculated";Input(inputs.Cells["C5"],20d);inputs.Cells["C6"].Formula="=C5*3";
            foreach(string name in new[]{"Statements","Assumptions & tenure inputs","Compliance Questions","FFR Key Defn"}){var s=wb.Worksheets.Add(name);s.Cells["A8"].Value="1";s.Cells["B8"].Value="Description";if(name=="FFR Key Defn")Input(s.Cells["C8"],"Definition",false);else s.Cells["C8"].Formula="='FFR Workings'!E18";}
            var validation=wb.Worksheets.Add("FFR Validation Summary");validation.Cells["A20"].Value="Current total";validation.Cells["B20"].Formula="='FFR Workings'!E10+'FFR Inputs Adj Stmt'!C5";
            wb.CalculateFullRebuild();wb.SaveDocument(path,DocumentFormat.Xlsx);
        }
        return path;
    }
    static void EditVertical(Control view,string gridName,string field,int record,object value)
    {
        var grid=(VGridControl)Field(view,gridName);grid.FocusedRow=grid.GetRowByFieldName(field);grid.FocusedRecord=record;grid.Focus();grid.ShowEditor();Pump();
        Check(grid.ActiveEditor!=null,"real vertical editor opens "+field);grid.ActiveEditor.EditValue=value;Check(grid.PostEditor(),"real vertical editor posts "+field);grid.CloseEditor();Pump();
    }
    internal static async Task Run(string source)
    {
        string path=Create(source);byte[] original=File.ReadAllBytes(path);var owners=new List<WorkbookCalculationSession>();
        foreach(var preference in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired}){
            var session=await WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(preference,false,false,120000,true));owners.Add(session);
            using(var model=new EngineChangeManagerTests.Model(path))try{
                var current=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Rebuild,new[]{new WorkbookReadArea("Data",0,0,1,3)});
                await model.Do(()=>{
                    typeof(ModelChangeManagerV2).GetMethod("BindEngineEditingTrial",F).Invoke(model.Manager,new object[]{session,current});var wb=model.Value.WB;
                    Console.WriteLine("FFR_NATIVE ENGINE "+session.EngineName);
                    using(var view=View("FFRWorkbookSheetView",0,"Data","A1:C1",null,false,false,false))using(var host=Host(view)){
                        var grid=(GridView)Field(view,"SheetView");grid.FocusedRowHandle=0;grid.FocusedColumn=grid.Columns["C0"];grid.Focus();grid.ShowEditor();Pump();
                        Check(grid.ActiveEditor!=null,"worksheet editor opens");grid.ActiveEditor.EditValue=12.5d;Check(grid.PostEditor()&&grid.UpdateCurrentRow(),"worksheet editor posts");grid.CloseEditor();Pump();
                        Check(wb.Worksheets["Data"].Cells["A1"].ModelValue().NumericValue==12.5&&wb.Worksheets["Data"].Cells["A1"].Value.NumericValue==10d,"FFR sheet changes native owner only");
                        var table=(DataTable)Field(view,"Snapshot");Check(Convert.ToDouble(table.Rows[0]["C1"])==25d,"FFR snapshot displays recalculated result");
                        grid.FocusedColumn=grid.Columns["C2"];grid.ShowEditor();Check(grid.ActiveEditor==null,"locked solid cell stays locked on unprotected sheet");
                        Check(model.Manager.Undo().BSuccess,"FFR sheet Undo admitted");Call(view,"RefreshFromWorkbook");Check(Convert.ToDouble(((DataTable)Field(view,"Snapshot")).Rows[0]["C1"])==20d,"FFR snapshot refreshes after Undo");
                    }
                    using(var view=View("FFRWorkingsVGridView",0))using(var host=Host(view)){
                        EditVertical(view,"Grid","Row_9",0,11d);Check(wb.Worksheets["FFR Workings"].Cells["E10"].ModelValue().NumericValue==11d,"Workings post reaches native owner");
                        var table=(DataTable)((VGridControl)Field(view,"Grid")).DataSource;Check(Convert.ToDouble(table.Rows[0]["Row_17"])==22d,"Workings calculated row refreshes");
                    }
                    using(var view=View("FFRInputsAdjStmtVGridView",0))using(var host=Host(view)){
                        EditVertical(view,"ActualGrid","Row_4",0,21d);Check(wb.Worksheets["FFR Inputs Adj Stmt"].Cells["C5"].ModelValue().NumericValue==21d,"Inputs post reaches native owner");
                    }
                    using(var view=View("FFRFrontSheetView",0))using(var host=Host(view)){
                        var editor=(BaseEdit)Field(view,"RPNumberEdit");editor.Focus();Pump();editor.EditValue="RP2";Call(view,"RPNumberValidated",editor,EventArgs.Empty);
                        Check(wb.Worksheets["Front Sheet"].Cells["B5"].ModelValue().TextValue=="RP2","front-sheet standalone editor uses captured native admission");
                        Check(((DateEdit)Field(view,"FirstForecastYearEdit")).DateTime==DateTime.FromOADate(46113d),"native numeric date renders as a date");
                    }
                    using(var view=View("FFRKeyDefinitionsVGridView",0))using(var host=Host(view)){
                        EditVertical(view,"Grid","Row_7",0,"Updated");Check(wb.Worksheets["FFR Key Defn"].Cells["C8"].ModelValue().TextValue=="Updated","key definitions retain unlocked-only rule");
                    }
                    foreach(string name in new[]{"FFRStatementsVGridView","FFRAssumptionsTenureVGridView","FFRComplianceQuestionsVGridView"}){
                        using(var view=name=="FFRStatementsVGridView"?View(name,0,"Statements","FFR Statements",151,false):View(name,0)){
                            var table=(DataTable)((VGridControl)Field(view,"Grid")).DataSource;Check(Convert.ToDouble(table.Rows[0]["Row_7"])==22d,name+" reads current native results");
                        }
                    }
                    using(var view=View("FFRValidationSummaryView",0)){
                        var table=(DataTable)((DevExpress.XtraGrid.GridControl)Field(view,"HardGrid")).DataSource;
                        Check((string)table.Rows[0]["Count"]=="32","validation summary reads final cross-sheet result");
                    }
                    using(var output=new Workbook()){
                        wb.Worksheets["FFR Workings"].Range["E10:E18"].CopyModelValuesTo(output.Worksheets[0].Range["A1:A9"]);
                        Check(output.Worksheets[0].Cells["A1"].Value.NumericValue==11d&&output.Worksheets[0].Cells["A9"].Value.NumericValue==22d,"return export uses current native inputs and formulas");
                        try{wb.Worksheets["Data"].Range["A1:B1"].CopyModelValuesTo(output.Worksheets[0].Range["B1"]);throw new Exception("Expected geometry rejection");}
                        catch(InvalidOperationException){Check(output.Worksheets[0].Cells["B1"].Value.IsEmpty,"mismatched export rejects before touching return cells");}
                    }
                    var support=typeof(ModelChangeManagerV2).Assembly.GetType("Abovo.ModelPostingChangeSupport");
                    var batch=new List<DataChangeEvent>{new DataChangeEvent{ModelID=0,WSName="Data",CellAddress="A1",ChangedValue=15d,DataFormat="N"},new DataChangeEvent{ModelID=0,WSName="FFR Inputs Adj Stmt",CellAddress="C5",ChangedValue=25d,DataFormat="N"}};
                    Check((bool)support.GetMethod("PostModelEngineBatch",F).Invoke(null,new object[]{0,batch,WorkbookValuePermission.UnlockedSolidFill,"FFR batch test"}),"FFR paste dispatcher accepts one native batch");
                    Check(wb.Worksheets["Data"].Cells["A1"].ModelValue().NumericValue==15d&&wb.Worksheets["FFR Inputs Adj Stmt"].Cells["C5"].ModelValue().NumericValue==25d,"both native paste values present");
                    Check(model.Manager.Undo().BSuccess&&wb.Worksheets["Data"].Cells["A1"].ModelValue().NumericValue==10d&&wb.Worksheets["FFR Inputs Adj Stmt"].Cells["C5"].ModelValue().NumericValue==21d,"one Undo restores whole FFR paste batch");
                    var unlockedPattern=new WorkbookCellState(1d,"","General",false,false,false,false,false);
                    var lockedSolid=new WorkbookCellState(1d,"","General",true,true,false,false,false);
                    Check(!unlockedPattern.AllowsValueEdit(WorkbookValuePermission.UnlockedSolidFill)&&!lockedSolid.AllowsValueEdit(WorkbookValuePermission.UnlockedSolidFill),"FFR conjunction rejects either patterned or locked input");
                    Check(wb.Worksheets["FFR Workings"].Cells["E10"].Value.NumericValue==10d&&wb.Worksheets["Front Sheet"].Cells["B5"].Value.TextValue=="RP1","all FFR edits leave display workbook inputs untouched");
                    return true;
                });
            }finally{await session.CloseAsync();}
        }
        Check(File.ReadAllBytes(path).SequenceEqual(original),"original generated file unchanged");await NativeProcessChecks.RequireOwnedProcesses(owners);Console.WriteLine("FFR_NATIVE ASSERTIONS="+checks);
    }
}
