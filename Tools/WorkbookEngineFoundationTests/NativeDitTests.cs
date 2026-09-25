using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Abovo;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraVerticalGrid;
using DevExpress.XtraVerticalGrid.Rows;

static class NativeDitTests
{
    const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
    static int checks;
    static object Field(object o,string name){var f=o.GetType().GetField(name,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(name,F).GetValue(o,null);}
    static object Call(object o,string name,params object[] args){return o.GetType().GetMethod(name,F).Invoke(o,args);}
    static void Check(bool ok,string message){if(!ok)throw new Exception(message);Console.WriteLine("DIT_NATIVE PASS "+(++checks)+" "+message);Console.Out.Flush();}
    static void Pump(){Application.DoEvents();Application.RaiseIdle(EventArgs.Empty);Application.DoEvents();}
    static IEnumerable<Control> Children(Control parent){foreach(Control child in parent.Controls){yield return child;foreach(var nested in Children(child))yield return nested;}}
    internal static Task Run(string source,bool excel,bool saveLifecycle=false,bool openingSelection=false)
    {
        var completion=new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread=new Thread(()=>{try{RunOnOwner(source,excel,saveLifecycle,openingSelection);completion.SetResult(true);}catch(Exception e){completion.SetException(e);}}){IsBackground=true};
        thread.SetApartmentState(ApartmentState.STA);thread.Start();return completion.Task;
    }
    static void RunOnOwner(string source,bool excel,bool saveLifecycle,bool openingSelection)
    {
        Thread.CurrentThread.CurrentCulture=CultureInfo.GetCultureInfo("en-GB");
        var before=NativeProcessChecks.ExcelIds();
        var owners=new List<WorkbookCalculationSession>();
        var app=typeof(ModelChangeManagerV2).Assembly;
        app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
        var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
        string folder=Path.Combine(Environment.CurrentDirectory,"obj","EngineDitTests",Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
        File.Copy(Path.Combine(Environment.CurrentDirectory,"Structure.xml"),Path.Combine(folder,"Structure.xml"));
        string copy=Path.Combine(folder,"private-native-dit"+Path.GetExtension(source));File.Copy(source,copy);
        var open=files.GetMethod("OpenModel");var opened=(AbovoAppCls.AbovoTransaction)open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
        Check(opened.BSuccess&&!opened.BError,"normal application opens a private model: "+opened.StrResponseMessage);
        var model=FileManager.ExcelModels[opened.IntegerReturn];var book=model.WB;WorkbookCalculationSession session=null;
        try
        {
            var options=new WorkbookEngineOptions(excel?WorkbookEnginePreference.ExcelRequired:WorkbookEnginePreference.DevExpressOnly,true,true,120000,true,saveLifecycle,saveLifecycle);
            if(openingSelection){
                app.GetType("Abovo.WorkbookEngines.ModelEngineSelection").GetMethod("Activate",F).Invoke(null,new object[]{model,options,null});
                session=(WorkbookCalculationSession)Field(model.ChangeManager,"engineTrial");
                Check(!model.IsDirty&&model.ChangeManager.GetHistoryTable().Rows.Count==0,"opening selection preserves clean/history state");
            }else{
                session=WorkbookCalculationSession.OpenAsync(copy,options).GetAwaiter().GetResult();
                var initial=session.CalculateAndReadAsync(0,WorkbookCalculationKind.Rebuild,new[]{new WorkbookReadArea("Check Sheet",0,0,1,1)}).GetAwaiter().GetResult();
                Call(model.ChangeManager,"BindEngineEditingTrial",session,initial);
            }
            owners.Add(session);
            Check(model.ChangeManager.HasEngineEditingTrial,"normal application model bound to "+session.EngineName);
            var unexpectedDialogs=new List<string>();
            using(var dialogReporter=new System.Windows.Forms.Timer())
            using(var form=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{model.ModelID,0,"Normal"}))
            {
                dialogReporter.Interval=200;
                dialogReporter.Tick+=(sender,args)=>{
                    foreach(Form dialog in Application.OpenForms.Cast<Form>().ToArray())if(dialog.Modal){
                        string message=dialog.Text+": "+String.Join(" | ",Children(dialog).Select(c=>c.Text).Where(t=>!String.IsNullOrWhiteSpace(t)).Distinct());
                        Console.WriteLine("DIT_NATIVE UNEXPECTED_DIALOG "+message);Console.Out.Flush();unexpectedDialogs.Add(message);
                        dialog.DialogResult=DialogResult.Cancel;dialog.Close();
                    }
                };
                dialogReporter.Start();
                form.Opacity=0;form.ShowInTaskbar=false;form.ClientSize=new Size(1750,1000);form.Show();Pump();
                var beforeSummary=model.ChangeManager.EngineEditingResult;
                Call(form,"CalculateSidebarWorkbook",book,"Native regression");
                Check(model.ChangeManager.EngineEditingResult.CalculationGeneration>beforeSummary.CalculationGeneration,
                    "explicit summary refresh calculates the selected native owner");
                foreach(int csid in new[]{35,33})
                {
                    var timer=Stopwatch.StartNew();Console.WriteLine("DIT_NATIVE OPEN section="+csid);Console.Out.Flush();
                    Call(form,"ShowInterface",0,csid,false,"None",null,-1);Pump();var dit=(Control)Field(form,"ActiveInterface");
                    if(csid==33){Call(dit,"BuildSection",1,false,false);var tabs=Field(dit,"XtraTabControlNewGIT");tabs.GetType().GetProperty("SelectedTabPageIndex").SetValue(tabs,1,null);Pump();}
                    Console.WriteLine("DIT_NATIVE BIND section="+csid+" ms="+timer.ElapsedMilliseconds);Console.Out.Flush();
                    var positions=((IEnumerable)Call(dit,"NavigationPositions")).Cast<object>().ToArray();
                    object selected=null;Cell target=null;GridView gridView=null;VGridControl vertical=null;
                    foreach(var point in positions)
                    {
                        if(Field(point,"Header")!=null||Field(point,"VHeader")!=null||Field(point,"Standalone")!=null)continue;
                        var view=Field(point,"View") as GridView;var vg=Field(point,"Host") as VGridControl;var column=Field(point,"Column") as GridColumn;var row=Field(point,"VRow") as EditorRow;
                        if(view==null&&row==null)continue;
                        var ds=Field(view!=null?view.GridControl.DataSource:vg.DataSource,"UBSTag");int dsid=(int)Field(ds,"DSIndex");
                        var datasets=(IList)Field(Field(dit,"DataPres"),"DataSets");var data=datasets[dsid];
                        int col=(int)(view!=null?Call(dit,"GetGridColumnIndex",column):Call(dit,"GetVGridColumnIndex",row,0));
                        int dataRow=view!=null?view.GetDataSourceRowIndex((int)Field(point,"RowHandle")):vg.GetDataSourceRecordIndex((int)Field(point,"Record"));
                        var rows=(IList)Field(data,"DataRows");var cells=(IList)Field(rows[dataRow],"DataCells");var cellMap=cells[col];
                        var cell=book.Worksheets[(string)Field(cellMap,"SourceSheet")].Cells[(string)Field(cellMap,"SourceAddress")];
                        var opening=csid==33?book.Worksheets["Funding Assumptions"].Range["Rep_Fund_08"]:null;
                        bool openingInput=opening!=null&&cell.Worksheet==opening.Worksheet&&cell.RowIndex>=opening.TopRowIndex&&cell.RowIndex<=opening.BottomRowIndex&&cell.ColumnIndex>=opening.LeftColumnIndex&&cell.ColumnIndex<=opening.RightColumnIndex;
                        if((csid==35&&cell.GetReferenceA1()=="D8")||openingInput){selected=point;target=cell;gridView=view;vertical=vg;break;}
                    }
                    Check(selected!=null,"native navigation resolves expected input in section "+csid);
                    var beforeValue=target.ModelValue();var template=target.Value;
                    Call(dit,"ActivateEditorPosition",selected);Pump();var editor=gridView!=null?gridView.ActiveEditor:vertical.ActiveEditor;
                    Check(editor!=null,"real native editor opens "+target.GetReferenceA1());
                    double changed=(beforeValue.IsNumeric?beforeValue.NumericValue:0)+0.125;
                    editor.EditValue=changed;
                    Check(gridView!=null?gridView.PostEditor()&&gridView.UpdateCurrentRow():vertical.PostEditor(),"real native editor posts");
                    Check(unexpectedDialogs.Count==0,"no edit failure dialog: "+String.Join("; ",unexpectedDialogs));
                    if(gridView!=null)gridView.CloseEditor();else vertical.CloseEditor();Pump();
                    Check(target.ModelValue().IsNumeric&&Math.Abs(target.ModelValue().NumericValue-changed)<1e-10,"actual DIT posts to native owner");
                    Check(target.Value.Equals(template),"actual DIT leaves presentation workbook input untouched");
                    var undone=model.ChangeManager.Undo();Pump();Check(undone.BSuccess&&target.ModelValue().Equals(beforeValue),"actual DIT Undo restores selected native input");
                }
            }
            Check(model.IsDirty&&model.ChangeManager.GetHistoryTable().Rows.Count>=2,"real model history and dirty state retained");
            if(saveLifecycle){
                var cell=book.Worksheets["Covenant Assumptions"].Cells["D8"];double old=cell.ModelValue().NumericValue;
                var posted=model.ChangeManager.ProcessEngineCommand(new[]{new DataChangeEvent{ModelID=model.ModelID,WSName=cell.Worksheet.Name,CellAddress="D8",DataFormat="N",ChangedValue=old+0.125,Description="Real model save lifecycle"}},new[]{WorkbookValuePermission.UnlockedCell},"Real model save lifecycle");
                Check(posted.BSuccess,"real model lifecycle edit admitted");
                string recovery=(string)app.GetType("Abovo.RecoveryBackupStore").GetMethod("Write",F).Invoke(null,new object[]{model});
                Check(model.IsDirty&&model.ChangeManager.CanUndo&&model.FileName==copy,"real model recovery retains dirty state, filename and Undo");
                using(var verify=new Workbook()){
                    verify.Options.CalculationMode=WorkbookCalculationMode.Manual;
                    using(var input=File.OpenRead(recovery))Check(verify.LoadDocument(input,DocumentFormat.Xlsm),"recovery verification workbook loads from a read-only stream");
                    Check(verify.Worksheets[cell.Worksheet.Name].Cells["D8"].Value.NumericValue==old+0.125,"independent recovery read contains current native edit");
                    Check(verify.CustomXmlParts.Any(p=>p.CustomXmlPartDocument.OuterXml.Contains("Real model save lifecycle")),"real model recovery retains current history XML");
                    Check((string)app.GetType("Abovo.RecoveryBackupStore").GetMethod("ReadSource",F).Invoke(null,new object[]{recovery})==copy,"real model recovery identifies original filename");
                }
                // Invoke the shared application save operation directly here
                // so a failure escapes with its stack instead of blocking this
                // unattended fixture on the public wrapper's error MessageBox.
                Check((bool)Call(model,"SaveNativeWorkbookTo",copy),"normal application Save operation routes to the native owner");
                owners.Add((WorkbookCalculationSession)Field(model.ChangeManager,"engineTrial"));
                Check(!model.IsDirty&&cell.ModelValue().NumericValue==old+0.125&&model.ChangeManager.CanUndo,"real Demo Save retains current values and Undo");
                Check(model.ChangeManager.Undo().BSuccess&&model.IsDirty&&cell.ModelValue().NumericValue==old,"real Demo Undo after Save restores native input");
                string another=Path.Combine(folder,"native-save-as"+Path.GetExtension(copy));
                Check((bool)Call(model,"SaveNativeWorkbookTo",another),"normal application Save As operation routes to the native owner");
                owners.Add((WorkbookCalculationSession)Field(model.ChangeManager,"engineTrial"));
                Check(!model.IsDirty&&model.FileName==another&&model.ChangeManager.CanRedo,"real Demo Save As adopts verified filename and keeps Redo");
                Check(model.ChangeManager.Redo().BSuccess&&model.IsDirty&&cell.ModelValue().NumericValue==old+0.125,"real Demo Redo after Save As updates current owner");
                foreach(var target in new[]{copy,another})using(var verify=new Workbook()){
                    verify.Options.CalculationMode=WorkbookCalculationMode.Manual;
                    using(var input=File.OpenRead(target))Check(verify.LoadDocument(input,DocumentFormat.Xlsb),"independent saved workbook loads from a read-only stream");
                    Check(verify.Worksheets[cell.Worksheet.Name].Cells["D8"].Value.NumericValue==(target==copy?old+0.125:old),"independent saved file retains its committed input, not unsaved Redo");
                }
            }
        }
        finally
        {
            // Explicit test discard: normal model close must own native cleanup.
            // The final close is only an idempotent/failure safety net for the test.
            var active=model.ChangeManager!=null?(WorkbookCalculationSession)Field(model.ChangeManager,"engineTrial"):session;
            try{
                FileManager.CloseModel(opened.IntegerReturn);
                if(active!=null)Check(active.NativeCleanupCompletion!=null&&active.NativeCleanupCompletion.Status==TaskStatus.RanToCompletion,"normal model close completes native owner cleanup");
            }finally{if(active!=null)active.CloseAsync().GetAwaiter().GetResult();if(session!=null)session.CloseAsync().GetAwaiter().GetResult();}
        }
        NativeProcessChecks.RequireOwnedProcesses(owners).GetAwaiter().GetResult();
        Console.WriteLine("DIT_NATIVE ASSERTIONS="+checks);
    }
}
