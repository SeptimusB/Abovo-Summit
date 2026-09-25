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
    internal static Task Run(string source,bool excel)
    {
        var completion=new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread=new Thread(()=>{try{RunOnOwner(source,excel);completion.SetResult(true);}catch(Exception e){completion.SetException(e);}}){IsBackground=true};
        thread.SetApartmentState(ApartmentState.STA);thread.Start();return completion.Task;
    }
    static void RunOnOwner(string source,bool excel)
    {
        Thread.CurrentThread.CurrentCulture=CultureInfo.GetCultureInfo("en-GB");
        var before=NativeProcessChecks.ExcelIds();
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
            session=WorkbookCalculationSession.OpenAsync(copy,new WorkbookEngineOptions(excel?WorkbookEnginePreference.ExcelRequired:WorkbookEnginePreference.DevExpressOnly,true,true,120000,true)).GetAwaiter().GetResult();
            var initial=session.CalculateAndReadAsync(0,WorkbookCalculationKind.Rebuild,new[]{new WorkbookReadArea("Check Sheet",0,0,1,1)}).GetAwaiter().GetResult();
            Call(model.ChangeManager,"BindEngineEditingTrial",session,initial);
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
        }
        finally
        {
            // Explicit test discard: nothing is saved; normal production close
            // ownership is a separate gate, not simulated by clearing IsDirty.
            FileManager.CloseModel(opened.IntegerReturn);
            if(session!=null)session.CloseAsync().GetAwaiter().GetResult();
        }
        NativeProcessChecks.RequireOriginalProcesses(before).GetAwaiter().GetResult();
        Console.WriteLine("DIT_NATIVE ASSERTIONS="+checks);
    }
}
