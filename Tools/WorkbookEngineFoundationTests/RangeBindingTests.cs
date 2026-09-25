using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using DevExpress.Spreadsheet;
using RangeDataSource=DevExpress.XtraSpreadsheet.Model.RangeDataSource;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Abovo;
using Abovo.WorkbookEngines;

static class RangeBindingTests
{
    static int assertions;
    static void Check(bool ok,string message){if(!ok)throw new Exception(message);Console.WriteLine("RANGE PASS "+(++assertions)+" "+message);}
    static object[][] Rows(IList values,PropertyDescriptorCollection fields)
    {return values.Cast<object>().Select(row=>fields.Cast<PropertyDescriptor>().Select(field=>field.GetValue(row)).ToArray()).ToArray();}
    internal static async Task Run(string original)
    {
        var owners=new List<WorkbookCalculationSession>();
        string root=Path.Combine(Path.GetDirectoryName(original),"range-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);
        string source=Path.Combine(root,"range.xlsx");
        using(var book=new Workbook())
        {
            var s=book.Worksheets[0];s.Name="Data";
            s.Cells["B3"].Value="Name";s.Cells["C3"].Value="Amount";s.Cells["D3"].Value="Hidden";s.Cells["E3"].Value="Description";
            s.Cells["F3"].Value="Error";s.Cells["H1"].Value=2;s.Cells["H1"].Protection.Locked=false;
            s.Cells["B4"].Value="A";s.Cells["C4"].Formula="=H1*1.25";s.Cells["E4"].Value="=literal";s.Cells["F4"].Value=CellValue.ErrorDivisionByZero;
            s.Cells["B5"].Value="Hidden row";s.Cells["C5"].Value=900;
            s.Cells["B6"].Value="B";s.Cells["C6"].Value=10;s.Cells["G4"].Value=true;
            s.Columns[3].Visible=false;s.Rows[4].Visible=false;
            book.CalculateFull();book.SaveDocument(source,DocumentFormat.Xlsx);
        }
        foreach(var preference in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired})
        {
            var session=await WorkbookCalculationSession.OpenAsync(source,new WorkbookEngineOptions(preference,false,false,120000,true));
            owners.Add(session);
            using(var model=new EngineChangeManagerTests.Model(source))
            try
            {
                object[][] expected=null;string[] names=null;Type[] types=null;
                var options=new RangeDataSourceOptions{UseFirstRowAsHeader=true,PreserveFormulas=false,SkipHiddenRows=true,SkipHiddenColumns=true,EditingOptions=DataSourceEditingOptions.ReadOnly,DataSourceColumnTypeDetector=new NativeDetector()};
                await model.Do(()=>{
                    var range=model.Value.WB.Worksheets[0].Range["B3:G6"];
                    using(var baseline=(RangeDataSource)range.GetDataSource(options))
                    {
                        var fields=((ITypedList)baseline).GetItemProperties(null);names=fields.Cast<PropertyDescriptor>().Select(p=>p.Name).ToArray();types=fields.Cast<PropertyDescriptor>().Select(p=>p.PropertyType).ToArray();expected=Rows((IList)baseline,fields);
                    }
                    using(var normal=ModelRangeDataSource.Create(range,options))Check(Rows(normal,((ITypedList)normal).GetItemProperties(null)).SelectMany(x=>x).SequenceEqual(expected.SelectMany(x=>x)),"ordinary DevExpress binding is unchanged");
                    return true;
                });
                var result=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{new WorkbookReadArea("Data",0,0,1,1)});
                await model.Do(()=>{typeof(ModelChangeManagerV2).GetMethod("BindEngineEditingTrial",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(model.Manager,new object[]{session,result});return true;});
                ModelRangeDataSource data=null;PropertyDescriptorCollection props=null;object retained=null;
                await model.Do(()=>{
                    var range=model.Value.WB.Worksheets[0].Range["B3:G6"];data=ModelRangeDataSource.Create(range,options);props=((ITypedList)data).GetItemProperties(null);retained=data[0];
                    Check(data.HasCurrentValues&&data.Count==2,session.EngineName+" correct hidden row/column mapping");
                    Check(props.Cast<PropertyDescriptor>().Select(p=>p.Name).SequenceEqual(names)&&props.Cast<PropertyDescriptor>().Select(p=>p.PropertyType).SequenceEqual(types),"native analyser preserves empty header normalisation and types");
                    Check(Rows(data,props).SelectMany(x=>x).SequenceEqual(expected.SelectMany(x=>x)),"native values, blanks, error, Boolean and literal formula text match the baseline");
                    bool rejected=false;try{props[1].SetValue(data[0],999d);}catch(NotSupportedException){rejected=true;}Check(rejected,"descriptor cannot edit calculated values");
                    using(var form=new System.Windows.Forms.Form())
                    using(var grid=new DevExpress.XtraGrid.GridControl())
                    using(var view=new DevExpress.XtraGrid.Views.Grid.GridView(grid))
                    {
                        grid.MainView=view;form.Controls.Add(grid);grid.DataSource=data;form.CreateControl();grid.CreateControl();grid.ForceInitialize();view.PopulateColumns();
                        view.ActiveFilterString="[Name] = 'A'";Check(view.DataRowCount==1&&Equals(view.GetRowCellValue(0,"Amount"),2.5d),"actual GridView binds, filters and reads native numeric field");
                        var amount=view.Columns["Amount"];amount.SummaryItem.SummaryType=DevExpress.Data.SummaryItemType.Sum;view.OptionsView.ShowFooter=true;view.UpdateTotalSummary();
                        Check(Convert.ToDouble(amount.SummaryItem.SummaryValue)==2.5d,"native grid totals use current numeric data");grid.DataSource=null;
                    }
                    return true;
                });
                await model.Do(()=>{
                    var publish=typeof(ModelChangeManagerV2).Assembly.GetType("Abovo.ModelEngineView").GetMethod("Publish",BindingFlags.Static|BindingFlags.NonPublic);
                    publish.Invoke(null,new object[]{model.Value.WB,null});
                    var cell=model.Value.WB.Worksheets[0].Cells["C4"];
                    Check(!data.HasCurrentValues&&props[1].GetValue(retained)==null,"pending publication hides already-bound figures");
                    Check(cell.ModelPaintText()=="…","paint callback uses a pending marker without a stale value");
                    bool refused=false;try{cell.ModelValue();}catch(InvalidOperationException){refused=true;}
                    Check(refused,"strict reads remain unavailable during calculation");
                    publish.Invoke(null,new object[]{model.Value.WB,result});
                    Check(data.HasCurrentValues&&Equals(props[1].GetValue(retained),2.5d),"accepted publication restores current bound values");
                    return true;
                });
                await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{new WorkbookReadArea("Data",0,0,1,1)});
                await model.Do(()=>{Check(!data.HasCurrentValues&&props[1].GetValue(retained)==null,"same-revision new calculation invalidates old bound figures");data.Dispose();Check(props[1].GetValue(retained)==null,"retained disposed row cannot leak values");typeof(ModelChangeManagerV2).GetMethod("RecalculateEngine",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(model.Manager,new object[]{WorkbookCalculationKind.Full,false});return true;});
                await model.Do(()=>{
                    var change=new DataChangeEvent{ModelID=0,WSName="Data",CellAddress="H1",ChangedValue=4d,DataFormat="N",Description="Range dependency"};
                    var response=model.Manager.ProcessEngineCommand(new[]{change},new[]{WorkbookValuePermission.UnlockedCell},"Range dependency");
                    Check(response.BSuccess,"native input updates analyser precedent");
                    using(var refreshed=ModelRangeDataSource.Create(model.Value.WB.Worksheets[0].Range["B3:G6"],options)){
                        var fields=((ITypedList)refreshed).GetItemProperties(null);Check(Equals(fields[1].GetValue(refreshed[0]),5d),"recreated analyser reads changed native dependent result");
                    }
                    Check(model.Value.WB.Worksheets[0].Cells["H1"].Value.NumericValue==2&&model.Value.WB.Worksheets[0].Cells["C4"].HasFormula,"presentation map formulas and inputs were never rewritten");return true;
                });
            }
            finally{await session.CloseAsync();}
        }
        await NativeProcessChecks.RequireOwnedProcesses(owners);
        Console.WriteLine("RANGE ASSERTIONS="+assertions);
    }
    sealed class NativeDetector:IDataSourceColumnTypeDetector
    {
        public string GetColumnName(int index,int offset,CellRange range){return range[-1,offset].ModelDisplayText();}
        public Type GetColumnType(int index,int offset,CellRange range){return offset==1?typeof(double):typeof(string);}
    }
    internal static async Task Agl(string source)
    {
        var owners=new List<WorkbookCalculationSession>();
        object[][] reference=null;string[] referenceNames=null;
        foreach(var preference in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired})
        {
            var session=await WorkbookCalculationSession.OpenAsync(source,new WorkbookEngineOptions(preference,true,true,120000,true));
            owners.Add(session);
            using(var model=new EngineChangeManagerTests.Model(source))
            try
            {
                var result=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Rebuild,new[]{new WorkbookReadArea("Check Sheet",0,0,1,1)});
                await model.Do(()=>{
                    typeof(ModelChangeManagerV2).GetMethod("BindEngineEditingTrial",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(model.Manager,new object[]{session,result});
                    var detector=(IDataSourceColumnTypeDetector)Activator.CreateInstance(typeof(ModelChangeManagerV2).Assembly.GetTypes().Single(t=>t.Name=="BPIEAColumnDetectorV2"),true);
                    var ws=model.Value.WB.Worksheets["Transactional DB"];
                    var range=(ws.DefinedNames.GetDefinedName("Transactional_Records")??model.Value.WB.DefinedNames.GetDefinedName("Transactional_Records")).Range;
                    var options=new RangeDataSourceOptions{UseFirstRowAsHeader=true,PreserveFormulas=false,SkipHiddenRows=true,SkipHiddenColumns=true,EditingOptions=DataSourceEditingOptions.ReadOnly,DataSourceColumnTypeDetector=detector};
                    var timer=System.Diagnostics.Stopwatch.StartNew();
                    using(var data=ModelRangeDataSource.Create(range,options))
                    {
                        var properties=((ITypedList)data).GetItemProperties(null);var names=properties.Cast<PropertyDescriptor>().Select(p=>p.Name).ToArray();var actual=Rows(data,properties);
                        Console.WriteLine("AGL_RANGE engine="+session.EngineName+" rows="+data.Count+" columns="+properties.Count+" bindAndReadMs="+timer.ElapsedMilliseconds);
                        Check(data.HasCurrentValues&&data.Count>1000&&properties.Count>50,"actual AGL analyser range bound to current native result");
                        if(reference==null){reference=actual;referenceNames=names;}else{
                            Check(referenceNames.SequenceEqual(names)&&reference.Length==actual.Length,"actual AGL analyser schema matches between engines");
                            int differences=0,cells=0;
                            for(int r=0;r<actual.Length;r++)for(int c=0;c<actual[r].Length;c++){
                                var a=reference[r][c];var b=actual[r][c];cells++;
                                bool same=a is double&&b is double?Math.Abs((double)a-(double)b)<=Math.Max(1e-6,Math.Abs((double)a)*1e-10):Equals(a,b);
                                if(!same){differences++;if(differences<=8)Console.WriteLine("AGL_RANGE_DIFFERENCE row="+r+" column="+names[c]+" dx="+a+" excel="+b);}
                            }
                            Check(differences==0,"current native analyser parity across "+cells+" values");
                        }
                    }
                    return true;
                });
            }
            finally{await session.CloseAsync();}
        }
        await NativeProcessChecks.RequireOwnedProcesses(owners);
    }
}
