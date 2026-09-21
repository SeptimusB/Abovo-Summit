using System;
using System.IO;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Nodes;

public static class StructureAuthoringFixture {
    const BindingFlags Flags=BindingFlags.Instance|BindingFlags.Public|BindingFlags.NonPublic;
    [STAThread] public static int Main(string[] args) {
        try {
            AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{
                string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");
                return File.Exists(p)?Assembly.LoadFrom(p):null;
            };
            Application.EnableVisualStyles();
            var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
            app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
            var files=app.GetType("Abovo.FileManager");
            files.GetMethod("Initialise").Invoke(null,new object[]{null});
            string scratch=null;
            using(var form=(Form)Activator.CreateInstance(app.GetType("StructureManagerForm"))) {
                form.Opacity=0;form.ShowInTaskbar=false;form.Show();Application.DoEvents();
                Call(form,"LoadWorkbook",args[1]);
                dynamic session=Field(form,"session");
                dynamic model=session.Model;
                scratch=(string)model.FileName;
                Check(scratch!=args[1] && File.Exists(scratch),"Isolated workbook copy");
                Check((bool)model.ModelSpreadsheetControl.ReadOnly,"Spreadsheet readonly");
                Check(!((bool)model.ModelSpreadsheetControl.Options.Behavior.SaveAllowed) && !((bool)model.ModelSpreadsheetControl.Options.Behavior.SaveAsAllowed) && !((bool)model.ModelSpreadsheetControl.Options.Behavior.OpenAllowed),"Native save/save-as/open commands disabled");
                dynamic draft=Field(form,"draft");
                XDocument doc=(XDocument)draft.Document;
                var group=doc.Root.Elements("GroupStructure").First();
                var child=group.Elements("ChildStructure").First(x=>(string)x.Element("CSID")=="1");
                var binding=child.Descendants("CellRangeDataSource").First(x=>x.Element("NRDSName")!=null && (string)x.Element("NRDSName")!="CR");
                var field=binding.Elements("DataFieldDefinition").First();
                Set(form,"selected",field);
                Call(form,"ReadProperties");
                Call(form,"RefreshPreview",true);
                Call(form,"LocateBinding");
                Application.DoEvents();
                var preview=(Control)Field(form,"preview");
                Check(preview!=null && preview.GetType().Name=="DataInterfaceTemplate","Actual DIT preview");
                var grids=Descendants(preview).OfType<GridControl>().ToList();
                Check(grids.Count>0,"Native grids built");
                foreach(var grid in grids)foreach(var v in grid.ViewCollection.OfType<GridView>())Check(!v.OptionsBehavior.Editable,"Preview grid readonly");
                IWorkbook book=(IWorkbook)model.WB;
                var range=book.DefinedNames.GetDefinedName((string)binding.Element("NRDSName")).Range;
                Check(book.Worksheets.ActiveWorksheet.Name==range.Worksheet.Name,"Tree to worksheet link");
                string before=doc.ToString(SaveOptions.DisableFormatting);
                string originalName=(string)field.Element("FieldName");
                draft.Apply(field,new Dictionary<string,string>{{"FieldName","Trial description"}},book);
                Call(form,"RefreshPreview",true);
                preview=(Control)Field(form,"preview");
                Check(Descendants(preview).OfType<GridControl>().SelectMany(g=>g.ViewCollection.OfType<GridView>()).Any(v=>v.Columns.Cast<DevExpress.XtraGrid.Columns.GridColumn>().Any(c=>c.Caption.Contains("Trial description"))),"Edited caption rendered by DIT");
                var sameBinding=binding;
                string prior=doc.ToString(SaveOptions.DisableFormatting);
                ExpectFailure(()=>draft.Apply(binding,new Dictionary<string,string>{{"NRDSName","__MISSING_RANGE__"}},book),"Invalid binding");
                Check(doc.ToString(SaveOptions.DisableFormatting)==prior && ReferenceEquals(field.Parent,sameBinding),"Atomic rollback preserves tree identities");
                ExpectFailure(()=>draft.Apply(field,new Dictionary<string,string>{{"DataFormat","UNKNOWN-TYPE"}},book),"Invalid field type rejected");
                draft.Apply(binding,new Dictionary<string,string>{{"NRDSName",(string)binding.Element("NRDSName")},{"Worksheet",range.Worksheet.Name}},book);
                Call(form,"RefreshPreview",true);
                Check(Descendants((Control)Field(form,"preview")).OfType<GridControl>().Any(),"Valid binding rebuild");
                var invalidType=new XDocument(doc);invalidType.Root.Element("GroupStructure").Element("ChildStructure").Element("CSID").Value="1";
                ExpectFailure(()=>app.GetType("Abovo.StructureAuthoringDraft").GetMethod("Validate").Invoke(null,new object[]{invalidType}),"Duplicate IDs");
                ExpectFailure(()=>Activator.CreateInstance(app.GetType("Abovo.StructureAuthoringDraft"),new object[]{"<!DOCTYPE a [<!ENTITY e SYSTEM 'file:///invalid'>]><Abovo_Model_Def>&e;</Abovo_Model_Def>"}),"DTD prohibited");
                doc.Root.Add(new XComment("unknown metadata preservation test"));
                doc.Root.Add(new XElement("FutureMetadata",new XAttribute("revision","test"),new XElement("Opaque","preserve")));
                string output=Path.Combine(args[2],"Structure-trial.xml");
                draft.SaveNew(output);
                var reopened=(dynamic)Activator.CreateInstance(app.GetType("Abovo.StructureAuthoringDraft"),new object[]{output});
                Check(XNode.DeepEquals(doc.Root,((XDocument)reopened.Document).Root),"Draft XML roundtrip including unknown metadata");
                ExpectFailure(()=>draft.SaveNew(output),"Overwrite blocked");
                Check(!draft.IsDirty,"Saved draft clean");
                ExpectFailure(()=>Activator.CreateInstance(app.GetType("Abovo.StructureAuthoringSession"),new object[]{args[1]}),"Shared-model session refused");
                var change=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));
                // Exercise the central change guard without relying only on disabled editors.
                var point=book.Worksheets["Stock Assumptions"].Cells["B10"];
                string pointBefore=point.Value.ToString();
                SetMember(change,"WSName",point.Worksheet.Name);SetMember(change,"CellAddress",point.GetReferenceA1());SetMember(change,"ChangedValue","SHOULD NOT WRITE");SetMember(change,"DataFormat","S");
                dynamic rejected=model.ChangeManager.ProcessChange((dynamic)change);
                Check(!rejected.BSuccess && point.Value.ToString()==pointBefore,"Central single edit guard");
                var changedArray=Array.CreateInstance(change.GetType(),1);changedArray.SetValue(change,0);
                dynamic rejectedBatch=model.ChangeManager.ProcessChanges((dynamic)changedArray,"Trial");
                Check(!rejectedBatch.BSuccess && point.Value.ToString()==pointBefore,"Central paste guard");
                preview=(Control)Field(form,"preview");
                var gridWithData=Descendants(preview).OfType<GridControl>().First(g=>g.MainView is GridView && ((GridView)g.MainView).DataRowCount>0);
                var view=(GridView)gridWithData.MainView;
                view.FocusedRowHandle=0;view.FocusedColumn=view.VisibleColumns[0];view.ClearSelection();view.SelectCell(0,view.FocusedColumn);
                object[] sourceArgs={gridWithData,null,null};
                Check((bool)preview.GetType().GetMethod("TryGetAuthoringSource",Flags).Invoke(preview,sourceArgs),"DIT source lookup");
                Console.WriteLine("DIT selected cell: "+sourceArgs[1]+"!"+sourceArgs[2]);
                Call(form,"PreviewSourceSelected",sourceArgs[1],sourceArgs[2]);
                Check(book.Worksheets.ActiveWorksheet.Name==(string)sourceArgs[1],"DIT to spreadsheet link");
                foreach(int width in new[]{1280,1900,3000}) {
                    form.ClientSize=new Size(width,950);Application.DoEvents();
                    Check(((Control)Field(form,"workbookHost")).Width>150 && ((Control)Field(form,"previewHost")).Width>150,"Pane widths "+width);
                    if(width==1900)using(var bitmap=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(bitmap,new Rectangle(Point.Empty,form.Size));bitmap.Save(Path.Combine(args[2],"structure-manager.png"));}
                }
                // Switch between ordinary, mapped and non-linear-period DITs in this same session.
                foreach(string id in new[]{"0","28","1"}) {
                    var target=group.Elements("ChildStructure").First(x=>(string)x.Element("CSID")==id);
                    var sections=target.Elements("CSInterfaceSection").ToList();
                    Set(form,"selected",id=="0"?sections.First(x=>(string)x.Element("ISName")=="Check Sheet"):sections.Last());
                    Call(form,"ReadProperties");Call(form,"RefreshPreview",true);Application.DoEvents();
                    var current=(Control)Field(form,"preview");
                    Check(current!=null,"Preview route "+id);
                    foreach(var g in Descendants(current).OfType<GridControl>())foreach(var v in g.ViewCollection.OfType<GridView>())Check(!v.OptionsBehavior.Editable,"Readonly after tab change "+id);
                    Check((int)model.WBCalcEngine.ActiveObjectCount==1,"One active preview registration "+id);
                }
                draft.IsDirty=false;
                Call(form,"ReleaseSession");
                Check(!File.Exists(scratch),"Temporary workbook removed");
                Check((int)files.GetField("OpenModelCount").GetValue(null)==0,"Private registry released");
                // Reopen after all disposal callbacks; this catches stale private registry state.
                Call(form,"LoadWorkbook",args[1]);Call(form,"ReleaseSession");
                Check((int)files.GetField("OpenModelCount").GetValue(null)==0,"Reopen and close again");
            }
            Console.WriteLine("PASS: isolated copy, real readonly DIT, linked selection, caption edit/rebuild, atomic invalid binding rollback, secure XML roundtrip/new-file-only, central single/paste guards and cleanup.");
            return 0;
        }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}
    }
    static IEnumerable<Control> Descendants(Control c){foreach(Control child in c.Controls){yield return child;foreach(var d in Descendants(child))yield return d;}}
    static object Field(object o,string n){return o.GetType().GetField(n,Flags).GetValue(o);}
    static void Set(object o,string n,object v){o.GetType().GetField(n,Flags).SetValue(o,v);}
    static void SetMember(object o,string n,object v){var f=o.GetType().GetField(n,Flags);if(f!=null)f.SetValue(o,v);else o.GetType().GetProperty(n,Flags).SetValue(o,v,null);}
    static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,Flags).Invoke(o,a);}
    static void Check(bool b,string m){if(!b)throw new Exception(m);Console.WriteLine("PASS: "+m);}
    static void ExpectFailure(Action action,string label){bool failed=false;try{action();}catch{failed=true;}Check(failed,label);}
}
