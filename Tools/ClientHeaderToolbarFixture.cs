using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;

class ClientHeaderToolbarFixture {
    const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
    static object Field(object o,string name) {var f=o.GetType().GetField(name,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(name,F).GetValue(o,null);}
    static object Call(object o,string name,params object[] args) {return o.GetType().GetMethod(name,F).Invoke(o,args);}
    static int checks;
    static void Check(bool ok,string message) {if(!ok)throw new Exception(message);checks++;Console.WriteLine("PASS: "+message);}
    static void Pump(){Application.DoEvents();}
    [STAThread] static int Main(string[] args) {try {
        AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
        var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
        app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
        var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
        string copy=Path.Combine(args[1],"private-headers.xlsb");File.Copy(args[2],copy);
        var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
        Check(!result.BError,"Private source copy opens");
        dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)result.IntegerReturn);
        try {Exercise(app,model);} finally {files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});}
        Console.WriteLine("PASS: "+checks+" checks; no source saved");return 0;
    } catch(Exception e){Console.Error.WriteLine(e);return 1;} }

    static void Exercise(Assembly app,dynamic model) {
        var pageType=Assembly.Load("DevExpress.XtraEditors.v25.2").GetType("DevExpress.XtraTab.XtraTabPage");
        using(var page=(Control)Activator.CreateInstance(pageType))
        using(var instance=(Control)Activator.CreateInstance(app.GetType("FileInstanceInterface"),new object[]{(int)model.ModelID}))
        using(var ffr=(Form)Activator.CreateInstance(app.GetType("FFRForm"),new object[]{(int)model.ModelID}))
        using(var group=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})) {
            string ffrSheetCaption=((Control)Field(ffr,"SheetCaption")).Text;
            page.Controls.Add(instance);Call(instance,"PopulateFileInfo");
            group.ShowInTaskbar=false;group.Opacity=0;group.ClientSize=new Size(1600,950);group.Show();Pump();
            Call(group,"ShowInterface",(int)model.ModelID,0,false,"None",null,-1);Pump();
            var dit=Field(group,"ActiveInterface");
            var panel=Field(dit,"WindowsUIButtonPanelActions");
            var buttons=((IEnumerable)Field(panel,"Buttons")).Cast<object>().ToList();
            var back=buttons.Single(b=>Convert.ToString(Field(b,"Tag"))=="Return");
            var history=buttons.Single(b=>Convert.ToString(Field(b,"Tag"))=="History");
            var separator=Field(dit,"ReturnSeparator");
            Check(Object.ReferenceEquals(buttons.Last(),back)&&Object.ReferenceEquals(buttons[buttons.Count-2],separator),"Return is last, after its own separator");
            Check(!(bool)Field(back,"Visible")&&!(bool)Field(separator,"Visible"),"Ordinary interface hides Return and its separator");
            var link=Activator.CreateInstance(app.GetTypes().Single(t=>t.Name=="ElementInterfaceLinkTag"));link.GetType().GetField("LinkReturnName").SetValue(link,"Test origin");
            Call(dit,"AddLink",link);
            Check((bool)Field(back,"Visible")&&(bool)Field(separator,"Visible"),"Linked interface shows Return and separator together");
            using(var picture=new Bitmap(((Control)panel).Width,((Control)panel).Height)) {
                ((Control)panel).DrawToBitmap(picture,new Rectangle(Point.Empty,picture.Size));
                picture.Save(Path.Combine(Path.GetDirectoryName(((Assembly.GetExecutingAssembly()).Location)),"linked-toolbar.png"));
            }
            Check(Convert.ToString(Field(back,"ToolTip"))=="Return to Test origin","Return retains the origin tooltip");
            Check(Field(Field(history,"ImageOptions"),"SvgImage")!=null,"History has vector artwork");
            Check(Field(Field(back,"ImageOptions"),"SvgImage")!=null,"Return has vector artwork");
            Check(!Object.ReferenceEquals(Field(Field(back,"ImageOptions"),"SvgImage"),Field(Field(history,"ImageOptions"),"SvgImage")),"Return and History use distinct artwork");
            Call(dit,"ClearLinks");Check(!(bool)Field(back,"Visible")&&!(bool)Field(separator,"Visible"),"Clearing links removes both conditional items");

            IWorkbook w=(IWorkbook)model.WB;var cell=w.DefinedNames.GetDefinedName("SelectTrust").Range[0,0];
            string oldName=(string)model.WBStructure.CompanyName;string suffix=group.Text.Substring(oldName.Length);
            string newName="Test <A & B> • Homes";int notifications=0;
            var changed=((object)model).GetType().GetEvent("MetadataChanged",F);
            EventHandler handler=(s,e)=>notifications++;
            changed.GetAddMethod(true).Invoke(model,new object[]{handler});
            Call(model,"RestoreRecoveryAutosaveHold",true);
            dynamic edit=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));edit.ModelID=(int)model.ModelID;edit.WSName=cell.Worksheet.Name;edit.CellAddress=cell.GetReferenceA1();edit.ChangedValue=newName;edit.DataFormat="S";edit.Description="Company metadata test";
            Check(model.ChangeManager.ProcessChange(edit).BSuccess,"Company-name change commits through production history service");Pump();
            Check((string)model.WBStructure.CompanyName==newName&&notifications==1,"Committed company name refreshes the model once");
            Check(group.Text==newName+suffix,"Company heading refresh preserves selected interface suffix");
            Check(ffr.Text.EndsWith(newName)&&((Control)Field(ffr,"SheetCaption")).Text==ffrSheetCaption,"FFR company title updates without replacing selected-sheet caption");
            Check(Convert.ToString(Field(instance,"MyCompanyName"))==newName&&page.Text.EndsWith(newName),"FileInstance and main tab refresh immediately");
            var html=((WebBrowser)Field(instance,"WebBrowserBPInfo")).DocumentText;
            Check(html.Contains("Test &lt;A &amp; B&gt;"),"FileInstance HTML escapes the new company name");
            Check((bool)Field(model,"CheckSheetWarningActive"),"Metadata refresh does not clear a Check Sheet warning");
            Check(model.ChangeManager.Undo().BSuccess,"Company-name Undo succeeds");Pump();
            Check((string)model.WBStructure.CompanyName==oldName&&group.Text==oldName+suffix&&notifications==2,"Undo updates cache and displayed heading");
            Check(ffr.Text.EndsWith(oldName),"Undo updates the open FFR company title");
            Check(model.ChangeManager.Redo().BSuccess,"Company-name Redo succeeds");Pump();
            Check((string)model.WBStructure.CompanyName==newName&&notifications==3,"Redo refreshes company metadata");
            Call(model,"RefreshModelMetadata");Check(notifications==3,"Unchanged metadata does not re-render views");
            changed.GetRemoveMethod(true).Invoke(model,new object[]{handler});

            var comboType=app.GetTypes().Single(t=>t.Name=="AbovoDEHeaderComboBox");
            using(var combo=(ComboBoxEdit)Activator.CreateInstance(comboType)) {
                comboType.GetField("ModelID").SetValue(combo,(int)model.ModelID);Call(combo,"InitialiseStandard","Rep_OrdinalYears");
                combo.EditValue=1;Console.WriteLine("YEAR REPOSITORY="+combo.Properties.GetDisplayText(1)+" LIVE="+combo.Text);
                Check(combo.Text.StartsWith("Year 1"),"Numeric year editor displays Year 1");
                string numericDisplay=combo.Text;combo.EditValue="Year 1";
                Check(combo.Text==numericDisplay,"Workbook-formatted Year 1 displays identically");
                var item=combo.Properties.Items[1];
                Check(Convert.ToString(Field(item,"StoredValue"))=="2"&&item.ToString().StartsWith("Year 2"),"Dropdown keeps numeric stored value separate from label");
                Check(Convert.ToString(Call(dit,"NormalizeInColumnEditorValue",item))=="2","Column editor converts selection back to numeric-year input");
            }
            using(var small=new Workbook()) {
                var c=small.Worksheets[0].Cells[0,0];c.Value=1;
                var formatter=app.GetType("Abovo.DataManager").GetMethod("RepeatingHeaderCaption",F);
                foreach(string caption in new[]{"","Year","Yr"})
                    Check((string)formatter.Invoke(null,new object[]{caption,"Rep_OrdinalYearsLess1",c})=="Year 1","Read-only first-year caption is consistent: '"+caption+"'");
                c.NumberFormat="\"Year \"0";
                Check((string)formatter.Invoke(null,new object[]{"Year","Rep_OrdinalYears",c})=="Year 1","Formatted first-year header avoids a duplicated Year prefix");
                Check(c.Value.NumericValue==1&&c.NumberFormat=="\"Year \"0","Presentation leaves year value and workbook format untouched");
            }
            group.Close();
        }
    }
}
