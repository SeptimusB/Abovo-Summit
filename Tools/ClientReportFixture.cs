using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraVerticalGrid;
using DevExpress.XtraVerticalGrid.Rows;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Columns;
using DevExpress.Spreadsheet;

public static class ClientReportFixture {
    static readonly System.Collections.Generic.List<Tuple<string,string,double>> savedInputs=new System.Collections.Generic.List<Tuple<string,string,double>>();
    const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
    static object Field(object o,string name){var field=o.GetType().GetField(name,F);return field!=null?field.GetValue(o):o.GetType().GetProperty(name,F).GetValue(o,null);}
    static object Call(object o,string name,params object[] args){return o.GetType().GetMethod(name,F).Invoke(o,args);}
    static void Check(bool value,string message){if(!value)throw new Exception(message);Console.WriteLine("PASS: "+message);}
    static string Child(XElement element,string name){return element.Elements().Where(x=>x.Name.LocalName==name).Select(x=>x.Value).FirstOrDefault();}
    static void Inspect(IWorkbook w){
        foreach(var name in w.DefinedNames.Where(n=>n.Name.IndexOf("SHG",StringComparison.OrdinalIgnoreCase)>=0&&n.Name.IndexOf("In",StringComparison.OrdinalIgnoreCase)>=0))Console.WriteLine("NAME "+name.Name+"="+name.RefersTo);
        foreach(string n in new[]{"SHGProfileIn","SHGMethodsIn","Rep_DevBP_01c","Rep_DevBP_01d","Rep_Jour_01","Rep_OCA_01"}){
            var range=w.DefinedNames.GetDefinedName(n).Range;Console.WriteLine("RANGE "+n+"="+range.GetReferenceA1());
            var first=range.Worksheet.Cells[range.TopRowIndex,range.LeftColumnIndex];Console.WriteLine("CELL "+first.GetReferenceA1()+" value="+first.Value+" format="+first.NumberFormat+" locked="+first.Protection.Locked);
            foreach(var c in range.Worksheet.Range.FromLTRB(0,range.TopRowIndex,Math.Min(5,range.RightColumnIndex),range.TopRowIndex).ExistingCells)Console.WriteLine("LABEL "+c.GetReferenceA1()+"="+c.DisplayText);
            foreach(var validation in range.Worksheet.DataValidations.GetDataValidations(first))Console.WriteLine("VALIDATION "+n+" formula="+validation.Criteria.FormulaInvariant);
        }
    }
    [STAThread] public static int Main(string[] args){try{
        AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var path=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(path)?Assembly.LoadFrom(path):null;};
        using(var w=new Workbook()){
            w.Options.CalculationMode=WorkbookCalculationMode.Manual;w.Options.Import.ThrowExceptionOnInvalidDocument=true;
            Check(w.LoadDocument(args[2]),"Native read-only inspection load");Inspect(w);
            if(args[3]=="True")return 0;
            var xml=XDocument.Load(Path.Combine(args[1],"Structure.xml"));
            var identified=xml.Descendants().Single(x=>x.Name.LocalName=="ChildStructure"&&Child(x,"CSName")=="Development Details");
            var grant=identified.Elements().Single(x=>x.Name.LocalName=="CSInterfaceSection"&&Child(x,"ISName")=="Grant/HFG");
            Func<string,XElement> field=repo=>grant.Descendants().Where(x=>x.Name.LocalName=="CellRangeDataSource").Single(x=>x.Descendants().Any(r=>r.Name.LocalName=="RepositaryItemID"&&r.Value==repo));
            string profile=Child(field("Rep_GrantMod"),"NRDSName"),basis=Child(field("Rep_SHGCalcBas"),"NRDSName");
            Check(profile!=basis,"P067 SHG profiling and calculation basis bind to separate named ranges");
            Check(profile=="SHGProfileIn"&&basis=="SHGMethodsIn","P067 bindings agree with master input definitions");
            var period=identified.Descendants().Single(x=>x.Name.LocalName=="CellRangeDataSource"&&Child(x,"NRDSName")=="Rep_DevBP_01d");
            Check(period.Descendants().Single(x=>x.Name.LocalName=="FieldName").Value.Replace("vblf ","")=="Period Units into Mgmt (to)","P062 management-period caption matches the master, not the works-completion row");
            var journals=xml.Descendants().Single(x=>x.Name.LocalName=="ChildStructure"&&Child(x,"CSName")=="Journal Assumptions");
            Check(journals.Descendants().Single(x=>x.Name.LocalName=="DataFieldDefinition"&&Child(x,"FieldName")=="Year").Elements().Single(x=>x.Name.LocalName=="DataFormat").Value=="I","Journal years remain integer inputs");
        }
        ExerciseEditors(args);
        return 0;
    }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
    static void ExerciseEditors(string[] args){
        var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
        app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
        var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
        string copy=Path.Combine(args[1],"private-client-report.xlsb");File.Copy(args[2],copy);
        var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
        Check(!result.BError,"Opened private model for real DIT regression");
        dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)result.IntegerReturn);
        try{
            using(var form=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})){
                form.Opacity=0;form.ShowInTaskbar=false;form.ClientSize=new Size(1750,1000);form.Show();Application.DoEvents();
                Call(form,"ShowInterface",0,26,false,"None",null,-1);
                var dit=(Control)Field(form,"ActiveInterface");
                dynamic tabs=Field(dit,"XtraTabControlNewGIT");
                int grantTab=-1;for(int i=0;i<tabs.TabPages.Count;i++)if(Convert.ToString(tabs.TabPages[i].Text).Trim()=="Grant/HFG"){grantTab=i;break;}
                Check(grantTab>=0,"Grant/HFG tab located by actual displayed caption");
                Call(dit,"BuildSection",grantTab,false,false);tabs.SelectedTabPageIndex=grantTab;
                Application.DoEvents();Application.RaiseIdle(EventArgs.Empty);Application.DoEvents();
                var positions=((IEnumerable)Call(dit,"NavigationPositions")).Cast<object>().ToList();
                Console.WriteLine("POSITIONS="+positions.Count+" tabs="+tabs.TabPages.Count+" selected="+tabs.SelectedTabPage.Text);
                foreach(var p in positions.GroupBy(p=>Field(p,"VRow")??Field(p,"Column")).Select(g=>g.First()))Console.WriteLine("EDITOR="+(Field(p,"VRow")!=null?Caption(p):Convert.ToString(Field(Field(p,"Column"),"Caption"))));
                var profile=positions.First(p=>Caption(p).Contains("SHG Profiling"));
                var basis=positions.First(p=>Caption(p).Contains("SHG Calculation Basis"));
                IWorkbook w=(IWorkbook)model.WB;var sheet=w.Worksheets["Development BP Assumptions"];
                int profileCol=w.DefinedNames.GetDefinedName("SHGProfileIn").Range.LeftColumnIndex+Record(profile);
                int basisCol=w.DefinedNames.GetDefinedName("SHGMethodsIn").Range.LeftColumnIndex+Record(basis);
                var profileCell=sheet.Cells[121,profileCol];var basisCell=sheet.Cells[124,basisCol];
                var oldProfile=profileCell.Value;var oldBasis=basisCell.Value;bool protection=sheet.IsProtected;
                var profileColumn=(GridColumn)Field(profile,"Column");profileColumn.Width=333;int expectedWidth=profileColumn.Width;
                Edit(dit,profile);Check(!profileCell.Value.Equals(oldProfile)&&basisCell.Value.Equals(oldBasis),"P067 editing Profiling changes only the correct workbook input");
                Console.WriteLine("COLUMN_WIDTH before="+expectedWidth+" after="+profileColumn.Width);
                Check(profileColumn.Width==expectedWidth,"Shared DIT edit preserves the existing column width");
                var newProfile=profileCell.Value;
                Edit(dit,basis);Check(!basisCell.Value.Equals(oldBasis)&&profileCell.Value.Equals(newProfile),"P067 editing Calculation Basis leaves Profiling unchanged");
                Check(model.IsDirty&&model.ChangeManager.CanUndo,"Real dropdown edits mark dirty and retain Undo");
                dynamic undo=model.ChangeManager.Undo();Check(!undo.BError&&basisCell.Value.Equals(oldBasis)&&profileCell.Value.Equals(newProfile),"Undo reverses only the basis edit");
                undo=model.ChangeManager.Undo();Check(!undo.BError&&profileCell.Value.Equals(oldProfile)&&basisCell.Value.Equals(oldBasis),"Undo restores both original independent inputs");
                Check(sheet.IsProtected==protection,"Input worksheet protection state preserved");
                for(int i=0;i<3;i++){Edit(dit,profile);Check(profileColumn.Width==expectedWidth,"Repeated dropdown edits preserve width, iteration "+(i+1));undo=model.ChangeManager.Undo();Check(!undo.BError,"Repeated edit Undo succeeds");}
                DecimalInput(form,model,41,"Amount",2345.67);
                DecimalInput(form,model,42,"Amount",1234.56);
                string saved=Path.Combine(args[1],"decimal-inputs.xlsb");
                Check((bool)model.SaveFileAsTo(saved,true),"Monetary inputs save through normal Summit XLSB policy to a separate file");
                using(var reopened=new Workbook()){
                    reopened.Options.CalculationMode=WorkbookCalculationMode.Manual;Check(reopened.LoadDocument(saved),"Saved monetary-input workbook reopens");
                    foreach(var input in savedInputs)Check(Math.Abs(reopened.Worksheets[input.Item1].Cells[input.Item2].Value.NumericValue-input.Item3)<1e-9,"Saved/reopened decimal preserved: "+input.Item1+"!"+input.Item2);
                }
                form.Close();
            }
        }finally{files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});}
    }
    static int Record(object point){return Field(point,"VRow")!=null?(int)Field(point,"Record"):((GridView)Field(point,"View")).GetDataSourceRowIndex((int)Field(point,"RowHandle"));}
    static string Caption(object point){return System.Text.RegularExpressions.Regex.Replace(Field(point,"VRow")!=null?((EditorRow)Field(point,"VRow")).Properties.Caption:((GridColumn)Field(point,"Column")).Caption,@"\s+"," ").Trim();}
    static void Edit(object dit,object point){
        Call(dit,"ActivateEditorPosition",point);Application.DoEvents();
        var grid=Field(point,"Host") as VGridControl;var view=Field(point,"View") as GridView;
        var editor=(grid!=null?grid.ActiveEditor:view.ActiveEditor) as ComboBoxEdit;
        Check(editor!=null,"Native SHG dropdown opens");
        var next=editor.Properties.Items.Cast<object>().First(v=>!String.IsNullOrWhiteSpace(Convert.ToString(v))&&Convert.ToString(v)!="<Blank>"&&Convert.ToString(v)!=Convert.ToString(editor.EditValue));
        editor.EditValue=next;
        Check(grid!=null?grid.PostEditor():view.PostEditor()&&view.UpdateCurrentRow(),"Native SHG dropdown commits through DIT");
        if(grid!=null)grid.CloseEditor();else view.CloseEditor();Application.DoEvents();
    }
    static void DecimalInput(Form form,dynamic model,int csid,string heading,double value){
        Call(form,"ShowInterface",0,csid,false,"None",null,-1);Application.DoEvents();Application.RaiseIdle(EventArgs.Empty);Application.DoEvents();
        var dit=(Control)Field(form,"ActiveInterface");
        var positions=((IEnumerable)Call(dit,"NavigationPositions")).Cast<object>().ToList();
        var point=positions.First(p=>Field(p,"View")!=null&&Caption(p).Contains(heading));
        var view=(GridView)Field(point,"View");var column=(GridColumn)Field(point,"Column");
        dynamic source=view.GridControl.DataSource;dynamic pres=Field(dit,"DataPres");
        int columnIndex=(int)Call(dit,"GetGridColumnIndex",column);
        dynamic cellPoint=pres.DataSets[(int)source.UBSTag.DSIndex].DataRows[Record(point)].DataCells[columnIndex];
        IWorkbook w=(IWorkbook)model.WB;var sheet=w.Worksheets[(string)cellPoint.SourceSheet];var cell=sheet.Cells[(string)cellPoint.SourceAddress];
        var original=cell.Value;string format=cell.NumberFormat;bool protection=sheet.IsProtected;
        column.Width=347;int expectedWidth=column.Width;
        Console.WriteLine("DECIMAL_TARGET CSID="+csid+" cell="+sheet.Name+"!"+cell.GetReferenceA1()+" format="+format+" locked="+cell.Protection.Locked);
        Call(dit,"ActivateEditorPosition",point);Application.DoEvents();Check(view.ActiveEditor!=null,"Decimal test opens actual monetary editor");
        view.ActiveEditor.EditValue=value;
        Check(view.PostEditor()&&view.UpdateCurrentRow(),"Decimal monetary entry posts");view.CloseEditor();Application.DoEvents();
        Console.WriteLine("DECIMAL_VALUE requested="+value+" stored="+cell.Value);
        Check(cell.Value.IsNumeric&&Math.Abs(cell.Value.NumericValue-value)<1e-9,"Monetary input retains both decimal places: CSID="+csid);
        Check(column.Width==expectedWidth,"Numeric edit preserves column width: CSID="+csid);
        Check(cell.NumberFormat==format&&sheet.IsProtected==protection,"Decimal edit preserves source formatting and worksheet protection");
        dynamic undo=model.ChangeManager.Undo();Check(!undo.BError&&cell.Value.Equals(original),"Decimal input Undo restores original value");
        dynamic redo=model.ChangeManager.Redo();Check(!redo.BError&&Math.Abs(cell.Value.NumericValue-value)<1e-9,"Decimal input Redo restores precise value");
        savedInputs.Add(Tuple.Create(sheet.Name,cell.GetReferenceA1(),value));
    }
}
