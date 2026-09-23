using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

public static class IntegrityOptionsFixture {
 const BindingFlags F=BindingFlags.Instance|BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic;
 static Assembly app;static string output;static Array models;static Type files;
 static object Field(object o,string n){return o.GetType().GetField(n,F).GetValue(o);}
 static object Prop(object o,string n){return o.GetType().GetProperty(n,F).GetValue(o,null);}
 static object Call(object o,string n,params object[] a){return(o as Type??o.GetType()).GetMethods(F).Single(m=>m.Name==n&&m.IsStatic==(o is Type)&&m.GetParameters().Length==a.Length).Invoke(o is Type?null:o,a);}
 static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("PASS: "+text);}
 static void Selection(int? requested,int expected,string label){
  using(var form=(Form)app.GetType("Abovo.ApplicationOptionsForm").GetConstructor(new[]{typeof(int?)}).Invoke(new object[]{requested})){
   dynamic selector=Field(form,"IntegrityPlan");var run=form.Controls.Find("RunIntegrityNow",true).Single();var path=form.Controls.Find("IntegrityTargetPath",true).Single();
   Check((int)selector.SelectedIndex==expected,label);
   Check(run.Enabled==(expected>=0),"Run is enabled only for a selected target");
   if(expected>=0){var plans=(System.Collections.IList)Field(form,"IntegrityPlans");dynamic target=plans[expected];Check(path.Text.Contains((string)target.FileName),"Full selected path shown");}
   form.Opacity=0;form.ShowInTaskbar=false;form.Show();dynamic tabs=form.Controls[0].Controls[0];tabs.SelectedTabPageIndex=2;Application.DoEvents();
   if(requested==1&&expected==1){
    using(var image=new System.Drawing.Bitmap(form.Width,form.Height)){form.DrawToBitmap(image,new System.Drawing.Rectangle(System.Drawing.Point.Empty,form.Size));image.Save(Path.Combine(output,"integrity-target.png"));}
    selector.SelectedIndex=0;Check(path.Text.Contains("New Blank.xlsb"),"Explicitly changing target updates path");
    selector.SelectedIndex=-1;Check(!run.Enabled,"Clearing selection prevents a check");
   }
   form.Close();
  }
 }
 [STAThread]public static int Main(string[] args){try{
  output=args[1];AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  using(var main=(Form)Activator.CreateInstance(app.GetType("FormMainScreen"))){
   files=app.GetType("Abovo.FileManager");var modelType=app.GetType("Abovo.FileManager+ExcelModel");var ctor=modelType.GetConstructors().Single();models=Array.CreateInstance(modelType,2);
   for(int i=0;i<2;i++){dynamic model=ctor.Invoke(new object[]{i,Enum.ToObject(ctor.GetParameters()[1].ParameterType,0)});model.FileName=Path.Combine(output,i==0?"New Blank.xlsb":"AGL.xlsb");models.SetValue(model,i);}
   files.GetField("ExcelModels").SetValue(null,models);files.GetField("ExcelModelCount").SetValue(null,1);
   var allocated=(Array)models.Clone();try{
    Selection(1,1,"Current DIT model 1 selected instead of first-opened Blank");Selection(0,0,"Explicit model 0 respected");Selection(null,-1,"Ambiguous context requires selection");Selection(99,-1,"Closed or unknown context cannot select another file silently");
    dynamic outer=Prop(main,"XtraTabControlMainNavigator"),tabs=Prop(main,"XtraTabControlModels");outer.SelectedTabPage=(dynamic)Prop(main,"XtraTabPageMainHABP");
    var tabType=tabs.TabPages[0].GetType();dynamic blank=Activator.CreateInstance(tabType),agl=Activator.CreateInstance(tabType);blank.Tag=0;agl.Tag=1;tabs.TabPages.Add(blank);tabs.TabPages.Add(agl);
    tabs.SelectedTabPage=agl;Check((int?)Call(main,"SelectedOptionsModelID")==1,"Main screen chooses visible AGL tab");
    main.GetType().GetField("ActiveModel",F).SetValue(main,1);tabs.SelectedTabPage=blank;Check((int?)Call(main,"SelectedOptionsModelID")==0,"Switching tabs wins over last-opened ActiveModel");
    tabs.SelectedTabPage=agl;agl.Tag=null;Check(Call(main,"SelectedOptionsModelID")==null,"Unbound main tab has no implicit target");
    dynamic dsa=Field(main,"DsaModelsTabControl"),dsaPage=Activator.CreateInstance(tabType);dsaPage.Tag=1;dsa.TabPages.Add(dsaPage);dsa.SelectedTabPage=dsaPage;outer.SelectedTabPage=(dynamic)Prop(main,"XtraTabPageEvolveDSA");Check((int?)Call(main,"SelectedOptionsModelID")==1,"DSA selection uses its own visible model tab");
    models.SetValue(null,0);Selection(null,0,"Single remaining plan can be selected unambiguously");models.SetValue(null,1);Selection(null,-1,"No open plan cannot start a check");
   }finally{foreach(dynamic model in allocated)if(model!=null)model.ModelSpreadsheetControl.Dispose();}
  }
  return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
