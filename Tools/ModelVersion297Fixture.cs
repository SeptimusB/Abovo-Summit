using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using DevExpress.XtraEditors;
using DevExpress.Spreadsheet;

// Private workbook and non-activating hidden host. No global keyboard/mouse input.
class ModelVersion297Fixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 [DllImport("user32.dll")]static extern IntPtr SendMessage(IntPtr h,int m,IntPtr w,IntPtr l);
 [DllImport("user32.dll")]static extern bool ShowWindow(IntPtr h,int n);
 static int checks;
 static object Field(object o,string n){var f=o.GetType().GetField(n,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(n,F).GetValue(o,null);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static void Check(bool ok,string s){if(!ok)throw new Exception(s);Console.WriteLine("PASS "+(++checks)+" "+s);}
 static IEnumerable<Control> Controls(Control c){foreach(Control child in c.Controls){yield return child;foreach(var d in Controls(child))yield return d;}}
 [STAThread]static int Main(string[] args){try{
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  var path=Path.Combine(args[1],"private-model-version.xlsb");File.Copy(args[2],path);var open=files.GetMethod("OpenModel");
  dynamic result=open.Invoke(null,new object[]{path,new FileInfo(path),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});Check(!result.BError,"Private workbook opens");
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)result.IntegerReturn);IWorkbook w=model.WB;
  try{using(var host=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})){
   host.Opacity=0;host.ShowInTaskbar=false;host.ClientSize=new System.Drawing.Size(1600,950);ShowWindow(host.Handle,4);Application.DoEvents();
   Call(host,"ShowInterface",(int)model.ModelID,0,false,"None",null,-1);Application.DoEvents();var dit=(Control)Field(host,"ActiveInterface");
   var editors=Controls(dit).OfType<TextEdit>().Where(c=>c.GetType().Name=="AbovoDETextEdit").ToArray();
   var versionRange=w.DefinedNames.GetDefinedName("ModelVersion").Range;var companyRange=w.DefinedNames.GetDefinedName("SelectTrust").Range;
   var version=editors.First(c=>string.Equals((string)Field(c,"TargetCell"),versionRange[0,0].GetReferenceA1(),StringComparison.OrdinalIgnoreCase));
   var company=editors.First(c=>string.Equals((string)Field(c,"TargetCell"),companyRange[0,0].GetReferenceA1(),StringComparison.OrdinalIgnoreCase));
   Check(version.Properties.ReadOnly&&(bool)Field(version,"IsReadOnly"),"Model Version applies the existing XML read-only contract");
   Check(version.Enabled&&version.Text==versionRange[0,0].DisplayText,"Model Version stays enabled and displays workbook value");
   Check(!company.Properties.ReadOnly,"Company name remains editable");
   var oldText=version.Text;var oldValue=versionRange[0,0].Value;bool oldDirty=model.IsDirty;long oldRevision=Field(model,"UserChangeRevision") is long?(long)Field(model,"UserChangeRevision"):0;
   version.SelectAll();var target=Controls(version).OfType<TextBoxBase>().FirstOrDefault();Check(target!=null,"Native text child exists for input rejection test");
   foreach(char ch in "9999.999")SendMessage(target.Handle,0x102,(IntPtr)ch,IntPtr.Zero);Application.DoEvents();
   Check(version.Text==oldText,"Native character messages cannot edit Model Version");
   Check(versionRange[0,0].Value.Equals(oldValue),"Workbook Model Version unchanged");
   Check((bool)model.IsDirty==oldDirty&&(long)Field(model,"UserChangeRevision")==oldRevision,"Rejected typing does not dirty the model or add user revisions");
   target.SelectAll();Check(target.SelectedText==oldText,"Read-only native text remains selectable for copy");host.Close();
  }}finally{files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});}
  Console.WriteLine("PASS TOTAL="+checks);return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
}
