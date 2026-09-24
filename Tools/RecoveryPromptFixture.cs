using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Security.Cryptography;
using DevExpress.XtraEditors;
using DevExpress.Spreadsheet;

class RecoveryPromptFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Static|BindingFlags.Instance;
 static Type store;static int count;static string output;
 static object Call(string method,params object[] args){return store.GetMethod(method,F).Invoke(null,args);}
 static void Check(bool value,string message){if(!value)throw new Exception(message);count++;Console.WriteLine("PASS "+message);}
 static string Hash(string path){using(var sha=SHA256.Create())return BitConverter.ToString(sha.ComputeHash(File.ReadAllBytes(path)));}
 static IEnumerable<Control> Controls(Control root){foreach(Control c in root.Controls){yield return c;foreach(var d in Controls(c))yield return d;}}
 static void Recovery(string path,string source){using(var w=new Workbook()){w.Worksheets[0].Cells["A1"].Value=42;w.DocumentProperties.Custom["Abovo.Summit.RecoverySource"]=source;w.DocumentProperties.Custom["Abovo.Summit.RecoveryVersion"]="1";w.SaveDocument(path,DocumentFormat.Xlsm);}File.SetLastWriteTimeUtc(path,DateTime.UtcNow);}
 static void Reject(string source,string candidate,long length,DateTime time,string message){bool rejected=false;try{Call("DeleteRecovery",source,candidate,length,time);}catch(TargetInvocationException e){rejected=e.InnerException is IOException||e.InnerException is InvalidOperationException;}Check(rejected,message);}
 static string Choose(string source,params string[] captions){
  int step=0;Exception failure=null;var deadline=DateTime.UtcNow.AddSeconds(12);
  using(var timer=new Timer()){timer.Interval=120;timer.Tick+=(s,e)=>{
   try{var form=Application.OpenForms.Cast<Form>().FirstOrDefault(f=>f.Modal);if(form==null)return;
    if(DateTime.UtcNow>deadline)throw new Exception("Recovery prompt timed out: "+form.Text);
    if(step>=captions.Length)throw new Exception("Unexpected extra dialog: "+form.Text);
    var buttons=Controls(form).OfType<SimpleButton>().ToArray();
    var button=buttons.FirstOrDefault(b=>b.Text.Replace("&","")==captions[step]);
    if(button==null)throw new Exception("Missing button "+captions[step]+": "+String.Join(",",buttons.Select(b=>b.Text)));
    if(form.GetType().Name=="RecoveryOpenPrompt"){
     Check(buttons.Length==4,"Recovery discovery offers four explicit choices");
     using(var image=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(image,new Rectangle(Point.Empty,image.Size));image.Save(Path.Combine(output,"recovery-prompt.png"));}
    }
    step++;button.PerformClick();
   }catch(Exception ex){failure=ex;timer.Stop();foreach(var f in Application.OpenForms.Cast<Form>().ToArray()){f.DialogResult=DialogResult.Cancel;f.Close();}}
  };timer.Start();var answer=(string)Call("SelectOpenPath",null,source);timer.Stop();if(failure!=null)throw failure;Check(step==captions.Length,"Expected native dialog actions consumed");return answer;}
 }
 [STAThread] static int Main(string[] args){try{
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));store=app.GetType("Abovo.RecoveryBackupStore");output=args[1];
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  string source=Path.Combine(output,"private plan.xlsb");using(var w=new Workbook()){w.Worksheets[0].Cells["A1"].Value=10;w.SaveDocument(source,DocumentFormat.Xlsb);}File.SetLastWriteTimeUtc(source,DateTime.UtcNow.AddHours(-1));string sourceHash=Hash(source);
  Check((string)Call("SelectOpenPath",null,source)==source,"No recovery opens original without a prompt");
  string recovery=(string)Call("RecoveryPath",source);
  File.WriteAllText(recovery,"not a recovery workbook");
  Check(Call("NewerRecovery",source)==null,"Malformed recovery is not offered as verified");
  Reject(source,recovery,new FileInfo(recovery).Length,File.GetLastWriteTimeUtc(recovery),"Malformed recovery is not deleted");
  string missingSource=Path.Combine(output,"missing original.xlsb"),missingRecovery=(string)Call("RecoveryPath",missingSource);Recovery(missingRecovery,missingSource);
  Reject(missingSource,missingRecovery,new FileInfo(missingRecovery).Length,File.GetLastWriteTimeUtc(missingRecovery),"Missing original keeps its recovery copy safe");
  Recovery(recovery,source);string recoveryHash=Hash(recovery);
  Check(Choose(source,"Open original")==source&&Hash(recovery)==recoveryHash,"Open original retains recovery bytes");
  Check(Choose(source,"Open recovery")==recovery&&Hash(source)==sourceHash,"Open recovery retains original bytes");
  Check(Choose(source,"Cancel")==null&&Hash(recovery)==recoveryHash,"Cancel opens nothing and retains recovery");
  Check(Choose(source,"Delete recovery\u2026","No","Cancel")==null&&Hash(recovery)==recoveryHash,"Declining deletion retains recovery and returns to choices");
  var info=new FileInfo(recovery);
  Reject(source,source,new FileInfo(source).Length,File.GetLastWriteTimeUtc(source),"Original cannot be a deletion target");
  Reject(source,recovery,info.Length+1,info.LastWriteTimeUtc,"Changed recovery length rejected");
  Reject(source,recovery,info.Length,info.LastWriteTimeUtc.AddSeconds(-1),"Changed recovery timestamp rejected");
  using(var held=new FileStream(recovery,FileMode.Open,FileAccess.ReadWrite,FileShare.None))Reject(source,recovery,info.Length,info.LastWriteTimeUtc,"File open in another application rejected");
  string other=Path.Combine(output,"unrelated.xlsm");Recovery(other,source);Reject(source,other,new FileInfo(other).Length,File.GetLastWriteTimeUtc(other),"Unrelated filename cannot be deleted");Check(File.Exists(other),"Unrelated file retained");
  var modelType=app.GetType("Abovo.FileManager+ExcelModel");var ctor=modelType.GetConstructors().Single();dynamic model=ctor.Invoke(new object[]{0,Enum.ToObject(ctor.GetParameters()[1].ParameterType,0)});model.FileName=recovery;
  var models=Array.CreateInstance(modelType,1);models.SetValue(model,0);files.GetField("ExcelModels").SetValue(null,models);
  Reject(source,recovery,info.Length,info.LastWriteTimeUtc,"Recovery open in Summit rejected");files.GetField("ExcelModels").SetValue(null,null);model.ModelSpreadsheetControl.Dispose();
  Recovery(recovery,Path.Combine(output,"another source.xlsb"));info=new FileInfo(recovery);Reject(source,recovery,info.Length,info.LastWriteTimeUtc,"Mismatched provenance rejected");
  Recovery(recovery,source);
  Check(Choose(source,"Delete recovery\u2026","Yes")==source&&!File.Exists(recovery),"Confirmed native Delete removes only recovery and selects original");
  Check(Hash(source)==sourceHash&&Call("NewerRecovery",source)==null,"Original byte-for-byte unchanged; deleted recovery no longer offered");
  string legacy=Path.Combine(output,"private plan_recovery.xlsm");Recovery(legacy,source);Recovery(recovery,source);File.SetLastWriteTimeUtc(legacy,DateTime.UtcNow.AddMinutes(1));
  Check((string)Call("NewerRecovery",source)==legacy,"Legacy recovery remains discoverable");info=new FileInfo(legacy);Call("DeleteRecovery",source,legacy,info.Length,info.LastWriteTimeUtc);
  Check(!File.Exists(legacy)&&File.Exists(recovery)&&Hash(source)==sourceHash,"Deleting displayed legacy copy preserves other recovery and original");
  Console.WriteLine("PASS "+count+" assertions; only synthetic recovery copies were recycled");return 0;
 }catch(Exception e){Console.WriteLine(e);return 1;}}
}
