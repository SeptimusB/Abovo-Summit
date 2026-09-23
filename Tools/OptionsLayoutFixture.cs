using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

public static class OptionsLayoutFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static object Field(object o,string n){return o.GetType().GetField(n,F).GetValue(o);}
 static void Check(bool ok,string message){if(!ok)throw new Exception(message);Console.WriteLine("PASS: "+message);}
 static void Drain(){for(int i=0;i<8;i++){Application.DoEvents();System.Threading.Thread.Sleep(15);}}
 [STAThread] public static int Main(string[] args){try{
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  foreach(int percent in new[]{75,97,100,150,200,98}){
   app.GetType("Abovo.PresentationScaleManager").GetMethod("SetInterfaceScale").Invoke(null,new object[]{percent,false});
   if(percent==98)AppDomain.CurrentDomain.GetAssemblies().Select(a=>a.GetType("DevExpress.XtraEditors.WindowsFormsSettings")).First(t=>t!=null).GetProperty("DefaultFont").SetValue(null,new Font("Segoe UI",18f),null);
   using(var options=(Form)Activator.CreateInstance(app.GetType("Abovo.ApplicationOptionsForm"))){
    if(percent==98)options.Font=new Font(options.Font.FontFamily,18f);
    options.Opacity=0;options.ShowInTaskbar=false;options.Show();Drain();
    dynamic track=Field(options,"ScaleTrack");track.EditValue=percent==98?97:percent;Drain();
    var title=(Control)Field(options,"PreviewTitle");var text=(Control)Field(options,"PreviewText");var preview=title.Parent;
    var reset=(Control)Field(options,"ResetButton");dynamic tabs=options.Controls[0].Controls[0];
    Check(((Control)tabs.TabPages[0]).Contains(reset),"Reset belongs to Display at "+percent+"%");
    Check(preview.Height>title.Height+text.Height&&title.Bounds.Bottom<=text.Top,"Preview reserves content height at "+percent+"%");
    Check(text.Bottom<=preview.ClientSize.Height-preview.Padding.Bottom,"Preview body not clipped at "+percent+"%");
    Console.WriteLine("LAYOUT scale="+percent+" dpi="+options.DeviceDpi+" form="+options.Size+" preview="+preview.Bounds+" title="+title.Bounds+" body="+text.Bounds);
    using(var bitmap=new Bitmap(options.Width,options.Height)){options.DrawToBitmap(bitmap,new Rectangle(Point.Empty,options.Size));bitmap.Save(Path.Combine(args[1],"display-"+percent+".png"));}
    tabs.SelectedTabPageIndex=1;Drain();Check(!reset.Visible,"Display reset absent from Recovery tab");
    dynamic enable=Field(options,"BackupEnabled");enable.Checked=true;
    var idle=(Control)Field(options,"BackupWhenIdle");var always=(Control)Field(options,"BackupAlways");
    var idleMinutes=(Control)Field(options,"BackupIdleMinutes");var minutes=(Control)Field(options,"BackupMinutes");
    foreach(var pair in new[]{new[]{idle,idleMinutes},new[]{always,minutes}}){
     dynamic checkbox=pair[0];Check(pair[0].Width>=checkbox.CalcBestSize().Width,"Timing caption fully fits at "+percent+"%");
     Check(pair[0].Right<=pair[1].Left&&pair[1].Right<=pair[1].Parent.ClientSize.Width,"Timing checkbox and editor do not overlap at "+percent+"%");
     Check(pair[1].Bottom<=pair[1].Parent.ClientSize.Height,"Timing editor height is not clipped at "+percent+"%");
    }
    Check(idle.Parent.Bottom<=always.Parent.Top,"Idle and maximum timing rows do not overlap");
    using(var bitmap=new Bitmap(options.Width,options.Height)){options.DrawToBitmap(bitmap,new Rectangle(Point.Empty,options.Size));bitmap.Save(Path.Combine(args[1],"recovery-"+percent+".png"));}
    tabs.SelectedTabPageIndex=2;Drain();Check(!reset.Visible,"Display reset absent from Integrity tab");
    tabs.SelectedTabPageIndex=0;options.Size=new Size(620,410);Drain();
    Check(text.Bottom<=preview.ClientSize.Height-preview.Padding.Bottom,"Short window scrolls instead of crushing preview");
    dynamic button=reset;button.PerformClick();Check(Convert.ToInt32(track.EditValue)==100,"Reset restores 100% preview");
    options.Close();
   }
  }
  // Deliberately lose the presentation cache; exports must still reflect the
  // authoritative log, as when a UI event subscriber missed a notification.
  var manager=app.GetType("Abovo.SystemMessageManager");object messages=manager.GetMethod("Acquire").Invoke(null,new object[]{987});
  manager.GetMethod("Publish").Invoke(null,new object[]{987,"Export regression sentinel",Enum.ToObject(app.GetType("Abovo.SystemMessageSeverity"),0),"Fixture","private"});
  ((System.Collections.IList)Field(messages,"Items")).Clear();
  Check(manager.GetMethod("CreateTextExport").Invoke(messages,null).ToString().Contains("Export regression sentinel"),"Text export survives missed UI notifications");
  Check(manager.GetMethod("CreateHtmlExport").Invoke(messages,null).ToString().Contains("Export regression sentinel"),"HTML export survives missed UI notifications");
  ((IDisposable)messages).Dispose();return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
}
