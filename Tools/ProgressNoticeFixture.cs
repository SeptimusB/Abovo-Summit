using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using DevExpress.XtraWaitForm;

public static class ProgressNoticeFixture {
    const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
    static void Check(bool ok,string message){if(!ok)throw new Exception(message);Console.WriteLine("PASS: "+message);}
    static void Pump(int ms){var clock=System.Diagnostics.Stopwatch.StartNew();while(clock.ElapsedMilliseconds<ms){Application.DoEvents();Thread.Sleep(10);}}
    [STAThread] public static int Main(string[] args){try{
        AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
        Application.EnableVisualStyles();Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
        var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
        var type=app.GetType("WaitFormA");
        foreach(float scale in new[]{1f,1.5f,2f})using(var form=(WaitForm)Activator.CreateInstance(type)){
            form.ShowInTaskbar=false;form.Opacity=0;form.Show();Pump(30);
            var panel=(ProgressPanel)type.GetProperty("progressPanel1",F).GetValue(form);
            using(var captionFont=new Font(panel.AppearanceCaption.Font.FontFamily,12f*scale))
            using(var descriptionFont=new Font(panel.AppearanceDescription.Font.FontFamily,8.25f*scale)){
                panel.AppearanceCaption.Font=captionFont;panel.AppearanceDescription.Font=descriptionFont;
                string longText="AGL - Business plan 2026-27 with Funding and Development changes.xlsb\r\nSaving committed inputs to the recovery XLSM. The original business plan and your unsaved changes remain unchanged; please wait while the recovery copy is safely written and checked.";
                foreach(string text in new[]{"Please wait...",longText,"C:\\A very long folder\\"+new string('X',200)+"_recovery.xlsm",string.Join("\r\n",Enumerable.Repeat(longText,20)),"Recovery copy saved. Normal Save still writes your XLSB plan."}){
                    form.SetCaption(text.StartsWith("Recovery copy saved")?"Complete":"Saving recovery copy");
                    form.SetDescription(text);Pump(25);
                    var work=Screen.FromControl(form).WorkingArea;
                    Check(work.Contains(form.Bounds),"Notice stays inside work area, font scale "+scale);
                    var vi=panel.ViewInfo;
                    Size caption=(Size)vi.GetType().GetMethod("GetCaptionSize",F).Invoke(vi,null);
                    Size description=(Size)vi.GetType().GetMethod("GetDescriptionSize",F).Invoke(vi,null);
                    Console.WriteLine("GEOMETRY: form="+form.ClientSize+" panel="+panel.Bounds+" caption="+caption+" description="+description+" scroll="+((ScrollableControl)panel.Parent).AutoScrollPosition);
                    Check(panel.Height>=caption.Height+description.Height,"Native wrapped text fits panel, font scale "+scale);
                    Check(panel.Right<=form.ClientSize.Width,"Panel remains inside available width");
                    if(text==longText){
                        Check(description.Height>descriptionFont.Height*2,"Multiline recovery text wraps");
                        using(var bitmap=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(bitmap,new Rectangle(Point.Empty,form.Size));bitmap.Save(Path.Combine(args[1],"progress-"+scale.ToString(System.Globalization.CultureInfo.InvariantCulture)+".png"));}
                    }
                    if(text.Length>5000)Check(((ScrollableControl)panel.Parent).VerticalScroll.Visible,"Oversized messages scroll instead of growing off screen");
                    if(text.StartsWith("Recovery copy saved"))Check(!((ScrollableControl)panel.Parent).VerticalScroll.Visible,"Completion shrinks and removes overflow scrolling");
                }
            }
            form.Close();
        }
        using(var owner=new Form()){owner.ShowInTaskbar=false;owner.Opacity=0;owner.Show();Pump(20);
            var wrapper=app.GetType("Abovo.FormSplashScreen");
            using(var notice=(IDisposable)Activator.CreateInstance(wrapper,new object[]{owner,"Saving recovery copy","Writing the private XLSM recovery copy; original unchanged..."})){
                Check((bool)wrapper.GetProperty("IsShowing",F).GetValue(notice),"Shared manager shows notice before work");
                Pump(150);
                wrapper.GetMethod("Complete").Invoke(notice,new object[]{"Recovery copy saved."});
                Check((bool)wrapper.GetProperty("OperationInProgress",F).GetValue(null),"Operation remains guarded until disposal");
            }
            Pump(1200);Check(!(bool)wrapper.GetProperty("OperationInProgress",F).GetValue(null),"Completion closes without leaking operation guard");owner.Close();
        }
        return 0;
    }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
