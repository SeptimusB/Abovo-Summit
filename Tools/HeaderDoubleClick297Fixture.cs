// Diagnostic only: native header routing, no workbook or user application input.
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.BandedGrid;

class HeaderDoubleClick297Fixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 sealed class QuietForm:Form{protected override bool ShowWithoutActivation{get{return true;}}protected override CreateParams CreateParams{get{var p=base.CreateParams;p.ExStyle|=0x08000000;return p;}}}
 [DllImport("user32.dll")]static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]static extern IntPtr SendMessage(IntPtr window,int message,IntPtr w,IntPtr l);
 static int checks;
 static object Read(object o,string n){for(var t=o.GetType();t!=null;t=t.BaseType){var p=t.GetProperty(n,F|BindingFlags.DeclaredOnly);if(p!=null)return p.GetValue(o,null);var f=t.GetField(n,F|BindingFlags.DeclaredOnly);if(f!=null)return f.GetValue(o);}throw new MissingMemberException(n);}
 static object Call(object o,string n,params object[] args){for(var t=o.GetType();t!=null;t=t.BaseType){var m=t.GetMethod(n,F|BindingFlags.DeclaredOnly);if(m!=null)return m.Invoke(o,args);}throw new MissingMemberException(n);}
 static void Check(bool okay,string message){if(!okay)throw new Exception(message);Console.WriteLine("PASS "+(++checks)+" "+message);}
 static void Pump(){Application.DoEvents();Thread.Sleep(10);Application.DoEvents();}
 [STAThread]static int Main(string[] args){try{
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));var helperType=app.GetType("Abovo.ColumnInplaceEditorHelper");var foreground=GetForegroundWindow();
  Console.WriteLine("SYSTEM doubleClickTime="+SystemInformation.DoubleClickTime+" doubleClickSize="+SystemInformation.DoubleClickSize+" app="+app.GetName().Version);
  foreach(bool date in new[]{false,true})using(var form=new QuietForm())using(var grid=new GridControl())using(var owner=date?(BaseEdit)new DateEdit():new ComboBoxEdit()){
   form.Opacity=0;form.ShowInTaskbar=false;form.ClientSize=new Size(900,500);grid.Dock=DockStyle.Fill;form.Controls.Add(grid);
   var view=new BandedGridView(grid);grid.MainView=view;view.OptionsView.ColumnAutoWidth=false;view.ColumnPanelRowHeight=44;
   var band=new GridBand(){Caption="Period",Width=400};view.Bands.Add(band);var first=new BandedGridColumn(){Visible=true,Width=160,Caption="Description"};var target=new BandedGridColumn(){Visible=true,Width=240,Caption="Target"};view.Columns.Add(first);view.Columns.Add(target);band.Columns.Add(first);band.Columns.Add(target);first.VisibleIndex=0;target.VisibleIndex=1;
   if(!date)((ComboBoxEdit)owner).Properties.Items.AddRange(new[]{"Year 1","Year 2"});owner.EditValue=date?(object)new DateTime(2026,1,1):"Year 1";
   object committed=owner.EditValue;int copied=0,commits=0;object callbackSender=null;
   object helper=Activator.CreateInstance(helperType,new object[]{target,owner.Properties,new EventHandler((s,e)=>{commits++;committed=((BaseEdit)s).EditValue;}),new EventHandler((s,e)=>{copied++;callbackSender=s;})});helperType.GetProperty("EditValue").SetValue(helper,owner.EditValue,null);
   form.Show();Pump();grid.Refresh();var bounds=(Rectangle)Read(helper,"_LastPaintedEditorBounds");Check(!bounds.IsEmpty,"Painted "+(date?"date":"combo")+" header has exact click bounds");
   Console.WriteLine("DPI control="+grid.DeviceDpi+" bounds="+bounds);
   Call(helper,"view_MouseDown",view,new MouseEventArgs(MouseButtons.Left,2,bounds.Left+8,bounds.Top+8,0));Pump();Check(copied==1&&ReferenceEquals(callbackSender,owner.Properties),"painted header double-click queues one original-repository callback");
   var editor=(BaseEdit)Read(helper,"ActiveEditor");Call(editor,"OnDoubleClick",EventArgs.Empty);Call(helper,"RaiseEditorDoubleClick",editor);Pump();Check(copied==2,"active and painted duplicate routes deduplicate within one message turn");
   object pending=date?(object)new DateTime(2031,1,1):"Year 2";editor.EditValue=pending;int before=commits;
   Call(editor,"OnDoubleClick",EventArgs.Empty);Pump();Check(copied==3&&commits==before&&Equals(committed,owner.EditValue),"active double-click leaves new defining value uncommitted (existing behaviour)");
   Check(Equals(Read(helper,"EditValue"),owner.EditValue),"copy callback reads previous helper header value while editor holds pending input");
   // Exercise the vendor's first-click timestamp bridge directly. This is not
   // claimed to be a physical double-click test: the host stays non-activating.
   var checkDouble=typeof(BaseEdit).GetMethod("CheckDoubleClick",F);
   editor.SendMouse(new Point(8,8),MouseButtons.Left);
   object[] message={Message.Create(editor.Handle,0x0201,new IntPtr(1),new IntPtr(8|(8<<16)))};checkDouble.Invoke(editor,message);
   Check(((Message)message[0]).Msg==0x0203,"DevExpress first-click timestamp recognises a second down inside Windows interval");
   editor.SendMouse(new Point(8,8),MouseButtons.Left);Thread.Sleep(SystemInformation.DoubleClickTime+30);
   message=new object[]{Message.Create(editor.Handle,0x0201,new IntPtr(1),new IntPtr(8|(8<<16)))};checkDouble.Invoke(editor,message);
   Check(((Message)message[0]).Msg==0x0201,"second click outside Windows interval is not promoted to double-click");
   Call(editor,"OnMaskBox_DoubleClick",editor,EventArgs.Empty);Pump();Check(copied==4,"native text-child double-click forwarding reaches the copy callback");
   Check(GetForegroundWindow()==foreground,"non-activating diagnostic leaves user foreground unchanged");form.Close();
  }
  Console.WriteLine("PASS "+checks+" diagnostic assertions; no workbook loaded.");return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
}
