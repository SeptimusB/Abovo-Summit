// Isolated native-control regression: no workbook, user file or application session.
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraVerticalGrid;
using DevExpress.XtraVerticalGrid.Rows;

class InplaceZoomFixture {
 sealed class NonActivatingForm:Form{protected override bool ShowWithoutActivation{get{return true;}}protected override CreateParams CreateParams{get{var cp=base.CreateParams;cp.ExStyle|=0x08000000;return cp;}}}
 const BindingFlags Flags=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static Assembly app; static Type presentation; static int assertions; static string evidence; static readonly List<BaseEdit> linkedOwners=new List<BaseEdit>();
 static void Check(bool okay,string message){if(!okay)throw new Exception(message);Console.WriteLine("PASS "+(++assertions)+" "+message);}
 static object Call(object target,string name,params object[] args){return target.GetType().GetMethod(name,Flags).Invoke(target,args);}
 static object Read(object target,string name){var p=target.GetType().GetProperty(name,Flags);return p!=null?p.GetValue(target,null):target.GetType().GetField(name,Flags).GetValue(target);}
 static void Zoom(Control grid,int value){presentation.GetMethod("SetZoom",Flags).Invoke(null,new object[]{grid,value,false});Application.DoEvents();}
 static RepositoryItemComboBox Repository(){var owner=new ComboBoxEdit();linkedOwners.Add(owner);var repo=owner.Properties;repo.Appearance.Font=new Font("Segoe UI",9);repo.Appearance.Options.UseFont=true;repo.AppearanceFocused.Assign(repo.Appearance);repo.AppearanceReadOnly.Assign(repo.Appearance);repo.AppearanceDisabled.Assign(repo.Appearance);repo.Items.AddRange(new[]{"Original","Pending"});owner.EditValue="Original";return repo;}
 static int Test(bool vertical,int number){
  int layouts=0,commits=0;var helpers=new List<object>();var repos=new List<RepositoryItemComboBox>();
  using(var host=new NonActivatingForm()) {
   host.ShowInTaskbar=false;host.Opacity=0;host.ClientSize=new Size(1100,650);
   Control grid;BandedGridView banded=null;VGridControl vgrid=null;
   var table=new DataTable();for(int i=0;i<number;i++)table.Columns.Add("C"+i);table.Rows.Add(Enumerable.Repeat<object>("Original",number).ToArray());
   if(vertical){
    vgrid=new VGridControl();grid=vgrid;vgrid.RowHeaderWidth=220;vgrid.RecordWidth=160;vgrid.OptionsView.AutoScaleBands=false;vgrid.Appearance.RecordValue.Font=new Font("Segoe UI",9);vgrid.Appearance.RecordValue.Options.UseFont=true;
    vgrid.DataSource=table;for(int i=0;i<number;i++){var row=new EditorRow();row.Properties.FieldName="C"+i;row.Properties.Caption="Schedule "+i;vgrid.Rows.Add(row);}
    foreach(EditorRow row in vgrid.Rows){row.Height=30;var repo=Repository();repos.Add(repo);var helper=Activator.CreateInstance(app.GetType("Abovo.VGridRowInplaceEditorHelper"),new object[]{vgrid,row,repo,new EventHandler((s,e)=>commits++)});helper.GetType().GetProperty("EditValue").SetValue(helper,"Original",null);helpers.Add(helper);}
    vgrid.Layout+=(s,e)=>layouts++;
   }else{
    var standard=new GridControl();grid=standard;banded=new BandedGridView();banded.OptionsView.ColumnAutoWidth=false;banded.ColumnPanelRowHeight=44;banded.Appearance.Row.Font=new Font("Segoe UI",9);banded.Appearance.Row.Options.UseFont=true;standard.DataSource=table;
    banded.Columns.Clear();banded.Bands.Clear();var band=new GridBand(){Caption="Repeating inputs"};banded.Bands.Add(band);
    var label=new BandedGridColumn(){Caption="Description",Visible=true,Width=160};banded.Columns.Add(label);band.Columns.Add(label);label.VisibleIndex=0;
    band.Width=160;
    for(int i=0;i<number;i++){var periodBand=new GridBand(){Caption="Period "+i,Width=160};banded.Bands.Add(periodBand);var column=new BandedGridColumn(){FieldName="C"+i,Caption="C"+i,Visible=true,Width=160};banded.Columns.Add(column);periodBand.Columns.Add(column);column.VisibleIndex=i+1;var repo=Repository();repos.Add(repo);var helper=Activator.CreateInstance(app.GetType("Abovo.ColumnInplaceEditorHelper"),new object[]{column,repo,new EventHandler((s,e)=>commits++),null});helper.GetType().GetProperty("EditValue").SetValue(helper,"Original",null);helpers.Add(helper);}
    Check(banded.GridControl==null&&helpers.Count==number,"lazy tab builds all column helpers before the view has a GridControl ("+number+")");
    standard.ViewCollection.Add(banded);standard.MainView=banded;banded.GridControl=standard;
    banded.Layout+=(s,e)=>layouts++;
   }
   grid.Dock=DockStyle.Fill;host.Controls.Add(grid);presentation.GetMethod("Configure",Flags).Invoke(null,new object[]{grid,"Regression/InplaceZoom/"+Guid.NewGuid()});host.Show();Application.DoEvents();
   if(!vertical)banded.LeftCoord=0;grid.Refresh();
   if(!vertical)Console.WriteLine("SETUP columns="+banded.VisibleColumns.Count+" left="+banded.LeftCoord+" band="+banded.Bands[0].Width+" control="+grid.Bounds+" cached="+helpers.Count(h=>!((Rectangle)Read(h,"_LastPaintedEditorBounds")).IsEmpty));
   var first=vertical?helpers[0]:helpers.Where(h=>!((Rectangle)Read(h,"_LastPaintedEditorBounds")).IsEmpty).OrderBy(h=>((Rectangle)Read(h,"_LastPaintedEditorBounds")).Left).First();Call(first,"ShowEditorFromKeyboard");Application.DoEvents();
   var editor=(BaseEdit)Read(first,vertical?"_ActiveEditor":"ActiveEditor");Check(editor!=null,"active "+(vertical?"vertical":"column")+" header created ("+number+")");editor.EditValue="Pending";
   float original=vertical?vgrid.Appearance.RecordValue.GetFont().SizeInPoints:banded.Appearance.Row.GetFont().SizeInPoints;layouts=0;var timer=System.Diagnostics.Stopwatch.StartNew();Zoom(grid,130);timer.Stop();int zoomLayouts=layouts;
   Check(commits==0&&ReferenceEquals(editor,Read(first,vertical?"_ActiveEditor":"ActiveEditor"))&&!editor.IsDisposed&&(string)editor.EditValue=="Pending","zoom keeps uncommitted header value and editor instance ("+number+")");
   Check(repos[0].Appearance.GetFont().SizeInPoints<=original*1.3f+0.05f,"header font follows owner without enlargement ("+number+")");
   foreach(int percent in new[]{50,200,100,50,200,100,130,100}){
    Zoom(grid,percent);
    int actualPercent=(int)presentation.GetMethod("ZoomPercent",Flags).Invoke(null,new object[]{grid});
    if(percent==50)Check(actualPercent==60,"legacy 50% request stops at the readable 60% zoom minimum");
    float effectiveFont=vertical?vgrid.Appearance.RecordValue.GetFont().SizeInPoints:banded.Appearance.Row.GetFont().SizeInPoints;
    foreach(var repo in repos){if(repo.Appearance.GetFont().SizeInPoints>effectiveFont+0.05f)throw new Exception("Header exceeds the effective owner size");foreach(var appearance in new[]{repo.AppearanceFocused,repo.AppearanceReadOnly,repo.AppearanceDisabled})CheckSize(appearance.GetFont().SizeInPoints,repo.Appearance.GetFont().SizeInPoints,percent);}
    if(commits!=0||editor.IsDisposed)Console.WriteLine("EDITOR LOST percent="+percent+" commits="+commits+" disposed="+editor.IsDisposed+" layouts="+layouts+" left="+(banded==null?0:banded.LeftCoord));
    Check(commits==0&&!editor.IsDisposed&&(string)editor.EditValue=="Pending","zoom "+percent+"% preserves pending input");
    Rectangle bounds;
    if(vertical){var header=(Rectangle)Read(first,"_LastHeaderBounds");bounds=(Rectangle)Call(first,"GetEditorBounds",header);Check(bounds.Height<=header.Height&&header.Contains(bounds),"vertical editor remains inside its row header at "+percent+"%");}
    else {bounds=(Rectangle)Read(first,"_LastPaintedEditorBounds");Check(bounds.Height>0&&bounds.Height<=banded.ColumnPanelRowHeight,"column editor height is bounded at "+percent+"%");}
    Check(editor.Height==bounds.Height,"live editor matches the bounded painted height at "+percent+"%");
    using(var graphics=grid.CreateGraphics()){
     CheckGlyphFits(editor.Properties,graphics,bounds.Size,"live "+(vertical?"vertical":"column")+" glyphs at "+percent+"%");
     CheckValueFitsWidth(editor.Properties,graphics,bounds.Size,editor.EditValue,"pending header text width at "+actualPercent+"%");
    }
    Check(editor.Properties.Appearance.GetFont().SizeInPoints>=8.0f-0.05f,"live header remains at least 8pt at "+actualPercent+"%");
   }
   Check(editor.Properties.AppearanceFocused.GetFont().SizeInPoints<=original+0.05f,"focused font returns without accumulated enlargement");
   Check((bool)Call(first,"CommitForSave")&&commits==1,"pending header value commits exactly once when explicitly requested");
   if(number==4){
    var repo=(RepositoryItem)Read(first,"_Item");repo.Appearance.ForeColor=Color.Magenta;repo.Appearance.Options.UseForeColor=true;Rectangle smallInk=Rectangle.Empty,largeInk=Rectangle.Empty;
    foreach(int percent in new[]{60,200}){
     Zoom(grid,percent);grid.Refresh();Rectangle bounds=vertical?(Rectangle)Call(first,"GetEditorBounds",(Rectangle)Read(first,"_LastHeaderBounds")):(Rectangle)Read(first,"_LastPaintedEditorBounds");
     using(var bitmap=new Bitmap(grid.Width,grid.Height)){
      grid.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));var ink=MagentaBounds(bitmap,Rectangle.Intersect(bounds,new Rectangle(Point.Empty,bitmap.Size)));
      if(evidence!=null)bitmap.Save(Path.Combine(evidence,(vertical?"vertical":"column")+"-inactive-"+percent+".png"));
      Console.WriteLine("PAINT linked="+(repo.OwnerEdit!=null)+" font="+repo.Appearance.GetFont().SizeInPoints+" colour="+repo.Appearance.ForeColor+" use="+repo.Appearance.Options.UseForeColor+" bounds="+bounds+" ink="+ink);
      Check(!ink.IsEmpty,"native inactive "+(vertical?"vertical":"column")+" header uses repository text appearance at "+percent+"%");
      if(percent==60)smallInk=ink;else largeInk=ink;
      if(evidence!=null)bitmap.Save(Path.Combine(evidence,(vertical?"vertical":"column")+"-inactive-"+percent+".png"));
     }
    }
    Check(largeInk.Height>smallInk.Height*1.5&&largeInk.Width>smallInk.Width*1.5,"native inactive header glyphs enlarge with zoom instead of remaining default8.25pt");
   }
   Console.WriteLine("MEASURE "+(vertical?"vertical":"column")+" helpers="+number+" layouts="+zoomLayouts+" zoomMs="+timer.ElapsedMilliseconds);
   foreach(var helper in helpers)if(vertical)Call(helper,"DetachForDisposal");host.Close();return zoomLayouts;
  }
 }
 static void CheckGlyphFits(RepositoryItem repo,Graphics graphics,Size size,string message){
  var paint=new DevExpress.Utils.AppearanceObject();paint.Assign(repo.Appearance);var info=repo.CreateViewInfo();info.PaintAppearance=paint;info.Bounds=new Rectangle(Point.Empty,size);info.EditValue=new DateTime(2051,9,30);info.CalcViewInfo(graphics);int available=info.GetTextBounds().Height;
  var font=repo.Appearance.GetFont();float actual=Math.Max(font.GetHeight(graphics),TextRenderer.MeasureText(graphics,"Ag",font,Size.Empty,TextFormatFlags.NoPadding|TextFormatFlags.SingleLine).Height);
  Check(actual<=available+1.0f,message+" text="+actual+" available="+available+" dpi="+graphics.DpiY);
 }
 static void CheckValueFitsWidth(RepositoryItem repo,Graphics graphics,Size size,object value,string message){
  var paint=new DevExpress.Utils.AppearanceObject();paint.Assign(repo.Appearance);var info=repo.CreateViewInfo();info.PaintAppearance=paint;info.Bounds=new Rectangle(Point.Empty,size);info.EditValue=value;info.CalcViewInfo(graphics);
  int actual=TextRenderer.MeasureText(graphics,repo.GetDisplayText(value),repo.Appearance.GetFont(),Size.Empty,TextFormatFlags.NoPadding|TextFormatFlags.SingleLine).Width;
  Check(actual<=info.GetTextBounds().Width+1,message+" text="+actual+" available="+info.GetTextBounds().Width);
 }
 static void LinkedDatePainting(){
  using(var owner=new DateEdit())using(var bitmap=new Bitmap(600,100))using(var graphics=Graphics.FromImage(bitmap)){
   var repo=owner.Properties;repo.DisplayFormat.FormatType=DevExpress.Utils.FormatType.DateTime;repo.DisplayFormat.FormatString="dd-MMM-yyyy";repo.Appearance.ForeColor=Color.Magenta;repo.Appearance.Options.UseForeColor=true;repo.Appearance.BackColor=Color.White;repo.Appearance.Options.UseBackColor=true;
   Rectangle small=Rectangle.Empty,large=Rectangle.Empty;
   foreach(float size in new[]{8f,9.386666f,18f,24f}){
    repo.Appearance.Font=new Font("Segoe UI",size);repo.Appearance.Options.UseFont=true;var original=repo.Appearance.Font;graphics.Clear(Color.White);
    app.GetType("Abovo.DrawEditorHelper").GetMethod("DrawEdit",Flags).Invoke(null,new object[]{graphics,repo,new Rectangle(5,5,550,80),new DateTime(2042,7,24),false,true});
    Check(ReferenceEquals(original,repo.Appearance.Font)&&repo.Appearance.Options.UseFont&&Math.Abs(repo.Appearance.Font.SizeInPoints-size)<0.01f,"linked date native paint preserves live repository font and UseFont at "+size+"pt (owner "+owner.Font.SizeInPoints+"pt)");
    Check(repo.Appearance.Font.GetHeight(graphics)>0,"linked date font remains valid after native painting");
    var ink=MagentaBounds(bitmap,new Rectangle(Point.Empty,bitmap.Size));Check(!ink.IsEmpty,"linked date native painter uses repository appearance at "+size+"pt");if(size==8)small=ink;if(size==24)large=ink;
   }
   Check(large.Height>small.Height*2&&large.Width>small.Width*2,"linked date native glyphs enlarge instead of reverting to the hidden owner's default font");
  }
 }
 static void DensityAndScale(){
  var formatting=app.GetType("Abovo.InplaceEditorFormatting");
  using(var host=new NonActivatingForm())using(var grid=new VGridControl())using(var repo=new RepositoryItemDateEdit()){
   host.ShowInTaskbar=false;host.Opacity=0;host.ClientSize=new Size(800,500);grid.Dock=DockStyle.Fill;host.Controls.Add(grid);var row=new EditorRow();grid.Rows.Add(row);
   repo.Appearance.Font=new Font("Segoe UI",24);repo.Appearance.Options.UseFont=true;repo.Appearance.BackColor=Color.FromArgb(0,91,170);repo.Appearance.ForeColor=Color.White;repo.Appearance.Options.UseBackColor=true;repo.Appearance.Options.UseForeColor=true;
   repo.DisplayFormat.FormatType=DevExpress.Utils.FormatType.DateTime;repo.DisplayFormat.FormatString="dd-MMM-yyyy";repo.Mask.EditMask="dd/MM/yy";var mask=repo.Mask.EditMask;var colour=repo.Appearance.BackColor;
   presentation.GetMethod("Configure",Flags).Invoke(null,new object[]{grid,"Regression/DensityDpi/"+Guid.NewGuid()});host.CreateControl();grid.CreateControl();Application.DoEvents();
   // Native GDI view-info uses the real HDC, not bitmap-resolution metadata.
   // Exercise presentation scales on that same HDC; true DPI is reported by
   // the application-config Funding fixture, not simulated with SetResolution.
   foreach(float scale in new[]{1.0f,1.5f,2.5f}){
    using(var graphics=grid.CreateGraphics())foreach(int density in new[]{2,1,2,1}){
     using(var layout=(IDisposable)presentation.GetMethod("BeginBaseLayout",Flags).Invoke(null,new object[]{grid})){
      grid.Font=new Font("Segoe UI",9*density*scale);grid.Appearance.RecordValue.Font=grid.Font;grid.Appearance.RecordValue.Options.UseFont=true;row.Height=(int)Math.Ceiling(grid.Font.GetHeight(graphics))+12;
     }
     foreach(int percent in new[]{50,100,200}){
      Zoom(grid,percent);
      var existingFont=repo.Appearance.Font;float existingFontSize=existingFont.SizeInPoints;
      int nativeHeight=(int)formatting.GetMethod("DesiredEditorHeight",Flags).Invoke(null,new object[]{repo,grid,graphics});
      int nativeWidth=(int)formatting.GetMethod("DesiredEditorWidth",Flags).Invoke(null,new object[]{repo,grid,graphics,new DateTime(2051,9,30)});
      Check(Object.ReferenceEquals(existingFont,repo.Appearance.Font)&&Math.Abs(repo.Appearance.Font.SizeInPoints-existingFontSize)<0.01f&&repo.Appearance.Font.GetHeight(graphics)>0,"native measurement leaves live repository font intact");
      var bounds=new Rectangle(0,0,Math.Max(nativeWidth,(int)(300*density*scale*percent/100)),Math.Max(nativeHeight,row.Height-4));
      Console.WriteLine("DENSITY dpi="+graphics.DpiY+" scale="+scale+" density="+density+" percent="+percent+" owner="+grid.Appearance.RecordValue.GetFont().SizeInPoints+" repoBefore="+repo.Appearance.GetFont().SizeInPoints+" nativeHeight="+nativeHeight+" rowHeight="+row.Height+" bounds="+bounds);
      formatting.GetMethod("FitEditorToBounds",Flags).Invoke(null,new object[]{repo,null,grid,graphics,bounds});
      Console.WriteLine("DENSITY repoAfter="+repo.Appearance.GetFont().SizeInPoints);
      CheckGlyphFits(repo,graphics,bounds.Size,"density="+density+" zoom="+percent);
      Check(Math.Abs(repo.Appearance.GetFont().SizeInPoints-grid.Appearance.RecordValue.GetFont().SizeInPoints)<=0.05f,"native minimum geometry preserves the effective owner font at presentation scale "+scale);
     }
    }
   }
   Check(repo.Mask.EditMask==mask&&repo.Appearance.BackColor==colour,"Density/zoom fitting preserves date mask and fill");host.Close();
  }
 }
 static void CheckSize(float actual,float expected,int percent){if(Math.Abs(actual-expected)>0.05f)throw new Exception("font-state mismatch at "+percent+"%: "+actual+" expected "+expected);}
 static Rectangle MagentaBounds(Bitmap bitmap,Rectangle area){int left=int.MaxValue,top=int.MaxValue,right=-1,bottom=-1;for(int y=area.Top;y<area.Bottom;y++)for(int x=area.Left;x<area.Right;x++){var c=bitmap.GetPixel(x,y);if(c.R-c.G>50&&c.B-c.G>50){left=Math.Min(left,x);top=Math.Min(top,y);right=Math.Max(right,x);bottom=Math.Max(bottom,y);}}return right<0?Rectangle.Empty:Rectangle.FromLTRB(left,top,right+1,bottom+1);}
 [System.Runtime.InteropServices.DllImport("user32.dll")]static extern IntPtr GetForegroundWindow();
 [System.Runtime.InteropServices.DllImport("user32.dll")]static extern bool SetForegroundWindow(IntPtr window);
 [System.Runtime.InteropServices.DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
 [STAThread]static int Main(string[] args){var previousWindow=GetForegroundWindow();try{
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));presentation=app.GetType("Abovo.GridPresentation");
  evidence=args.Length>1?args[1]:null;
  if(args.Length>4&&args[4]=="density"){DensityAndScale();Console.WriteLine("PASS "+assertions+" density assertions");return 0;}
  if(args.Length>4&&args[4]=="linked"){LinkedDatePainting();Console.WriteLine("PASS "+assertions+" linked assertions");return 0;}
  int small=Test(false,4),large=Test(false,100);Check(large<=small+3,"100 column headers do not cause per-editor grid layouts");
  small=Test(true,4);large=Test(true,100);Check(large<=small+3,"100 vertical headers do not cause per-editor grid layouts");
  LinkedDatePainting();DensityAndScale();
  Console.WriteLine("PASS "+assertions+" assertions");return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}finally{foreach(var owner in linkedOwners)owner.Dispose();uint process;GetWindowThreadProcessId(GetForegroundWindow(),out process);if(previousWindow!=IntPtr.Zero&&process==(uint)System.Diagnostics.Process.GetCurrentProcess().Id)SetForegroundWindow(previousWindow);}}
}
