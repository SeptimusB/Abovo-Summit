using System;
using System.IO;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.Serialization;
using System.Windows.Forms;
using DevExpress.XtraVerticalGrid;

// Exercise production menu population before the popup is visible. No workbook
// or DIT initialization is needed: this method only uses the three menu fields
// and an empty view-state dictionary, all supplied explicitly below.
public static class FundingFilterMenuFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static int assertions;
 static void Check(bool value,string text) { if(!value)throw new Exception(text);Console.WriteLine("PASS "+(++assertions)+" "+text); }
 [STAThread] public static int Main(string[] args) { try {
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string path=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(path)?Assembly.LoadFrom(path):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  Type dit=app.GetType("DataInterfaceTemplate");object owner=FormatterServices.GetUninitializedObject(dit);
  var views=dit.GetField("FundingViews",F);views.SetValue(owner,Activator.CreateInstance(views.FieldType));
  using(var menu=new ContextMenuStrip())using(var copy=new ToolStripMenuItem("Copy with headings")) {
   dit.GetField("ClipboardContextMenu",F).SetValue(owner,menu);
   dit.GetField("ClipboardCopyWithHeadersMenuItem",F).SetValue(owner,copy);
   dynamic identity=Activator.CreateInstance(app.GetType("Abovo.FundingScheduleFacility"));
   identity.FunderName="Fixture funder";identity.FacilityName="Fixture facility";identity.LoanName="Fixture loan";
   MethodInfo update=dit.GetMethod("UpdateFundingFilterMenu",F);update.Invoke(owner,new object[]{identity});
   var funder=(ToolStripMenuItem)dit.GetField("FundingFunderMenu",F).GetValue(owner);
   var facility=(ToolStripMenuItem)dit.GetField("FundingFacilityMenu",F).GetValue(owner);
   Check(!menu.Visible&&!funder.Visible,"Closed popup suppresses effective Visible despite configured menu availability");
   Check(funder.Available&&funder.Enabled,"Funder filter is available before opening the popup");
   if(args.Length>4&&args[4]=="reproduce") { Check(!facility.Available,"Original defect reproduced: reading hidden owner's Visible hides Facility filter");return 0; }
   Check(facility.Available&&facility.Enabled,"Facility filter is available before opening the popup");
   Check(funder.Text=="Filter to funder Fixture funder"&&facility.Text=="Filter to facility Fixture facility","Distinct funder/facility identities populate their own labels");
   int count=menu.Items.Count;update.Invoke(owner,new object[]{identity});
   Check(menu.Items.Count==count&&facility.Available,"Repeated popup preparation retains one visible facility action");
   identity.FacilityName="";update.Invoke(owner,new object[]{identity});
   Check(facility.Available&&!facility.Enabled&&funder.Enabled,"Empty facility leaves a disabled action while funder stays available");
   update.Invoke(owner,new object[]{null});Check(!funder.Available&&!facility.Available,"Non-funding target hides both filters");
   identity.HeaderName="Rep_Fund_29";identity.FacilityName="Fixture facility";
   // IsInvestment is derived from the named header used by the production model.
   var investment=identity.GetType().GetProperty("IsInvestment");
   if((bool)investment.GetValue(identity,null)){update.Invoke(owner,new object[]{identity});Check(!funder.Available&&!facility.Available,"Investment target hides both loan filters");}
   TestBoundary(app,dit);
   Console.WriteLine("PASS "+assertions+" assertions; no workbook opened or changed.");return 0;
  }
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;} }
 static void TestBoundary(Assembly app,Type dit) {
  using(var host=new Form())using(var grid=new VGridControl()) {
   host.Opacity=0;host.ShowInTaskbar=false;host.ClientSize=new Size(600,220);
   grid.Dock=DockStyle.Fill;grid.RowHeaderWidth=100;grid.RecordWidth=150;grid.RecordHeaderHeight=36;grid.OptionsView.ShowRecordHeaders=true;
   grid.LayoutStyle=LayoutViewStyle.MultiRecordView;
   var table=new DataTable();table.Columns.Add("Amount");for(int i=0;i<12;i++)table.Rows.Add("Loan "+i);grid.DataSource=table;
   var row=new DevExpress.XtraVerticalGrid.Rows.EditorRow();row.Properties.FieldName="Amount";row.Properties.Caption="Schedule date";row.Height=48;grid.Rows.Add(row);
   var item=new DevExpress.XtraEditors.Repository.RepositoryItemDateEdit();item.Appearance.BackColor=Color.LightGoldenrodYellow;item.Appearance.Options.UseBackColor=true;
   var helperType=app.GetType("Abovo.VGridRowInplaceEditorHelper");
   object helper=Activator.CreateInstance(helperType,new object[]{grid,row,item,new EventHandler((s,e)=>{})});
   helperType.GetProperty("EditValue").SetValue(helper,new DateTime(2026,9,24),null);
   MethodInfo drawRow=dit.GetMethod("DrawFundingRowHeaderBoundary",F);Rectangle rowBounds=Rectangle.Empty;bool paintBoundary=true;
   // Production order: the existing editor painter runs first, then the boundary.
   grid.CustomDrawRowHeaderCell+=(s,e)=>{if(paintBoundary)drawRow.Invoke(null,new object[]{s,e});if(e.Row==row)rowBounds=e.Bounds;};
   MethodInfo draw=dit.GetMethod("DrawFundingRecordBoundary",F);var bounds=new List<Rectangle>();
   grid.CustomDrawRecordHeader+=(s,e)=>{draw.Invoke(null,new object[]{s,e});bounds.Add(e.Bounds);};
   host.Controls.Add(grid);host.Show();Application.DoEvents();
   var presentation=app.GetType("Abovo.GridPresentation");presentation.GetMethod("Configure",F).Invoke(null,new object[]{grid,"Regression/FundingBoundary/"+Guid.NewGuid()});
   foreach(int zoom in new[]{100,60,200,100}) {
    presentation.GetMethod("SetZoom",F).Invoke(null,new object[]{grid,zoom,false});Application.DoEvents();int fixedEdge=-1;
    foreach(int firstRecord in new[]{0,4}) {
     grid.LeftVisibleRecord=firstRecord;Application.DoEvents();bounds.Clear();
     using(var baseline=new Bitmap(grid.Width,grid.Height))using(var bitmap=new Bitmap(grid.Width,grid.Height)) {
     paintBoundary=false;grid.DrawToBitmap(baseline,new Rectangle(Point.Empty,baseline.Size));paintBoundary=true;bounds.Clear();
     grid.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));Check(bounds.Count>0,"Pinned loan header paints at "+zoom+"%");
     var rect=bounds[0];int x=rect.Left+rect.Width/2;
     Check(bitmap.GetPixel(x,rect.Bottom-1).ToArgb()==Color.SteelBlue.ToArgb()&&bitmap.GetPixel(x,rect.Bottom-2).ToArgb()==Color.SteelBlue.ToArgb()&&bitmap.GetPixel(x,rect.Bottom-3).ToArgb()==Color.SteelBlue.ToArgb()&&bitmap.GetPixel(x,rect.Bottom-4).ToArgb()!=Color.SteelBlue.ToArgb(),"Pinned loan header boundary is exactly three pixels at "+zoom+"%");
     Check(!rowBounds.IsEmpty&&grid.LeftVisibleRecord==firstRecord,"Date header and requested horizontal position rendered at "+zoom+"% / "+firstRecord);
     int y=rowBounds.Top+rowBounds.Height/2;
     Check(bitmap.GetPixel(rowBounds.Right-1,y).ToArgb()==Color.SteelBlue.ToArgb()&&bitmap.GetPixel(rowBounds.Right-2,y).ToArgb()==Color.SteelBlue.ToArgb()&&bitmap.GetPixel(rowBounds.Right-3,y).ToArgb()==Color.SteelBlue.ToArgb()&&bitmap.GetPixel(rowBounds.Right-4,y).ToArgb()!=Color.SteelBlue.ToArgb(),"Fixed date/title boundary is exactly three pixels at "+zoom+"% / "+firstRecord);
     if(firstRecord==0)fixedEdge=rowBounds.Right;else Check(rowBounds.Right==fixedEdge,"Horizontal scrolling does not move the fixed header boundary at "+zoom+"%");
     var editorBounds=(Rectangle)helperType.GetMethod("GetEditorBounds",F).Invoke(helper,new object[]{rowBounds});
     Check(editorBounds.Right<=rowBounds.Right-3,"Native date editor remains inside the separator at "+zoom+"% / "+firstRecord);
     bool unchanged=true;for(int px=editorBounds.Left;px<editorBounds.Right;px++)for(int py=editorBounds.Top;py<editorBounds.Bottom;py++)if(bitmap.GetPixel(px,py)!=baseline.GetPixel(px,py))unchanged=false;
     Check(unchanged,"Existing in-place date editor pixels are preserved at "+zoom+"% / "+firstRecord);
     helperType.GetMethod("ShowEditor",F).Invoke(helper,new object[]{editorBounds,true});Application.DoEvents();
     var active=(Control)helperType.GetField("_ActiveEditor",F).GetValue(helper);
     var activeRight=grid.PointToClient(active.Parent.PointToScreen(new Point(active.Right,active.Top))).X;
     Check(activeRight<=rowBounds.Right-3,"Active date editor does not cover the three-pixel separator at "+zoom+"% / "+firstRecord);
     helperType.GetMethod("CloseEditor",F).Invoke(helper,null);
    }
    }
   }
   helperType.GetMethod("DetachForDisposal",F).Invoke(helper,null);item.Dispose();
   host.Close();
  }
 }
}
