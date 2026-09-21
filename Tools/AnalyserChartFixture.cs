using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.XtraCharts;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraTab;
using DevExpress.Spreadsheet;
using DevExpress.Utils.Menu;
using DevExpress.XtraBars.Docking2010;

// Native controls on a private, unsaved workbook copy. No user windows touched.
public static class AnalyserChartFixture {
    const BindingFlags Flags=BindingFlags.Instance|BindingFlags.Public|BindingFlags.NonPublic;
    [STAThread] public static int Main(string[] args) {
        try { Run(args); return 0; } catch(Exception e) { Console.Error.WriteLine(e); return 1; }
    }
    static void Run(string[] args) {
        AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{
            string path=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");
            return File.Exists(path)?Assembly.LoadFrom(path):null;
        };
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
        Application.EnableVisualStyles();
        var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
        app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
        var files=app.GetType("Abovo.FileManager");
        files.GetMethod("Initialise").Invoke(null,new object[]{null});
        string scratch=Path.Combine(args[2],"Chart-fixture.xlsb");
        File.Copy(args[1],scratch);
        var open=files.GetMethod("OpenModel");
        dynamic result=open.Invoke(null,new object[]{scratch,new FileInfo(scratch),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
        Check(!result.BError,"Open private workbook copy");
        dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);
        using(var form=new Form{Opacity=0,ShowInTaskbar=false,ClientSize=new Size(1650,950)})
        using(var analyser=(Control)Activator.CreateInstance(app.GetType("BPIncomeExpenditureAnalyserV2"),new object[]{0,null})) {
            analyser.Dock=DockStyle.Fill;form.Controls.Add(analyser);form.Show();Application.DoEvents();
            var charts=(IList)Field(analyser,"ChartViews");
            var tabs=(XtraTabControl)Field(analyser,"XtraTabControlAnalyser");
            Check(charts.Count==3,"Chart offered on all three statements");
            var mainToggle=(WindowsUIButton)Field(analyser,"ChartViewButton");
            var mainButtons=(WindowsUIButtonPanel)Field(analyser,"WindowsUIButtonPanelAnalyser");
            Check(!mainToggle.UseCaption && mainToggle.ImageOptions.HasSvgImage,"Icon-only main chart button");
            Check(mainButtons.Buttons[mainButtons.Buttons.Count-2] is WindowsUISeparator,"Chart button has its own separator");
            object datasource=Field(analyser,"DSAnalDataRange");
            int groupsChecked=0,leavesChecked=0;
            for(int i=0;i<3;i++) {
                tabs.SelectedTabPageIndex=i; Application.DoEvents();
                object state=charts[i];
                var view=(GridView)Field(state,"View");
                var before=Expansion(view);
                Check(mainToggle.ToolTip=="Show line chart","Per-tab figures state reflected in main button");
                Call(analyser,"ToggleCurrentStatementChart"); Application.DoEvents();
                var chart=(ChartControl)Field(state,"Chart");
                Check(chart.Series.Count>0,"Overview series: "+tabs.SelectedTabPage.Text);
                var toolbar=(Control)Field(state,"Toolbar");
                Check(toolbar.Visible && toolbar.Height>0,"Chart toolbar visible and sized");
                Check(mainToggle.ToolTip=="Show figures grid" && mainToggle.ImageOptions.HasSvgImage,"Main toggle offers grid icon");
                Check(chart.Series.Cast<Series>().All(s=>((LineSeriesView)s.View).LineStyle.Thickness==5),"Easier-to-target series lines");
                Check(chart.Series.Cast<Series>().All(s=>((LineSeriesView)s.View).LineMarkerOptions.Size==8),"Larger series point markers");
                var diagram=(XYDiagram)chart.Diagram;
                Check(diagram.EnableAxisXZooming && diagram.EnableAxisYZooming && diagram.ZoomingOptions.UseMouseWheel,"Native wheel zoom enabled on both axes");
                chart.Focus();chart.Refresh();Application.DoEvents();
                double beforeX=diagram.AxisX.VisualRange.MaxValueInternal-diagram.AxisX.VisualRange.MinValueInternal;
                double beforeY=diagram.AxisY.VisualRange.MaxValueInternal-diagram.AxisY.VisualRange.MinValueInternal;
                var point=diagram.DiagramToPoint(chart.Series[0].Points[15].Argument,diagram.AxisY.VisualRange.MinValueInternal+beforeY/2);
                typeof(Control).GetMethod("OnMouseWheel",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(chart,new object[]{new MouseEventArgs(MouseButtons.None,0,point.Point.X,point.Point.Y,120)});
                Application.DoEvents();
                double afterX=diagram.AxisX.VisualRange.MaxValueInternal-diagram.AxisX.VisualRange.MinValueInternal;
                double afterY=diagram.AxisY.VisualRange.MaxValueInternal-diagram.AxisY.VisualRange.MinValueInternal;
                Console.WriteLine("WHEEL RANGES "+tabs.SelectedTabPage.Text+": X "+beforeX+"->"+afterX+" Y "+beforeY+"->"+afterY);
                Check(afterX<beforeX && (beforeY<=0 || afterY<beforeY),"Mouse wheel narrows axes with a nondegenerate range");
                ((SimpleButton)Field(state,"ResetZoomButton")).PerformClick();Application.DoEvents();
                Check(Math.Abs((diagram.AxisX.VisualRange.MaxValueInternal-diagram.AxisX.VisualRange.MinValueInternal)-beforeX)<0.001,"Reset zoom restores forecast range");
                Console.WriteLine("ROOTS: "+String.Join(" | ",chart.Series.Cast<Series>().Select(s=>s.Name)));
                // Every reachable grouping level uses exactly the existing group summaries.
                Walk(state,ref groupsChecked,ref leavesChecked,0);
                Check(before==Expansion(view),"Chart drill does not alter grid expansion");
                if(i==0) {
                    var rental=chart.Series.Cast<Series>().FirstOrDefault(s=>s.Name.Equals("Rental Income",StringComparison.OrdinalIgnoreCase));
                    Check(rental!=null,"Rental Income at overview");
                    Call(state,"DrillInto",rental.Tag);Application.DoEvents();
                    var names=chart.Series.Cast<Series>().Select(s=>s.Name).ToArray();
                    Check(names.Any(n=>n.Equals("Rents Receivable",StringComparison.OrdinalIgnoreCase)) && names.Any(n=>n.Equals("Service Income",StringComparison.OrdinalIgnoreCase)),"Rental drill shows required headings including zero rows");
                    string path=PathState(state);
                    Call(analyser,"ToggleCurrentStatementChart");
                    Check(!toolbar.Visible && view.GridControl.Visible,"Grid restores with chart toolbar hidden");
                    Call(analyser,"ToggleCurrentStatementChart");
                    Check(PathState(state)==path,"Figures/chart toggle retains branch");
                    Application.DoEvents();chart.Refresh();
                    if(Path.GetFileName(args[1]).StartsWith("Demo")) {
                        var rents=chart.Series.Cast<Series>().First(s=>s.Name.Equals("Rents Receivable",StringComparison.OrdinalIgnoreCase));
                        var rentPoint=rents.Points[15];
                        var marker=((XYDiagram)chart.Diagram).DiagramToPoint(rentPoint.Argument,rentPoint.Values[0]).Point;
                        var nearMarker=new Point(marker.X+2,marker.Y+2);
                        Check(Object.ReferenceEquals(chart.CalcHitInfo(nearMarker).Series,rents),"Larger marker is hit away from its exact centre");
                        Call(state,"ChartClicked",chart,new MouseEventArgs(MouseButtons.Left,1,nearMarker.X,nearMarker.Y,0));
                        Check(PathState(state)==path,"Left-click does not change the drill branch");
                    }
                    chart.ExportToImage(Path.Combine(args[2],"rental-chart.png"),System.Drawing.Imaging.ImageFormat.Png);
                    Call(state,"GoUp");
                    // Exercise the actual MouseClick hit-test route, including legend hits.
                    Application.DoEvents();chart.Refresh();
                    var target=chart.Series.Cast<Series>().First(s=>s.Name.Equals("Rental Income",StringComparison.OrdinalIgnoreCase));
                    Point? hitPoint=null;
                    for(int x=chart.Width-2;x>=0 && hitPoint==null;x-=4)
                        for(int y=2;y<chart.Height;y+=4) {
                            var hit=chart.CalcHitInfo(new Point(x,y));
                            if(hit.InLegend && Object.ReferenceEquals(hit.Series,target)) {hitPoint=new Point(x,y);break;}
                        }
                    Check(hitPoint.HasValue,"Legend hit testing identifies Rental Income");
                    var menu=(DXPopupMenu)Field(state,"NavigationMenu");
                    // ShowPopup runs its own native menu message loop until dismissed.
                    using(var closeMenu=new Timer{Interval=250}) {
                        closeMenu.Tick+=(s,e)=>{closeMenu.Stop();menu.HidePopup();};
                        closeMenu.Start();
                        Call(state,"ChartClicked",chart,new MouseEventArgs(MouseButtons.Right,1,hitPoint.Value.X,hitPoint.Value.Y,0));
                    }
                    Check(((IList)Field(state,"Navigation")).Count==0,"Right-click opens menu without immediately drilling");
                    Check(menu.Items.Count==2 && menu.Items[0].Caption=="Go Up" && !menu.Items[0].Enabled,"Overview Go Up disabled");
                    Check(menu.Items[1].Enabled && menu.Items[1].Caption=="Drill down to Rental Income","Menu identifies clicked series");
                    menu.HidePopup();menu.Items[1].GenerateClickEvent();Application.DoEvents();
                    Check(((IList)Field(state,"Navigation")).Count==1,"Context menu drills to children");
                    Call(state,"BuildNavigationMenu",new object[]{null});
                    Check(menu.Items[0].Enabled && !menu.Items[1].Enabled,"No-hit context allows Go Up but not drill");
                    menu.Items[0].GenerateClickEvent();Application.DoEvents();
                    Check(((IList)Field(state,"Navigation")).Count==0,"Context Go Up returns to overview");
                }
                ((CheckEdit)Field(state,"Totals")).Checked=true; Application.DoEvents(); Verify(state);
                Console.WriteLine("WITH TOTALS: "+String.Join(" | ",chart.Series.Cast<Series>().Select(s=>s.Name)));
                ((CheckEdit)Field(state,"Totals")).Checked=false;
                Check(Object.ReferenceEquals(datasource,Field(analyser,"DSAnalDataRange")),"Drill and view switch do not rebind/recalculate");
            }
            Console.WriteLine("PARITY: "+groupsChecked+" group series, "+leavesChecked+" leaf series");
            tabs.SelectedTabPageIndex=0;Application.DoEvents();
            object soci=charts[0];
            var sociChart=(ChartControl)Field(soci,"Chart");
            Call(soci,"DrillInto",sociChart.Series.Cast<Series>().First(s=>s.Name.Equals("Rental Income",StringComparison.OrdinalIgnoreCase)).Tag);
            string retained=PathState(soci);
            Call(analyser,"CreateTransactionalDBSnapshot",new object[]{null});Application.DoEvents();
            var modeType=analyser.GetType().GetNestedType("AnalyserDataSourceMode",BindingFlags.NonPublic);
            foreach(string mode in new[]{"Snapshot","Comparison","Live"}) {
                Call(analyser,"SwitchDataSource",Enum.Parse(modeType,mode));Application.DoEvents();
                foreach(object state in charts) {
                    tabs.SelectedTabPage=(XtraTabPage)Field(state,"Page");Application.DoEvents();Verify(state);
                    Check(((LabelControl)Field(state,"Trail")).Text.Contains(mode=="Comparison"?"Differences":mode),"Source clearly labelled: "+mode);
                    if(mode=="Comparison") {
                        var c=(ChartControl)Field(state,"Chart");
                        Check(c.Series.Cast<Series>().SelectMany(s=>s.Points.Cast<SeriesPoint>()).Where(p=>!p.IsEmpty).All(p=>Math.Abs(p.Values[0])<0.0001),"New snapshot differences are zero");
                    }
                }
                Check(PathState(soci)==retained,"Source switch retains chart path: "+mode);
            }
            tabs.SelectedTabPageIndex=0; Application.DoEvents();
            // An ordinary typed rent edit must leave Snapshot frozen and update Live/Differences.
            Call(analyser,"SwitchDataSource",Enum.Parse(modeType,"Snapshot"));Application.DoEvents();
            double[] frozen=Values(sociChart,"Rents Receivable");
            IWorkbook book=(IWorkbook)model.WB;
            var rent=book.DefinedNames.GetDefinedName("TransRents").Range[0,0];
            // Mirror a second open assumptions interface: its worksheet participates
            // in CalculateWSs alongside the analyser, exactly as a DIT edit does.
            int editRegistration=model.WBCalcEngine.AddActiveObject(new object());
            model.WBCalcEngine.AddActiveWorksheet(editRegistration,rent.Worksheet,false);
            dynamic change=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));
            change.ModelID=0;change.WSName=rent.Worksheet.Name;change.CellAddress=rent.GetReferenceA1();
            change.OriginalValue=rent.Value.NumericValue;change.ChangedValue=rent.Value.NumericValue*0.5+1;change.DataFormat="D2";
            change.Description="Unsaved chart regression rent edit";change.TimeStamp=DateTime.Now;
            dynamic edit=model.ChangeManager.ProcessChange(change);
            Check(edit.BSuccess,"Typed rent edit through ChangeManager");Application.DoEvents();
            Check(frozen.SequenceEqual(Values(sociChart,"Rents Receivable")),"Snapshot chart remains frozen after edit");
            Call(analyser,"SwitchDataSource",Enum.Parse(modeType,"Live"));Application.DoEvents();Verify(soci);
            double[] live=Values(sociChart,"Rents Receivable");
            if(Path.GetFileName(args[1]).StartsWith("Demo"))Check(!frozen.SequenceEqual(live),"Populated live chart changes after rent edit");
            Call(analyser,"SwitchDataSource",Enum.Parse(modeType,"Comparison"));Application.DoEvents();Verify(soci);
            double[] difference=Values(sociChart,"Rents Receivable");
            for(int y=0;y<live.Length;y++)Check(Math.Abs(difference[y]-(live[y]-frozen[y]))<0.001,"Difference equals live minus snapshot year "+(y+1));
            Call(analyser,"SwitchDataSource",Enum.Parse(modeType,"Live"));Application.DoEvents();
            model.WBCalcEngine.RemoveActiveObject(editRegistration);
            // Refresh/reconnect re-reads group summaries without losing branch or grid layout.
            string expansion=Expansion((GridView)Field(soci,"View"));
            Call(analyser,"RefreshCalculatedData");Application.DoEvents();Verify(soci);
            Check(retained==PathState(soci) && expansion==Expansion((GridView)Field(soci,"View")),"Refresh retains branch and expansion");
            Call(analyser,"DisconnectRDS");Call(analyser,"DeferStructuralRefresh");
            Check(sociChart.Series.Count==0 && !tabs.Enabled,"Deferred structural refresh clears stale chart and disables view");
            Call(analyser,"ReconnectRDS",false);
            Call(analyser,"EnsureDeferredAnalysisCurrent");Application.DoEvents();Verify(soci);
            Check(tabs.Enabled,"Explicit refresh restores chart");
            Check(((SimpleButton)Field(soci,"UpButton")).Enabled && ((SimpleButton)Field(soci,"OverviewButton")).Enabled,"Return controls enabled after deferred refresh");
            var activeView=(GridView)Field(soci,"View");
            string filter=activeView.ActiveFilterString;
            activeView.ActiveFilterString="[UseInSOCI] > 999999";Application.DoEvents();
            Check(sociChart.Series.Count==0 && ((IList)Field(soci,"Navigation")).Count==0,"Removed/filtered branch returns safely to empty overview");
            activeView.ActiveFilterString=filter;Application.DoEvents();Verify(soci);
            Check(sociChart.Series.Count>0,"Chart recovers after restoring filter");
            Call(soci,"DrillInto",sociChart.Series.Cast<Series>().First(s=>s.Name.Equals("Rental Income",StringComparison.OrdinalIgnoreCase)).Tag);
            var parser=soci.GetType().GetMethod("TryChartNumber",BindingFlags.Static|BindingFlags.NonPublic);
            foreach(object raw in new object[]{null,DBNull.Value,"#N/A","",Double.NaN,Double.PositiveInfinity})
                Check(!(bool)parser.Invoke(null,new object[]{raw,0d}),"Invalid numeric chart value rejected");
            foreach(int width in new[]{1100,1900,3000}) {
                form.ClientSize=new Size(width,950); Application.DoEvents();
                Check(sociChart.Width>0 && sociChart.Height>200,"Chart fills available space: "+width);
                using(var bmp=new Bitmap(analyser.Width,analyser.Height)) {
                    analyser.DrawToBitmap(bmp,new Rectangle(Point.Empty,bmp.Size));
                    bmp.Save(Path.Combine(args[2],"analyser-"+width+".png"));
                }
            }
        }
        Console.WriteLine("PASS: native chart/grid parity, drill, sources, state preservation, deferred safety and sizing. No source saved.");
    }
    static void Walk(object state,ref int groups,ref int leaves,int depth) {
        if(depth>5)throw new Exception("Unbounded drill");
        Verify(state);
        var chart=(ChartControl)Field(state,"Chart");
        var nodes=chart.Series.Cast<Series>().Select(s=>s.Tag).ToArray();
        foreach(var node in nodes) {
            if((bool)Field(node,"CanDrill")) {
                groups++;Call(state,"DrillInto",node);Walk(state,ref groups,ref leaves,depth+1);Call(state,"GoUp");
            }else leaves++;
        }
    }
    static void Verify(object state) {
        var view=(GridView)Field(state,"View");
        var chart=(ChartControl)Field(state,"Chart");
        if(((LabelControl)Field(state,"Status")).Text.StartsWith("Chart unavailable"))throw new Exception(((LabelControl)Field(state,"Status")).Text);
        foreach(Series series in chart.Series) {
            int handle=(int)Field(series.Tag,"RowHandle");
            var summaries=view.IsGroupRow(handle)?view.GetGroupSummaryValues(handle):null;
            int point=0;
            foreach(GridGroupSummaryItem item in view.GroupSummary) {
                if(item.SummaryType!=DevExpress.Data.SummaryItemType.Sum)continue;
                object raw=summaries==null?view.GetRowCellValue(handle,item.FieldName):summaries[item];
                double value;
                bool valid=raw!=null && raw!=DBNull.Value && Double.TryParse(Convert.ToString(raw),out value);
                var p=series.Points[point++];
                if(valid && (p.IsEmpty || Math.Abs(p.Values[0]-Convert.ToDouble(raw))>0.0001))throw new Exception("Parity: "+series.Name+" "+item.FieldName);
                if(!valid && !p.IsEmpty)throw new Exception("Missing values must be gaps");
            }
            if(point!=series.Points.Count)throw new Exception("Incomplete periods");
        }
    }
    static string Expansion(GridView v) {var parts=new List<string>();for(int h=-1;v.IsValidRowHandle(h);h--)if(v.GetRowExpanded(h))parts.Add(h.ToString());return String.Join(",",parts);}
    static double[] Values(ChartControl chart,string name){return chart.Series.Cast<Series>().First(s=>s.Name.Equals(name,StringComparison.OrdinalIgnoreCase)).Points.Cast<SeriesPoint>().Select(p=>p.Values[0]).ToArray();}
    static string PathState(object o){return String.Join("|",((IList)Field(o,"Navigation")).Cast<string>());}
    static object Field(object o,string n){var t=o.GetType();var f=t.GetField(n,Flags);if(f!=null)return f.GetValue(o);var p=t.GetProperty(n,Flags);if(p!=null)return p.GetValue(o,null);throw new Exception("Missing member "+t.Name+"."+n);}
    static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,Flags).Invoke(o,a);}
    static void Check(bool b,string m){if(!b)throw new Exception(m);Console.WriteLine("PASS: "+m);}
}
