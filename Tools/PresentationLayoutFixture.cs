using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using DevExpress.XtraBars.Docking;
using DevExpress.XtraBars.Docking2010;
using DevExpress.XtraBars.Navigation;
using DevExpress.XtraEditors;

public static class PresentationLayoutFixture {
    static Assembly app;
    [STAThread] public static int Main(string[] args) {
        try { Run(args); return 0; } catch(Exception e) { Console.Error.WriteLine(e); return 1; }
    }
    static void Run(string[] args) {
        AppDomain.CurrentDomain.AssemblyResolve += (sender,e) => {
            string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");
            return File.Exists(p)?Assembly.LoadFrom(p):null;
        };
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
        Application.EnableVisualStyles();
        app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
        app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
        Type files=app.GetType("Abovo.FileManager");
        files.GetMethod("Initialise").Invoke(null,new object[]{null});
        MethodInfo open=files.GetMethod("OpenModel");
        dynamic opened=open.Invoke(null,new object[]{args[1],new FileInfo(args[1]),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
        Check(!opened.BError,"Model open");
        dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);
        int funding=-1;
        foreach(dynamic child in model.WBStructure.GroupStructures[0].ChildStructures)
            if(child.CSName=="Funding Assumptions") {
                funding=int.Parse(child.CSID);
                Check(child.NavigationText=="Funding","XML alias");
            }
        Check(funding>=0,"Funding route exists");
        if(args.Length<4 || args[3]!="--stress-only")
        using(Form group=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{0,0,"Normal"}))
        using(Form other=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{0,2,"Normal"})) {
            group.Opacity=0; group.ShowInTaskbar=false; group.ClientSize=new Size(2800,1300);
            other.Opacity=0; other.ShowInTaskbar=false;
            group.Show(); Application.DoEvents();
            var nav=(AccordionControl)Field(group,"AccordionControlNavigator");
            Check(Elements(nav.Elements).Any(x=>x.Text=="Funding"),"Navigator caption");
            Check(Elements(nav.Elements).Where(x=>x.HeaderVisible).All(x=>x.Appearance.Normal.TextOptions.WordWrap==DevExpress.Utils.WordWrap.NoWrap),"No wrapping");
            var sidebar=(AccordionControl)Field(group,"AccordionControlSum");
            Check(sidebar.Elements.Count>=6,"Dynamic history header included");
            var expanded=sidebar.Elements.Select(x=>x.Expanded).ToArray();
            foreach(var item in sidebar.Elements) {
                Check(item.Style==ElementStyle.Item && item.ContentContainer!=null,"Sidebar keeps its content container: "+item.Text);
                Check(item.Appearance.Normal.BackColor==nav.Appearance.Group.Default.BackColor && item.Appearance.Normal.ForeColor==Color.White,"Navigator-matched sidebar header: "+item.Text);
                Check(item.Appearance.Normal.Options.UseBackColor && item.Appearance.Normal.Options.UseForeColor,"Explicit sidebar header appearance");
            }
            var left=(DockPanel)Field(group,"DockPanelNewNavigator");
            var right=(DockPanel)Field(group,"DockPanelDetail");
            left.Visibility=DockVisibility.Visible; left.Width=430;
            right.Visibility=DockVisibility.AutoHide;
            var second=(DockPanel)Field(other,"DockPanelDetail");
            DockVisibility secondState=second.Visibility;
            var states=new[]{(DockPanel)Field(group,"DockPanelNavigator"),left,right};
            var visibility=states.Select(p=>p.Visibility).ToArray();
            int leftWidth=left.Width;
            using(var first=new WindowsUIButtonPanel())
            using(var next=new WindowsUIButtonPanel()) {
                Invoke(group,"AttachPanelsButton",first);
                Invoke(group,"AttachPanelsButton",first);
                Check(first.Buttons.Count==1,"No duplicated toggle");
                var compactIcon=((WindowsUIButton)first.Buttons[0]).ImageOptions.SvgImage;
                ((WindowsUIButton)first.Buttons[0]).ImageOptions.SvgImageSize=new Size(42,42);
                Invoke(group,"ToggleInterfacePanels");
                Check(states.All(p=>p.Visibility==DockVisibility.Hidden),"All panels hidden");
                Check(second.Visibility==secondState,"Other window unchanged");
                Invoke(group,"AttachPanelsButton",next);
                var nextButton=(WindowsUIButton)next.Buttons[0];
                var restoreIcon=nextButton.ImageOptions.SvgImage;
                Check(!Object.ReferenceEquals(compactIcon,restoreIcon),"Restore has a distinct matching icon");
                Check(Object.ReferenceEquals(((WindowsUIButton)first.Buttons[0]).ImageOptions.SvgImage,restoreIcon),"Existing and new interfaces show the same restore icon");
                Check(nextButton.Caption=="Restore" && !nextButton.UseCaption && nextButton.ToolTip.StartsWith("Restore"),"New interface shares icon-only restore state");
                Invoke(group,"ToggleInterfacePanels");
                for(int i=0;i<states.Length;i++) Check(states[i].Visibility==visibility[i],"Visibility restored");
                Check(left.Width==leftWidth,"Navigator width restored");
                var firstButton=(WindowsUIButton)first.Buttons[0];
                Check(firstButton.Caption=="Compact" && !firstButton.UseCaption && firstButton.ToolTip.StartsWith("Compact"),"Icon-only compact button restored");
                Check(firstButton.ImageOptions.HasSvgImage && nextButton.ImageOptions.HasSvgImage,"Panel icon retained in both states");
                Check(Object.ReferenceEquals(firstButton.ImageOptions.SvgImage,compactIcon),"Approved Compact icon restored");
                Check(firstButton.ImageOptions.SvgImageSize==new Size(42,42),"Icon switch preserves display size");
            }
            Console.WriteLine("PASS: XML routing, no-wrap navigator, independent hide/restore, exact visibility/width and new-interface toggle state.");
            if(args.Length>3 && args[3]=="--sidebar-only") {
                right.Visibility=DockVisibility.Visible;
                var wait=System.Diagnostics.Stopwatch.StartNew();
                while(wait.ElapsedMilliseconds<1500){Application.DoEvents();System.Threading.Thread.Sleep(10);}
                foreach(string browserName in new[]{"WebBrowserBPSum","WebBrowserFundSum","WebBrowserFile"}) {
                    var browser=(WebBrowser)Field(group,browserName);
                    Console.WriteLine("BROWSER "+browserName+" visible="+browser.Visible+" bounds="+browser.Bounds+" state="+browser.ReadyState+" text="+(browser.Document==null || browser.Document.Body==null?"<no body>":browser.Document.Body.InnerText));
                    File.WriteAllText(Path.Combine(args[2],browserName+".html"),browser.DocumentText);
                    Check(browser.Document!=null && browser.Document.Body!=null && !String.IsNullOrWhiteSpace(browser.Document.Body.InnerText),"Sidebar document populated on initial open: "+browserName);
                    if(browser.Document!=null && browser.Document.Body!=null) {
                        dynamic dom=browser.Document.Body.DomElement;
                        Console.WriteLine("BROWSERSTYLE "+browserName+" color="+dom.currentStyle.color+" visibility="+dom.currentStyle.visibility+" display="+dom.currentStyle.display+" dimensions="+dom.offsetWidth+"x"+dom.offsetHeight);
                    }
                }
                foreach(int width in new[]{1900,2800,5000,1900,1280}) {
                    group.ClientSize=new Size(width,1200);Invoke(group,"ApplyPresentationScale");Application.DoEvents();
                    Check(sidebar.Elements.Select(x=>x.Expanded).SequenceEqual(expanded),"Sidebar expansion retained on resize");
                    Check(sidebar.Elements.All(x=>x.Appearance.Normal.BackColor==nav.Appearance.Group.Default.BackColor && x.Appearance.Normal.ForeColor==Color.White),"Sidebar colours retained on resize");
                    Save(sidebar,args[2],"sidebar-"+width+".png");
                }
                Invoke(group,"RefreshSummaryData","Automatic");
                Invoke(group,"RefreshSummaryData","Automatic");
                wait.Restart();while(wait.ElapsedMilliseconds<500){Application.DoEvents();System.Threading.Thread.Sleep(10);}
                foreach(string name in new[]{"WebBrowserBPSum","WebBrowserFundSum","WebBrowserFile"}) {
                    var browser=(WebBrowser)Field(group,name);
                    Check(browser.ReadyState==WebBrowserReadyState.Complete && !String.IsNullOrWhiteSpace(browser.Document.Body.InnerText),"Document retained after repeated refresh/resize: "+name);
                }
                var fileBrowser=(WebBrowser)Field(group,"WebBrowserFile");
                Invoke(group,"SetSidebarDocument",fileBrowser,"<html><body>older test</body></html>");
                Invoke(group,"SetSidebarDocument",fileBrowser,"<html><body>latest test</body></html>");
                Application.DoEvents();
                Check(fileBrowser.Document.Body.InnerText=="latest test","Latest queued content wins without blank navigation");
                Invoke(group,"RefreshSummaryData","Automatic");Application.DoEvents();
                Check(fileBrowser.Document.Body.InnerText.Contains("Model details"),"Normal file details restored");
                var messages=Field(group,"SidebarMessageView");
                var messageView=(DevExpress.XtraGrid.Views.Grid.GridView)Field(messages,"MessageGridView");
                var history=Field(group,"SidebarHistoryView");
                var historyView=(DevExpress.XtraGrid.Views.Grid.GridView)Field(history,"HistoryGridView");
                var messageWidths=messageView.VisibleColumns.Select(c=>c.Width).ToArray();
                var historyWidths=historyView.VisibleColumns.Select(c=>c.Width).ToArray();
                var heights=sidebar.Elements.Select(x=>x.ContentContainer.Height).ToArray();
                var record=Activator.CreateInstance(app.GetType("Abovo.SystemMessageRecord"));
                record.GetType().GetProperty("EventID").SetValue(record,Int32.MaxValue,null);
                record.GetType().GetProperty("TimeStamp").SetValue(record,DateTime.Now,null);
                record.GetType().GetProperty("Message").SetValue(record,"Isolated fixture: "+new string('x',300),null);
                ((IList)Field(Field(messages,"MessageManager"),"Items")).Add(record);
                Invoke((object)model.InterfaceHistory,"RecordStandalone",Enum.ToObject(app.GetType("Abovo.InterfaceHistoryDestinationKind"),1),"Isolated fixture history entry with a long caption","FFR");
                for(int pass=0;pass<3;pass++){Invoke(messages,"RefreshMessages");Invoke(history,"RefreshHistory");Application.DoEvents();}
                Check(messageView.VisibleColumns.Select(c=>c.Width).SequenceEqual(messageWidths) && historyView.VisibleColumns.Select(c=>c.Width).SequenceEqual(historyWidths),"Grid refresh retains column widths");
                Check(sidebar.Elements.Select(x=>x.ContentContainer.Height).SequenceEqual(heights),"Grid refresh retains accordion container heights");
                Check(sidebar.AnimationType==AnimationType.None,"Accordion layout animation disabled");
                Invoke(group,"ToggleInterfacePanels");Invoke(group,"ToggleInterfacePanels");Application.DoEvents();
                Check(fileBrowser.Document.Body.InnerText.Contains("Model details"),"Sidebar retains content across hide/restore");
                // A retained second group gets its HTML when its browser handles
                // are first created, rather than borrowing another group's content.
                other.Show();Application.DoEvents();
                wait.Restart();while(wait.ElapsedMilliseconds<500){Application.DoEvents();System.Threading.Thread.Sleep(10);}
                var otherBrowser=(WebBrowser)Field(other,"WebBrowserFile");
                Check(otherBrowser.Document!=null && otherBrowser.Document.Body.InnerText.Contains("Model details"),"Delayed second-window opening retains file details");
                Console.WriteLine("PASS: sidebar headers, content containers, expansion, independent panels and resize.");return;
            }
            int repairs=-1;
            foreach(dynamic child in model.WBStructure.GroupStructures[0].ChildStructures)
                if(((string)child.CSName).StartsWith("Repairs & Maint")) { repairs=int.Parse(child.CSID); break; }
            Check(repairs>=0,"Repairs route exists");
            Invoke(group,"ShowInterface",0,repairs,false,"None",null,-1);
            var repairsDit=(Control)Field(group,"ActiveInterface");
            var toolbar=(WindowsUIButtonPanel)Field(repairsDit,"WindowsUIButtonPanelActions");
            right.Visibility=DockVisibility.Visible;
            group.WindowState=FormWindowState.Normal;
            foreach(int width in new[]{1900,2800,5000,1900,1280}) {
                group.ClientSize=new Size(width,1200);
                Invoke(group,"ApplyPresentationScale");
                Application.DoEvents(); Application.DoEvents();
                Check(sidebar.Elements.Select(x=>x.Expanded).SequenceEqual(expanded),"Sidebar expansion retained on resize");
                Save(group,args[2],"group-repairs-"+width+".png");
                Save(repairsDit,args[2],"dit-repairs-"+width+".png");
                Console.WriteLine("DIT scale check: host="+group.ClientSize+", actual="+app.GetType("Abovo.AbovoAppCls").GetMethod("GetDisplayScale").Invoke(null,new object[]{repairsDit})+", toolbarFont="+toolbar.Font.SizeInPoints+", toolbar="+toolbar.Bounds);
                Check(left.Width+right.Width<=width*0.49,"Panels reserve the document workspace at "+width);
                Check(toolbar.Height<200 || width>=2800,"Compact restored toolbar at "+width);
                Check(toolbar.PointToScreen(new Point(0,toolbar.Height)).Y<=((Control)Field(repairsDit,"XtraTabControlNewGIT")).PointToScreen(Point.Empty).Y,"DIT toolbar does not overlap tabs");
                Console.WriteLine("PASS: group width="+width+", panels="+left.Width+"/"+right.Width+", document="+repairsDit.Width+", toolbar="+toolbar.Bounds+", font="+toolbar.Font.SizeInPoints);
            }
            // Refreshing an unbound/disposed editor must not raise even a caught
            // first-chance NullReferenceException (VS breaks on these).
            int refreshNulls=0;
            EventHandler<System.Runtime.ExceptionServices.FirstChanceExceptionEventArgs> watch=(s,e)=>{
                if(e.Exception is NullReferenceException && e.Exception.StackTrace!=null && e.Exception.StackTrace.Contains("AbovoDESpinEdit.RefreshData"))refreshNulls++;
            };
            AppDomain.CurrentDomain.FirstChanceException+=watch;
            var spinType=app.GetType("Abovo.AbovoExtendedDEControls+AbovoDESpinEdit");
            var spin=(Control)Activator.CreateInstance(spinType);
            Invoke(spin,"RefreshData"); spin.Dispose(); Invoke(spin,"RefreshData");
            AppDomain.CurrentDomain.FirstChanceException-=watch;
            Check(refreshNulls==0,"Spin refresh teardown guard");
            var bpBrowser=(WebBrowser)Field(group,"WebBrowserBPSum");
            dynamic body=bpBrowser.Document.Body.DomElement;
            Console.WriteLine("Sidebar DOM: font="+body.currentStyle.fontSize+", width="+body.offsetWidth+", client="+bpBrowser.ClientSize.Width);
            File.WriteAllText(Path.Combine(args[2],"sidebar-summary.html"),bpBrowser.DocumentText);
            group.ClientSize=new Size(1900,1200); Application.DoEvents();
            Invoke(repairsDit,"BuildSection",2,false,false); // Stock Survey Allocation percentage
            int boundSpins=0;
            foreach(Control editor in Descendants(repairsDit).Where(c=>c.GetType()==spinType)) {
                Check(Field(editor,"TargetWorksheet")!=null && !string.IsNullOrEmpty(Convert.ToString(Field(editor,"TargetCell"))),"Percentage editor bound for refresh");
                Invoke(editor,"RefreshData"); boundSpins++;
            }
            Check(boundSpins>0,"Real percentage editor exercised");
            Console.WriteLine("PASS: bound percentage editors="+boundSpins);
            group.WindowState=FormWindowState.Maximized; Application.DoEvents(); Application.DoEvents();
            group.WindowState=FormWindowState.Normal; Application.DoEvents(); Application.DoEvents();
            Check(toolbar.Height<200 && toolbar.Font.SizeInPoints<12,"GIT native maximise/restore returns compact toolbar");
            Invoke(group,"ToggleInterfacePanels");
            group.ClientSize=new Size(1280,1000); Application.DoEvents();
            Invoke(group,"ToggleInterfacePanels"); Application.DoEvents();
            Check(left.Width+right.Width<=group.ClientSize.Width*0.49,"Hidden panels adapt if the host was resized");
            Invoke(repairsDit,"ReleaseInterfaceResources");
            Invoke(repairsDit,"RefreshData",false);
            Console.WriteLine("PASS: native maximise/restore, hidden-window resize and refresh after resource release.");
            using(Control dit=(Control)Activator.CreateInstance(app.GetType("DataInterfaceTemplate"),new object[]{0,0,funding,group,"Normal",null})) {
                dit.CreateControl();
                Invoke(dit,"BuildSection",1,false,false);
                Invoke(dit,"ResizeFonts");
                var commands=Descendants(dit).OfType<SimpleButton>().Where(b=>b.Text.IndexOf("Funding",StringComparison.OrdinalIgnoreCase)>=0).ToArray();
                Check(commands.Length>0,"Funding command exists");
                foreach(var command in commands) {
                    Check(command.AutoSize,"Command auto sized");
                    Size needed=command.CalcBestSize();
                    Check(command.Width>=needed.Width && command.Height>=needed.Height,"Funding caption fits");
                    Console.WriteLine("PASS: "+command.Text+" font="+command.Font.Size+" size="+command.Size);
                }
                Check(((WindowsUIButtonPanel)Field(dit,"WindowsUIButtonPanelActions")).Buttons.OfType<WindowsUIButton>().Any(b=>object.Equals(b.Tag,"TogglePanels")),"DIT toggle");
                Save(dit,args[2],"funding.png");
            }
            using(Control analyser=(Control)Activator.CreateInstance(app.GetType("BPIncomeExpenditureAnalyserV2"),new object[]{0,group})) {
                Check(((WindowsUIButtonPanel)Field(analyser,"WindowsUIButtonPanelAnalyser")).Buttons.OfType<WindowsUIButton>().Any(b=>object.Equals(b.Tag,"TogglePanels")),"Analyser toggle");
                Console.WriteLine("PASS: analyser toggle attaches to its actual parent.");
            }
        }
        if(args.Length<4 || args[3]!="--group-only")
        using(Form stress=(Form)Activator.CreateInstance(app.GetType("StressTest"),new object[]{0})) {
            stress.Opacity=0; stress.ShowInTaskbar=false; stress.WindowState=FormWindowState.Normal;
            stress.Show(); Application.DoEvents();
            stress.WindowState=FormWindowState.Normal; Application.DoEvents();
            foreach(int width in new[]{1600,2800,5000}) {
                stress.ClientSize=new Size(width,1800);
                Invoke(stress,"ApplyPresentationScale");
                Application.DoEvents();
                Control page=(Control)Field(stress,"XtraTabPageLMVP");
                Control area=(Control)Field(stress,"TablePanelStressInputs");
                Check(area.Width==page.ClientSize.Width && area.Left==0,"Stress fills width "+width);
                Console.WriteLine("PASS: Stress width="+area.Width+", available="+page.ClientSize.Width);
            }
            Save((Control)Field(stress,"TablePanelStressInputs"),args[2],"stress-wide.png");
            // A displayed form can rebuild these controls without changing display scale.
            // This was not covered by the original first-paint/resize-only checks.
            var inputs=(DevExpress.XtraGrid.Views.Grid.GridView)Field(stress,"View_WrapCG_Stresses");
            float expectedFont=inputs.Appearance.Row.Font.SizeInPoints;
            Invoke(stress,"RefreshCovenantSummary");
            var summary=(Control)Field(stress,"CovenantSummaryPanel");
            Check(Descendants(summary).OfType<LabelControl>().All(l=>Math.Abs(l.Appearance.Font.SizeInPoints-expectedFont)<0.1),"Summary font survives recreation");
            Invoke(stress,"ProcessBreachesGrid",true);
            var results=(DevExpress.XtraGrid.Views.Grid.GridView)((DevExpress.XtraGrid.GridControl)Field(stress,"GridControlBreaches")).MainView;
            Check(Math.Abs(results.Appearance.Row.Font.SizeInPoints-expectedFont)<0.1 && results.OptionsView.ColumnAutoWidth,"Results font/width survive rebinding");
            stress.GetType().GetField("STMode",BindingFlags.NonPublic|BindingFlags.Public|BindingFlags.Instance).SetValue(stress,"Y");
            ((Control)Field(stress,"PanelControlCovSel")).Visible=true;
            Invoke(stress,"ApplyPresentationScale");
            Application.DoEvents();
            var navigation=(WindowsUIButtonPanel)Field(stress,"WindowsUIButtonPanelStressNavigator");
            Console.WriteLine("Toolbar: bounds="+navigation.Bounds+", fontHeight="+navigation.Font.Height+", circle="+((DevExpress.Utils.ImageCollection)navigation.ButtonBackgroundImages).ImageSize+", parent="+navigation.Parent.Bounds+", state="+stress.WindowState);
            var table=(DevExpress.Utils.Layout.TablePanel)Field(stress,"TablePanelLMVPlan");
            Console.WriteLine("Toolbar table rows: "+string.Join(", ",table.Rows.Select(r=>r.Style+"="+r.Height)));
            Save(stress,args[2],"stress-after-refresh.png");
            Check(navigation.ContentAlignment==ContentAlignment.MiddleLeft,"Toolbar left aligned");
            Check(navigation.Height>=((DevExpress.Utils.ImageCollection)navigation.ButtonBackgroundImages).ImageSize.Height+navigation.Font.Height,"Toolbar glyph and caption fit");
            Console.WriteLine("PASS: refreshed summary/results typography and toolbar alignment/height.");
            float wideFont=inputs.Appearance.Row.Font.SizeInPoints;
            foreach(Size restored in new[]{new Size(1792,1200),new Size(1280,900),new Size(2800,1500),new Size(5000,1800),new Size(1792,1200)}) {
                stress.ClientSize=restored;
                Invoke(stress,"ApplyPresentationScale");
                Invoke(stress,"RefreshCovenantSummary");
                Invoke(stress,"ProcessBreachesGrid",true);
                Application.DoEvents();
                expectedFont=inputs.Appearance.Row.Font.SizeInPoints;
                Check(Descendants(summary).OfType<LabelControl>().All(l=>Math.Abs(l.Appearance.Font.SizeInPoints-expectedFont)<0.1),"Restored summary matches input font");
                Check(Math.Abs(results.Appearance.Row.Font.SizeInPoints-expectedFont)<0.1,"Restored results match input font");
                if(restored.Width<1920)Check(expectedFont<wideFont,"Restored window reduces automatic density");
                if(restored.Width==5000)Check(Math.Abs(expectedFont-wideFont)<0.1,"Wide window retains approved density");
                Save(stress,args[2],"stress-restored-"+restored.Width+".png");
                AssertButtonsReachable(navigation);
                foreach(string name in new[]{"SimpleButtonCapture","TextEditMultivariableName","ComboBoxBreachMode"}) {
                    Control editor=(Control)Field(stress,name);
                    Rectangle bounds=editor.RectangleToScreen(editor.ClientRectangle);
                    for(Control parent=editor.Parent;parent!=null && parent!=stress;parent=parent.Parent)
                        Check(parent.RectangleToScreen(parent.ClientRectangle).Contains(bounds),"Restored capture control fits: "+name+" at "+restored.Width);
                }
                Check(navigation.PointToScreen(new Point(0,navigation.Height)).Y<=((Control)Field(stress,"XtraTabControlStressTest")).PointToScreen(Point.Empty).Y,"Restored toolbar does not overlap content");
                Console.WriteLine("PASS: restored "+restored+", font="+expectedFont+", toolbar="+navigation.Bounds+", inputAutoWidth="+inputs.OptionsView.ColumnAutoWidth);
            }
            float restoredFont=inputs.Appearance.Row.Font.SizeInPoints;
            stress.WindowState=FormWindowState.Maximized; Application.DoEvents();
            Check(inputs.Appearance.Row.Font.SizeInPoints>=restoredFont,"Maximise restores wide scale");
            stress.WindowState=FormWindowState.Normal; Application.DoEvents();
            Check(Math.Abs(inputs.Appearance.Row.Font.SizeInPoints-restoredFont)<0.1,"Maximise/restore is not cumulative");
            Save(stress,args[2],"stress-maximise-restore.png");
            stress.ClientSize=new Size(5000,1800); Invoke(stress,"ApplyPresentationScale");
            Type stressScales=app.GetType("Abovo.PresentationScaleManager");
            stressScales.GetMethod("ApplyScaleToNewForms",BindingFlags.NonPublic|BindingFlags.Static).Invoke(null,new object[]{null,EventArgs.Empty});
            int originalStressScale=(int)stressScales.GetProperty("InterfaceScalePercent").GetValue(null,null);
            try {
                foreach(int percent in new[]{100,150,200,100}) {
                    stressScales.GetMethod("SetInterfaceScale").Invoke(null,new object[]{percent,false});
                    Invoke(stress,"RefreshCovenantSummary");
                    Invoke(stress,"ProcessBreachesGrid",true);
                    Application.DoEvents();
                    expectedFont=inputs.Appearance.Row.Font.SizeInPoints;
                    Console.WriteLine("Stress fonts at "+percent+": input="+expectedFont+", result="+results.Appearance.Row.Font.SizeInPoints+", summary="+Descendants(summary).OfType<LabelControl>().First().Appearance.Font.SizeInPoints);
                    Check(Descendants(summary).OfType<LabelControl>().All(l=>Math.Abs(l.Appearance.Font.SizeInPoints-expectedFont)<0.1),"Summary refreshed at user scale "+percent);
                    Check(Math.Abs(results.Appearance.Row.Font.SizeInPoints-expectedFont)<0.1,"Results refreshed at user scale "+percent);
                    Check(navigation.Height>=((DevExpress.Utils.ImageCollection)navigation.ButtonBackgroundImages).ImageSize.Height+navigation.Font.Height,"Toolbar fits at user scale "+percent);
                    Check(navigation.PointToScreen(new Point(0,navigation.Height)).Y<=((Control)Field(stress,"XtraTabControlStressTest")).PointToScreen(Point.Empty).Y,"Toolbar does not overlap content");
                    // Transparent fixture forms are not guaranteed a WM_PAINT.
                    // Render first so the native hit map reflects the new metrics.
                    Console.WriteLine("Stress rendered bounds at "+percent+": form="+stress.Bounds+", navigation="+navigation.Bounds+", table="+navigation.Parent.Bounds);
                    Console.WriteLine("Toolbar constraints: min="+navigation.MinimumSize+", max="+navigation.MaximumSize+", table padding="+navigation.Parent.Padding+", columns="+string.Join(",",table.Columns.Select(c=>c.Style+"="+c.Width)));
                    Save(navigation,args[2],"stress-toolbar-"+percent+".png");
                    AssertButtonsReachable(navigation);
                    Console.WriteLine("PASS: Stress refreshed at user scale="+percent+", font="+expectedFont+", toolbar="+navigation.Bounds);
                }
            } finally { stressScales.GetMethod("SetInterfaceScale").Invoke(null,new object[]{originalStressScale,false}); }
        }
        if(args.Length<4 || args[3]!="--stress-only")
        using(Control instance=(Control)Activator.CreateInstance(app.GetType("FileInstanceInterface"),new object[]{0})) {
            instance.Size=new Size(2500,1500); instance.CreateControl();
            Invoke(instance,"PopulateFileInfo");
            var top=(WindowsUIButtonPanel)Field(instance,"WindowsUIButtonPanelBPActions");
            var badge=(WindowsUIButtonPanel)Field(instance,"WindowsUIButtonPanelBPBadge");
            var left=(WindowsUIButtonPanel)Field(instance,"WindowsUIButtonPanelSaveClose");
            Check(top.AppearanceButton.Normal.Font.Size==left.AppearanceButton.Normal.Font.Size &&
                  left.AppearanceButton.Normal.Font.Size==badge.AppearanceButton.Normal.Font.Size,"Consistent native captions");
            Check(top.AppearanceButton.Normal.Font.Size>9,"Initial display font applied on large workspace");
            var browser=(WebBrowser)Field(instance,"WebBrowserBPInfo");
            DateTime deadline=DateTime.UtcNow.AddSeconds(5);
            while(browser.ReadyState!=WebBrowserReadyState.Complete && DateTime.UtcNow<deadline) Application.DoEvents();
            Console.WriteLine("File HTML: tag="+browser.Tag+", ready="+browser.ReadyState+", length="+browser.DocumentText.Length);
            Check(Convert.ToString(browser.Tag)=="SummitOwnsFontScale" && browser.DocumentText.Contains("pt;"),"Single owner of HTML scale");
            Type scales=app.GetType("Abovo.PresentationScaleManager");
            int original=(int)scales.GetProperty("InterfaceScalePercent").GetValue(null,null);
            try {
                foreach(int percent in new[]{100,150,200,100}) {
                    scales.GetMethod("SetInterfaceScale").Invoke(null,new object[]{percent,false});
                    Invoke(instance,"SetScale");
                    float font=top.AppearanceButton.Normal.Font.Size;
                    Size image=((WindowsUIButton)top.Buttons[0]).ImageOptions.Image.Size;
                    Invoke(instance,"SetScale");
                    Check(font==top.AppearanceButton.Normal.Font.Size && image==((WindowsUIButton)top.Buttons[0]).ImageOptions.Image.Size,"Scaling is not cumulative");
                    var circles=(DevExpress.Utils.ImageCollection)top.ButtonBackgroundImages;
                    Check(circles.ImageSize.Width>image.Width,"Glyph fits native button circle");
                    Console.WriteLine("PASS: user scale="+percent+", font="+font+", glyph="+image+", circle="+circles.ImageSize);
                }
            } finally { scales.GetMethod("SetInterfaceScale").Invoke(null,new object[]{original,false}); Invoke(instance,"SetScale"); }
            Save(instance,args[2],"file-instance.png");
        }
        Console.WriteLine("PASS: presentation only; no workbook writes or saves.");
    }
    static IEnumerable<AccordionControlElement> Elements(AccordionControlElementCollection list) {
        foreach(AccordionControlElement e in list) { yield return e; foreach(var c in Elements(e.Elements)) yield return c; }
    }
    static IEnumerable<Control> Descendants(Control root) {
        foreach(Control c in root.Controls) { yield return c; foreach(var n in Descendants(c)) yield return n; }
    }
    static object Field(object obj,string name) {
        for(Type t=obj.GetType();t!=null;t=t.BaseType) {
            var f=t.GetField(name,BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance);
            if(f!=null)return f.GetValue(obj);
            var p=t.GetProperty(name,BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance);
            if(p!=null)return p.GetValue(obj,null);
        } throw new Exception("Missing member "+name);
    }
    static object Invoke(object obj,string name,params object[] args) {
        return obj.GetType().GetMethod(name,BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance).Invoke(obj,args);
    }
    static void Save(Control c,string folder,string name) {
        using(var b=new Bitmap(c.Width,c.Height)) { c.DrawToBitmap(b,new Rectangle(Point.Empty,c.Size)); b.Save(Path.Combine(folder,name)); }
    }
    static void Check(bool ok,string message) { if(!ok)throw new Exception(message); }
    static void AssertButtonsReachable(WindowsUIButtonPanel panel) {
        var seen=new HashSet<object>();
        var firstX=new Dictionary<string,int>();
        for(int y=0;y<panel.Height;y+=5)
            for(int x=0;x<panel.Width;x+=5) {
                var hit=panel.CalcHitInfo(new Point(x,y));
                if(hit!=null) {
                    seen.Add(hit);
                    var button=hit as WindowsUIButton;
                    if(button!=null) {
                        string tag=Convert.ToString(button.Tag);
                        if(!firstX.ContainsKey(tag))firstX[tag]=x;
                    }
                }
            }
        foreach(var button in panel.Buttons.OfType<WindowsUIButton>().Where(b=>b.Visible))
            Check(seen.Contains(button),"Unreachable toolbar button: "+button.Tag);
        Check(firstX["Home"]<firstX["LiMVP"] && firstX["LiMVP"]<firstX["History"],"Toolbar visual order: Home first, history last");
    }
}
