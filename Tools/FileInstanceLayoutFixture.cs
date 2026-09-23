using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Windows.Forms;
using DevExpress.XtraBars.Docking2010;

// UI-only fixture: no workbook is opened, calculated, edited or saved.
public static class FileInstanceLayoutFixture
{
    [STAThread]
    public static int Main(string[] args)
    {
        AppDomain.CurrentDomain.AssemblyResolve += (sender, e) => {
            string path = Path.Combine(args[0], new AssemblyName(e.Name).Name + ".dll");
            return File.Exists(path) ? Assembly.LoadFrom(path) : null;
        };
        try { Run(args); return 0; }
        catch (Exception e) { Console.Error.WriteLine(e); return 1; }
    }

    static void Run(string[] args)
    {
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
        Application.EnableVisualStyles();
        Assembly app = Assembly.LoadFrom(Path.Combine(args[0], "Abovo-summit.exe"));
        Type manager = app.GetType("Abovo.FileManager");
        Type modelType = app.GetType("Abovo.FileManager+ExcelModel");
        object model = FormatterServices.GetUninitializedObject(modelType);
        object structure = Activator.CreateInstance(app.GetType("Abovo.Abovo_Model_Def"));
        var groupProperty = structure.GetType().GetProperty("GroupStructures");
        IList groups = (IList)Activator.CreateInstance(groupProperty.PropertyType);
        foreach (string name in new[] { "Assumptions", "Workings", "Outputs" }) {
            object group = Activator.CreateInstance(app.GetType("Abovo.GroupStructure"));
            group.GetType().GetField("GSName").SetValue(group, name);
            group.GetType().GetField("GSID").SetValue(group, groups.Count.ToString());
            groups.Add(group);
        }
        groupProperty.SetValue(structure, groups, null);
        modelType.GetField("WBStructure").SetValue(model, structure);
        modelType.GetField("Profile").SetValue(model, Activator.CreateInstance(app.GetType("Abovo.AbovoBPWorkbookProfile")));
        Array models = Array.CreateInstance(modelType, 1);
        models.SetValue(model, 0);
        manager.GetField("ExcelModels").SetValue(null, models);

        var fonts = new List<Font>();
        using (var layoutHost = new Form())
        using (Control ui = (Control)Activator.CreateInstance(app.GetType("FileInstanceInterface"), new object[] { 0 })) {
            layoutHost.ShowInTaskbar=false;layoutHost.Opacity=0;layoutHost.ClientSize=new Size(1600,720);
            layoutHost.Controls.Add(ui);layoutHost.Show();Application.DoEvents();
            var top = (WindowsUIButtonPanel)Find(ui, "WindowsUIButtonPanelBPActions");
            var left = (WindowsUIButtonPanel)Find(ui, "WindowsUIButtonPanelSaveClose");
            var badge = (WindowsUIButtonPanel)Find(ui, "WindowsUIButtonPanelBPBadge");
            var browser = Find(ui, "WebBrowserBPInfo");
            Check(Tags(top) == "GoAssumpt,GoWorkings,GoOutputs,GoData,GoFFR,StressTest,Spreadsheet", "Navigation order");
            Check(Tags(left) == "SaveBP,SaveBPAs,CloseBP", "File-action order");
            var save=left.Buttons.OfType<WindowsUIButton>().Single(b=>Convert.ToString(b.Tag)=="SaveBP");
            var saveAs=left.Buttons.OfType<WindowsUIButton>().Single(b=>Convert.ToString(b.Tag)=="SaveBPAs");
            Check(!save.Enabled && saveAs.Enabled,"Clean File Instance Save disabled, Save As available");
            modelType.GetProperty("IsDirty").SetValue(model,true,null);
            Check(save.Enabled && saveAs.Enabled,"Dirty File Instance Save enabled immediately");
            modelType.GetProperty("IsDirty").SetValue(model,false,null);
            Check(!save.Enabled && saveAs.Enabled,"Clear dirty state disables File Instance Save immediately");
            var offer=modelType.GetMethod("OfferSaveAfterCheckSheetClear",BindingFlags.Instance|BindingFlags.NonPublic);
            using(var ditOwner=new Form())using(var ditPanel=new WindowsUIButtonPanel()) {
                ditOwner.Controls.Add(ditPanel);var ditSave=new WindowsUIButton(){Tag="SaveBP"};ditPanel.Buttons.Add(ditSave);
                using(var binding=(IDisposable)Activator.CreateInstance(app.GetType("Abovo.ModelSaveButtonBinding"),BindingFlags.Instance|BindingFlags.Public|BindingFlags.NonPublic,null,new object[]{ditOwner,model,ditPanel,true},null)) {
                    Check(!ditSave.Enabled,"Clean DIT-style Save binding starts disabled");
                    offer.Invoke(model,new object[]{true});
                    Check(save.Enabled&&ditSave.Enabled&&saveAs.Enabled&&!(bool)modelType.GetProperty("IsDirty").GetValue(model,null),"Warning-clear save offer enables File Instance and DIT buttons immediately without dirtying");
                    offer.Invoke(model,new object[]{false});
                    Check(!save.Enabled&&!ditSave.Enabled&&saveAs.Enabled,"Successful-save reset greys both Save buttons; Save As stays available");
                }
            }
            Console.WriteLine("PASS: File Instance and DIT native Save buttons follow dirty and warning-clear transitions; Save As unaffected.");
            if(args.Length>2&&args[2]=="--save-state")return;
            Check(Tags(badge) == "BusinessPlan", "BP placeholder");
            Check(((WindowsUIButton)badge.Buttons[0]).Caption == "HA BP" &&
                ((WindowsUIButton)badge.Buttons[0]).UseCaption, "HA BP model caption");
            var layout = ui.GetType().GetMethod("LayoutFileActionControls", BindingFlags.NonPublic | BindingFlags.Instance);
            foreach (int width in new[] { 750, 1100, 1600 })
            foreach (float points in new[] { 9f, 14f, 18f }) {
                {
                    // Keep every assigned font alive until the controls have been disposed.
                    var font = new Font("Segoe UI", points);
                    fonts.Add(font);
                    ui.Size = new Size(width, 720);
                    foreach (var panel in new[] { top, left, badge }) {
                        panel.Font = font;
                        panel.AppearanceButton.Normal.Font = font;
                        panel.AppearanceButton.Hovered.Font = font;
                        panel.AppearanceButton.Pressed.Font = font;
                    }
                    layout.Invoke(ui, null);
                    ui.CreateControl();
                    foreach (Control control in ui.Controls) control.CreateControl();
                    using (var bitmap = new Bitmap(ui.Width, ui.Height)) {
                        ui.DrawToBitmap(bitmap, new Rectangle(Point.Empty, ui.Size));
                        bitmap.Save(Path.Combine(args[1], "layout-" + width + "-" + points + ".png"));
                    }
                    Check(badge.Top == top.Top, "Badge/navigation top alignment");
                    Check(badge.Left == left.Left && badge.Width == left.Width, "Left column alignment");
                    Check(browser.Left == top.Left && browser.Top == left.Top, "Details alignment");
                    Check(top.Bottom < browser.Top, "Toolbar/details overlap");
                    Check(left.Right < browser.Left, "File actions/details overlap");
                    ui.PerformLayout();Application.DoEvents();
                    Console.WriteLine("CHECK: width="+width+", font="+points+", left="+left.Bounds);
                    AssertButtonsReachable(top);
                    AssertButtonsReachable(left);
                    AssertButtonsReachable(badge);
                    Console.WriteLine("PASS: width=" + width + ", font=" + points + ", top=" + top.Bounds + ", left=" + left.Bounds);
                }
            }
        }
        foreach (var font in fonts) font.Dispose();
        modelType.GetField("Profile").SetValue(model, Activator.CreateInstance(app.GetType("Abovo.AbovoDSAWorkbookProfile")));
        using (Control ui = (Control)Activator.CreateInstance(app.GetType("FileInstanceInterface"), new object[] { 0 })) {
            var badge = (WindowsUIButtonPanel)Find(ui, "WindowsUIButtonPanelBPBadge");
            Check(((WindowsUIButton)badge.Buttons[0]).Caption == "DSA", "DSA model caption");
        }
        Console.WriteLine("PASS: layout, button order and native hit targets; no workbook used.");
        var render=app.GetType("FileInstanceInterface").GetMethod("BuildFileSummaryHtml",BindingFlags.Static|BindingFlags.NonPublic);
        using(var form=new Form())using(var browser=new WebBrowser()) {
            form.Opacity=0;form.ShowInTaskbar=false;form.Controls.Add(browser);browser.Dock=DockStyle.Fill;form.Show();
            foreach(int width in new[]{360,800,1400}) {
                form.ClientSize=new Size(width,800);
                string html=(string)render.Invoke(null,new object[]{"HA Business Plan","Example & Partners <Housing>","2026-04-01",@"C:\Sandbox\A deliberately long folder name for a client business plan\BP v26_0001 - New Blank.xlsb","21/09/2026 12:30:00","17/09/2026 09:33:28","Not recorded","11.68 MB",10f,15f,false,null,false,false});
                Check(!html.Contains("editbpdate") && !html.Contains("<a ") && html.Contains("&amp;") && html.Contains("&lt;Housing&gt;"),"Read-only summary encodes user text, without Edit link");
                browser.DocumentText=html;
                var wait=System.Diagnostics.Stopwatch.StartNew();
                while(wait.ElapsedMilliseconds<2000 && (browser.ReadyState!=WebBrowserReadyState.Complete || browser.Document==null || browser.Document.Body==null)){Application.DoEvents();System.Threading.Thread.Sleep(10);}
                Application.DoEvents();
                Check(browser.Document.Body.InnerText.Contains("Example & Partners <Housing>") && browser.Document.GetElementsByTagName("tr").Count==4,"Summary contains all metadata rows");
                dynamic body=browser.Document.Body.DomElement;
                Check((int)body.scrollWidth<=width,"No horizontal overflow at summary width "+width);
                File.WriteAllText(Path.Combine(args[1],"summary-"+width+".html"),html);
                using(var bitmap=new Bitmap(width,800))using(var g=Graphics.FromImage(bitmap)) {
                    var unk=System.Runtime.InteropServices.Marshal.GetIUnknownForObject(browser.Document.DomDocument);
                    var hdc=g.GetHdc();var bounds=new Rectangle(0,0,width,800);
                    try{int result=OleDraw(unk,1,hdc,ref bounds);Check(result==0,"Native HTML render");}
                    finally{g.ReleaseHdc(hdc);System.Runtime.InteropServices.Marshal.Release(unk);}
                    bitmap.Save(Path.Combine(args[1],"summary-"+width+".png"));
                }
            }
        }
        Console.WriteLine("PASS: modern HTML summary, escaping, no links, responsive native rendering.");
    }
    [System.Runtime.InteropServices.DllImport("ole32.dll")]
    static extern int OleDraw(IntPtr unknown,uint aspect,IntPtr hdc,ref Rectangle bounds);

    static void AssertButtonsReachable(WindowsUIButtonPanel panel)
    {
        var seen = new HashSet<object>();
        for (int y = 0; y < panel.Height; y += 3)
            for (int x = 0; x < panel.Width; x += 3) {
                var hit = panel.CalcHitInfo(new Point(x, y));
                if (hit != null) seen.Add(hit);
            }
        foreach (var button in panel.Buttons.OfType<WindowsUIButton>().Where(b => b.Visible))
            Check(seen.Contains(button), "Unreachable button: " + button.Tag + " in " + panel.Name);
    }
    static string Tags(WindowsUIButtonPanel panel) { return string.Join(",", panel.Buttons.OfType<WindowsUIButton>().Select(b => Convert.ToString(b.Tag))); }
    static Control Find(Control root, string name) { return root.Controls.Find(name, true).Single(); }
    static void Check(bool value, string message) { if (!value) throw new Exception(message); }
}
