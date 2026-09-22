using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraVerticalGrid;
using DevExpress.XtraVerticalGrid.Rows;
using DevExpress.XtraTab;
using DevExpress.Spreadsheet;
using DevExpress.XtraBars.Docking2010;

public static class EditorNavigationFixture {
    const BindingFlags F=BindingFlags.Instance|BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic;
    static object Field(object o,string name){var f=o.GetType().GetField(name,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(name,F).GetValue(o,null);}
    static object Call(object o,string name,params object[] args){return o.GetType().GetMethod(name,F).Invoke(o,args);}
    static void Check(bool value,string text){if(!value)throw new Exception(text);Console.WriteLine("PASS: "+text);}
    static void CheckToolbar(Control dit,Form owner){
        var panel=(WindowsUIButtonPanel)Field(dit,"WindowsUIButtonPanelActions");
        Func<string> order=()=>String.Join(",",panel.Buttons.Cast<object>().Where(b=>!(b is WindowsUIButton)||((WindowsUIButton)b).Visible).Select(b=>b is WindowsUISeparator?"|":Convert.ToString(((WindowsUIButton)b).Tag)));
        const string expected="MainMenu,|,SaveBP,SaveBPAs,|,Copy,Paste,|,ExportExcel,ExportPdf,|,History,Refresh,Spreadsheet,TogglePanels,|,Options,Help";
        Check(order()==expected,"Exact DIT action order and five separators: "+order());
        var original=panel.Buttons.OfType<WindowsUIButton>().ToArray();
        Call(dit,"OrderActionButtons");
        Check(order()==expected&&original.All(b=>panel.Buttons.Contains(b)),"Toolbar reorder is idempotent and preserves command objects");
        var compact=original.Single(b=>Convert.ToString(b.Tag)=="TogglePanels");
        Check(compact.Caption=="Maximise working area"&&compact.ToolTip=="Maximise working area"&&!compact.UseCaption,"Expanded-state working-area label remains icon-only");
        var icon=compact.ImageOptions.SvgImage;
        Call(owner,"ToggleInterfacePanels");Application.DoEvents();
        Check(compact.Caption=="Restore sidebars"&&compact.ToolTip=="Restore sidebars"&&!compact.UseCaption&&!Object.ReferenceEquals(icon,compact.ImageOptions.SvgImage),"Compacted-state restore label and alternate icon");
        Call(owner,"ToggleInterfacePanels");Application.DoEvents();
        Check(compact.ToolTip=="Maximise working area"&&order()==expected,"Restoring sidebars preserves button order and label");
        panel.Refresh();Application.DoEvents();
        var firstX=new Dictionary<string,int>();
        for(int y=0;y<panel.Height;y+=4)for(int x=0;x<panel.Width;x+=4){
            var hit=panel.CalcHitInfo(new Point(x,y)) as WindowsUIButton;
            if(hit!=null){string tag=Convert.ToString(hit.Tag);if(!firstX.ContainsKey(tag)||x<firstX[tag])firstX[tag]=x;}
        }
        var visible=panel.Buttons.OfType<WindowsUIButton>().Where(b=>b.Visible).ToArray();
        Check(visible.All(b=>firstX.ContainsKey(Convert.ToString(b.Tag))),"Every visible toolbar button has a native hit target");
        var visualOrder=String.Join(",",firstX.OrderBy(pair=>pair.Value).Select(pair=>pair.Key));
        Check(visualOrder==String.Join(",",visible.Select(b=>Convert.ToString(b.Tag))),"Actual rendered left-to-right order: "+visualOrder);
    }
    static List<object> Positions(object dit){return ((IEnumerable)Call(dit,"NavigationPositions")).Cast<object>().ToList();}
    static bool Same(object a,object b){return a!=null && b!=null && Object.ReferenceEquals(Field(a,"Host"),Field(b,"Host")) && (int)Field(a,"X")== (int)Field(b,"X") && (int)Field(a,"Y")== (int)Field(b,"Y");}
    static void AssertFocused(object p){
        var standalone=Field(p,"Standalone") as BaseEdit;
        if(standalone!=null){Check(standalone.ContainsFocus,"Standalone destination has focus");return;}
        var view=Field(p,"View") as GridView;
        var vertical=Field(p,"Host") as VGridControl;
        if(Field(p,"Header")!=null || Field(p,"VHeader")!=null)return;
        Check(view!=null ? view.FocusedRowHandle==(int)Field(p,"RowHandle") && Object.ReferenceEquals(view.FocusedColumn,Field(p,"Column")) : Object.ReferenceEquals(vertical.FocusedRow,Field(p,"VRow")) && vertical.FocusedRecord==(int)Field(p,"Record"),"Destination focus matches permitted cell");
    }
    static BaseEdit Active(object p){
        if(Field(p,"Standalone")!=null)return (BaseEdit)Field(p,"Standalone");
        if(Field(p,"Header")!=null)return (BaseEdit)Field(Field(p,"Header"),"ActiveEditor");
        if(Field(p,"VHeader")!=null)return (BaseEdit)Field(Field(p,"VHeader"),"_ActiveEditor");
        var view=Field(p,"View") as GridView;
        return view!=null?view.ActiveEditor:((VGridControl)Field(p,"Host")).ActiveEditor;
    }
    static List<object> EnterOrder(object dit,bool vertical){
        var points=((IEnumerable)Call(dit,"EnterNavigationPositions")).Cast<object>().ToList();
        var hosts=points.Select(p=>(Control)Field(p,"Host")).Distinct().OrderBy(h=>h.PointToScreen(Point.Empty).Y).ThenBy(h=>h.PointToScreen(Point.Empty).X);
        return hosts.SelectMany(h=>points.Where(p=>Field(p,"Host")==h).OrderBy(p=>(int)Field(p,vertical?"X":"Y")).ThenBy(p=>(int)Field(p,vertical?"Y":"X"))).ToList();
    }
    static object ExpectedEnter(object dit,object origin,bool vertical,bool reverse){
        var order=EnterOrder(dit,vertical);int index=order.FindIndex(p=>Same(p,origin));
        if(index<0)throw new Exception("Independent expected origin is missing");
        return order[(index+(reverse?-1:1)+order.Count)%order.Count];
    }
    static void PressEnter(object dit,object origin,object expected,bool reverse=false,bool closed=false,object edit=null){
        Call(dit,"ActivateEditorPosition",origin);Application.DoEvents();
        var editor=Active(origin);Check(editor!=null,"Enter origin editor available");
        if(edit!=null)editor.EditValue=edit;
        var key=new KeyEventArgs(Keys.Enter|(reverse?Keys.Shift:Keys.None));
        if(closed){
            var view=Field(origin,"View") as GridView;
            if(view!=null){view.CloseEditor();Call(dit,"GridControl_ProcessGridKey",Field(origin,"Host"),key);}
            else{((VGridControl)Field(origin,"Host")).CloseEditor();Call(dit,"VGrid_KeyDown",Field(origin,"Host"),key);}
        }else typeof(Control).GetMethod("OnKeyDown",F).Invoke(editor,new object[]{key});
        Check(key.Handled&&key.SuppressKeyPress,"Enter handled once: reverse="+reverse+", closed="+closed);
        Application.DoEvents();Application.DoEvents();
        AssertFocused(expected);
        var destination=Active(expected);
        Check(destination!=null&&!destination.IsDisposed&&destination.ContainsFocus,"Enter destination editor is focused (including headers)");
        Control page=destination.Parent;
        while(page!=null&&!(page is XtraTabPage))page=page.Parent;
        if(page!=null)Check(page.RectangleToScreen(page.ClientRectangle).IntersectsWith(destination.RectangleToScreen(destination.ClientRectangle)),"Enter destination is scrolled into view");
    }
    static void EnterContract(object dit){
        var initial=EnterOrder(dit,false);
        Check(initial.Count>1,"Enter has multiple permitted destinations");
        var routeCandidates=Call(dit,"EnterNavigationPositions");
        foreach(bool columnMajor in new[]{false,true}){
            var expectedOrder=EnterOrder(dit,columnMajor);
            for(int i=0;i<expectedOrder.Count;i++)foreach(bool reverse in new[]{false,true}){
                var target=Call(dit,"FindEnterTarget",expectedOrder[i],routeCandidates,columnMajor,reverse);
                if(!Same(target,expectedOrder[(i+(reverse?-1:1)+expectedOrder.Count)%expectedOrder.Count]))throw new Exception("Enter route differs from independent circular order");
            }
        }
        Check(true,"Every permitted Enter route matches independent horizontal/vertical forward/reverse order");
        var normal=initial.First(p=>Field(p,"Header")==null&&Field(p,"VHeader")==null&&Field(p,"Standalone")==null);
        foreach(Keys direction in new[]{Keys.Left,Keys.Up,Keys.Right,Keys.Down,Keys.Tab}){
            Exercise(dit,normal,direction);
            bool vertical=direction==Keys.Up||direction==Keys.Down;
            Check((bool)Field(dit,"EnterNavigationVertical")==vertical,"Last navigation axis remembered: "+direction);
            PressEnter(dit,normal,ExpectedEnter(dit,normal,vertical,false));
            PressEnter(dit,normal,ExpectedEnter(dit,normal,vertical,true),true);
        }
        foreach(bool vertical in new[]{false,true}){
            Exercise(dit,normal,vertical?Keys.Down:Keys.Right);
            var order=EnterOrder(dit,vertical);
            PressEnter(dit,order.Last(),order.First());
            PressEnter(dit,order.First(),order.Last(),true);
            // Independent row-/column-major ordering, not the production selector,
            // supplies the boundary expectation and the reverse round trip.
            var host=Field(normal,"Host");
            var own=order.Where(p=>Field(p,"Host")==host).ToList();
            var dimension=vertical?"X":"Y";
            int boundary=own.FindIndex(p=>(int)Field(p,dimension)!=(int)Field(own[0],dimension));
            if(boundary>0){PressEnter(dit,own[boundary-1],own[boundary]);PressEnter(dit,own[boundary],own[boundary-1],true);}
            for(int i=0;i<order.Count-1;i++)if(Field(order[i],"Host")!=Field(order[i+1],"Host")){
                PressEnter(dit,order[i],order[i+1]);PressEnter(dit,order[i+1],order[i],true);break;
            }
            PressEnter(dit,normal,ExpectedEnter(dit,normal,vertical,false),false,true);
        }
        // Validation failure consumes Enter without moving or posting.
        Call(dit,"ActivateEditorPosition",normal);Application.DoEvents();
        var rejected=Active(normal);System.ComponentModel.CancelEventHandler reject=(s,e)=>e.Cancel=true;
        rejected.Validating+=reject;rejected.IsModified=true;
        var rejectedKey=new KeyEventArgs(Keys.Enter);
        typeof(Control).GetMethod("OnKeyDown",F).Invoke(rejected,new object[]{rejectedKey});Application.DoEvents();
        Check(rejectedKey.Handled&&Object.ReferenceEquals(Active(normal),rejected),"Rejected Enter retains original editor");AssertFocused(normal);
        rejected.Validating-=reject;
        rejected.IsModified=true;
        Check(rejected.DoValidate(),"Corrected validation can be accepted");
        PressEnter(dit,normal,ExpectedEnter(dit,normal,true,false));
        var comboPoint=initial.FirstOrDefault(p=>Field(p,"Header")==null&&Field(p,"VHeader")==null&&
            ((Field(p,"Column") as DevExpress.XtraGrid.Columns.GridColumn)!=null&&((DevExpress.XtraGrid.Columns.GridColumn)Field(p,"Column")).ColumnEdit is DevExpress.XtraEditors.Repository.RepositoryItemComboBox ||
             (Field(p,"VRow") as EditorRow)!=null&&((EditorRow)Field(p,"VRow")).Properties.RowEdit is DevExpress.XtraEditors.Repository.RepositoryItemComboBox));
        if(comboPoint!=null){
            Call(dit,"ActivateEditorPosition",comboPoint);Application.DoEvents();
            var combo=(ComboBoxEdit)Active(comboPoint);combo.ShowPopup();Application.DoEvents();
            Check(combo.IsPopupOpen,"Native dropdown open for keyboard test");
            bool axis=(bool)Field(dit,"EnterNavigationVertical");
            Check(!(bool)Call(dit,"NavigateGrid",Field(comboPoint,"Host"),Keys.Enter,null,null)&&!(bool)Call(dit,"NavigateGrid",Field(comboPoint,"Host"),Keys.Down,null,null),"Open dropdown retains native Enter/arrows");
            Check((bool)Field(dit,"EnterNavigationVertical")==axis,"Dropdown selection does not change traversal axis");
            combo.ClosePopup();Application.DoEvents();
        }
        Check(!(bool)Call(dit,"IsNavigationKey",Keys.Control|Keys.Enter)&&!(bool)Call(dit,"IsNavigationKey",Keys.Alt|Keys.Enter),"Ctrl/Alt Enter not hijacked");
        // A value becoming locked after commit must not reset traversal.
        var list=Call(dit,"EnterNavigationPositions");var all=((IEnumerable)list).Cast<object>().ToList();
        var middle=all.Where(p=>Field(p,"Host")==hostOf(normal)).OrderBy(p=>(int)Field(p,"Y")).ThenBy(p=>(int)Field(p,"X")).ToList();
        if(middle.Count>2){var origin=middle[1];((IList)list).Remove(origin);var target=Call(dit,"FindEnterTarget",origin,list,false,false);Check(Same(target,middle[2]),"Removed/locked origin still advances to next coordinates");}
        Console.WriteLine("PASS: Enter contract for actual grid/header editors");
    }
    static object hostOf(object point){return Field(point,"Host");}
    static void Exercise(object dit,object origin,Keys key){
        var candidates=Positions(dit);
        var find=dit.GetType().GetMethod("FindNavigationTarget",F);
        object expected=find.Invoke(null,new object[]{origin,Call(dit,"NavigationPositions"),key});
        Call(dit,"ActivateEditorPosition",origin);Application.DoEvents();
        var regular=Field(origin,"View") as GridView;
        var vertical=Field(origin,"Host") as VGridControl;
        BaseEdit active=regular!=null?regular.ActiveEditor:(vertical!=null?vertical.ActiveEditor:null);
        if(Field(origin,"Header")==null && Field(origin,"VHeader")==null && active!=null) {
            var args=new KeyEventArgs(key);
            typeof(Control).GetMethod("OnKeyDown",F).Invoke(active,new object[]{args});
            Check(args.Handled,"Active editor key handled: "+key);
        }else Check((bool)Call(dit,"NavigateGrid",Field(origin,"Host"),key,Field(origin,"Header"),Field(origin,"VHeader")),"Navigation key handled: "+key);
        Application.DoEvents();
        AssertFocused(expected??origin);
        // A subsequent edit must still be possible after refresh and navigation.
        var v=Field(expected??origin,"Host") as VGridControl;
        if(v!=null && Field(expected??origin,"VHeader")==null)Check(v.ActiveEditor!=null,"Editor remains available after navigation");
    }
    static void Header(object dit,object point) {
        Call(dit,"ActivateEditorPosition",point);Application.DoEvents();
        object helper=Field(point,"VHeader")??Field(point,"Header");
        bool vertical=Field(point,"VHeader")!=null;
        var editor=(BaseEdit)Field(helper,vertical?"_ActiveEditor":"ActiveEditor");
        Check(editor!=null && !editor.IsDisposed,"Header editor opens by keyboard");
        Check(!(editor.Parent is DevExpress.Utils.Layout.TablePanel),"Temporary header editor does not become a layout-table cell");
        var old=editor.EditValue;
        object next=old;
        if(editor is DateEdit)next=((DateEdit)editor).DateTime.AddDays(1);
        else if(editor is ComboBoxEdit) {
            var combo=(ComboBoxEdit)editor;
            next=combo.Properties.Items.Cast<object>().FirstOrDefault(v=>!Object.Equals(v,old) && Convert.ToString(v)!="<Blank>")??old;
        }
        editor.EditValue=next;
        // A value change is buffered until navigation/leave, not a calculation per keystroke.
        Check(Object.Equals(Field(helper,"EditValue"),old),"Header buffers entry until commit");
        Call(helper,vertical?"Editor_KeyDown":"editor_KeyDown",editor,new KeyEventArgs(Keys.Tab));Application.DoEvents();
        var expected=dit.GetType().GetMethod("NormalizeInColumnEditorValue",F).Invoke(null,new object[]{next});
        Console.WriteLine("HEADER old="+old+" next="+expected+" actual="+Field(helper,"EditValue"));
        Check(Convert.ToString(Field(helper,"EditValue"))==Convert.ToString(expected),"Header value commits and survives refresh");
        var target=dit.GetType().GetMethod("FindNavigationTarget",F).Invoke(null,new object[]{point,Call(dit,"NavigationPositions"),Keys.Tab});
        AssertFocused(target??point);
        Call(dit,"ActivateEditorPosition",point);Application.DoEvents();
        editor=(BaseEdit)Field(helper,vertical?"_ActiveEditor":"ActiveEditor");
        if(editor is DateEdit)editor.EditValue=((DateEdit)editor).DateTime.AddDays(1);
        else if(editor is ComboBoxEdit){var c=(ComboBoxEdit)editor;editor.EditValue=c.Properties.Items.Cast<object>().FirstOrDefault(v=>!Object.Equals(v,c.EditValue) && Convert.ToString(v)!="<Blank>")??c.EditValue;}
        var requested=editor.EditValue;
        Check((bool)Call(dit,"CommitEditorsForSave"),"Save accepts pending header value");
        Check(Field(helper,vertical?"_ActiveEditor":"ActiveEditor")==null,"Save closes header editor");
        Check(Convert.ToString(Field(helper,"EditValue"))==Convert.ToString(requested),"Save commits pending header value through existing handler");
    }
    [STAThread] public static int Main(string[] args){
        try {
            AppDomain.CurrentDomain.AssemblyResolve+=(sender,e)=>{string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
            Application.EnableVisualStyles();
            var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
            app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
            var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
            string copy=Path.Combine(args[2],"navigation-copy.xlsb");File.Copy(args[1],copy);
            var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
            Check(!result.BError,"Opened private workbook");
            using(var form=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{0,0,"Normal"})){
                form.Opacity=0;form.ShowInTaskbar=false;form.ClientSize=new Size(1750,1000);form.Show();Application.DoEvents();
                Call(form,"ShowInterface",0,33,false,"None",null,-1);
                var dit=(Control)Field(form,"ActiveInterface");
                CheckToolbar(dit,form);
                Call(dit,"BuildSection",1,false,false);Application.DoEvents();
                var tabs=(XtraTabControl)Field(dit,"XtraTabControlNewGIT");tabs.SelectedTabPageIndex=1;Application.DoEvents();
                var points=Positions(dit);
                Console.WriteLine("METRIC Funding permitted cells="+points.Count+" grids="+points.Select(p=>Field(p,"Host")).Distinct().Count());
                Check(points.Count>10,"Funding editable cells discovered from actual rules/protection");
                if(args.Length>3 && args[3]=="save-only"){
                    dynamic saveModel=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);
                    var saveToolbar=(WindowsUIButtonPanel)Field(dit,"WindowsUIButtonPanelActions");
                    var saveActions=saveToolbar.Buttons.OfType<WindowsUIButton>().Where(b=>Convert.ToString(b.Tag)=="SaveBP" || Convert.ToString(b.Tag)=="SaveBPAs").ToArray();
                    var saveButton=saveActions.Single(b=>Convert.ToString(b.Tag)=="SaveBP");
                    var saveAsButton=saveActions.Single(b=>Convert.ToString(b.Tag)=="SaveBPAs");
                    Check(!saveModel.IsDirty && !saveButton.Enabled && saveAsButton.Enabled,"Clean model disables DIT Save only");
                    using(var fileInstance=(Control)Activator.CreateInstance(app.GetType("FileInstanceInterface"),new object[]{0})){
                    var fileButtons=(WindowsUIButtonPanel)Field(fileInstance,"WindowsUIButtonPanelSaveClose");
                    var fileSave=fileButtons.Buttons.OfType<WindowsUIButton>().Single(b=>Convert.ToString(b.Tag)=="SaveBP");
                    var fileSaveAs=fileButtons.Buttons.OfType<WindowsUIButton>().Single(b=>Convert.ToString(b.Tag)=="SaveBPAs");
                    Check(!fileSave.Enabled && fileSaveAs.Enabled,"Clean model disables File Instance Save only");
                    // Let the existing global first-idle font/layout pass finish
                    // before opening an editor (a human cannot type before it).
                    Application.RaiseIdle(EventArgs.Empty);Application.DoEvents();
                    var savePoint=points.First(p=>Field(p,"VRow")!=null && ((EditorRow)Field(p,"VRow")).Properties.Caption.Replace("\n"," ").Contains("Description") && Field(p,"VHeader")==null);
                    Call(dit,"ActivateEditorPosition",savePoint);Application.DoEvents();
                    var saveGrid=(VGridControl)Field(savePoint,"Host");
                    Check(saveGrid.ActiveEditor!=null,"Blank model description editor available");
                    Application.RaiseIdle(EventArgs.Empty);
                    Check(!saveButton.Enabled,"Opening an unchanged editor keeps Save disabled");
                    saveGrid.ActiveEditor.EditValue="Cancelled pending edit";
                    Application.RaiseIdle(EventArgs.Empty);
                    Check(!saveModel.IsDirty && saveButton.Enabled,"Pending unposted edit enables DIT Save without dirtying workbook");
                    saveGrid.HideEditor();Application.RaiseIdle(EventArgs.Empty);
                    Check(!saveModel.IsDirty && !saveButton.Enabled,"Cancelling an unposted editor restores disabled Save");
                    Call(dit,"ActivateEditorPosition",savePoint);Application.DoEvents();
                    saveGrid.ActiveEditor.EditValue="Blank model save test";
                    Application.RaiseIdle(EventArgs.Empty);
                    Check(saveButton.Enabled && saveAsButton.Enabled,"Save is enabled while the real input remains in its editor");
                    Check(saveActions.Length==2 && saveActions.All(b=>b.ImageOptions.Image!=null),"Blank model has both DIT Save icons");
                    using(var bitmap=new Bitmap(saveToolbar.Width,saveToolbar.Height)){saveToolbar.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(args[2],"dit-toolbar.png"));}
                    Call(dit,"WindowsUIButtonPanelSaveClose_ButtonClick",saveToolbar,new ButtonEventArgs(saveActions.First(b=>Convert.ToString(b.Tag)=="SaveBP")));
                    Check(!saveModel.IsDirty && saveGrid.ActiveEditor==null,"Blank model Save commits editor and clears dirty state");
                    Check(!saveButton.Enabled && !fileSave.Enabled && saveAsButton.Enabled && fileSaveAs.Enabled,"Successful save disables both Save buttons but leaves both Save As buttons enabled");
                    dynamic undone=saveModel.ChangeManager.Undo();
                    Check(!undone.BError && saveButton.Enabled && fileSave.Enabled,"Undo enables Save in both interfaces");
                    dynamic redone=saveModel.ChangeManager.Redo();
                    Check(!redone.BError && saveButton.Enabled && fileSave.Enabled,"Redo keeps both interfaces dirty until saved");
                    using(var saved=new Workbook()){
                        saved.Options.CalculationMode=WorkbookCalculationMode.Manual;saved.LoadDocument(copy);
                        Check(saved.Worksheets["Funding Assumptions"].Cells["E47"].Value.TextValue=="Blank model save test","Blank model Save persists current typed value");
                    }
                    string separate=Path.Combine(args[2],"save-as-copy.xlsb");
                    Check((bool)saveModel.SaveFileAsTo(separate,true) && File.Exists(separate),"Blank model Save As service writes separate copy");
                    Check(!saveButton.Enabled && !fileSave.Enabled,"Save As also restores the disabled clean state");
                    Check((bool)saveModel.SaveFileAsTo(Path.Combine(args[2],"clean-save-as-copy.xlsb"),true),"Save As remains usable on a clean model with a new name");
                    }
                    Application.RaiseIdle(EventArgs.Empty);
                    Console.WriteLine("PASS: Targeted DIT save controls. Only private workbook copies saved.");
                    return 0;
                }
                Check(!(bool)Field(dit,"EnterNavigationVertical"),"New DIT defaults to horizontal Enter");
                EnterContract(dit);
                var payment=points.First(p=>Field(p,"VRow")!=null && ((EditorRow)Field(p,"VRow")).Properties.Caption.Replace("\n"," ").Contains("First Interest Payment Month"));
                var paymentRow=(EditorRow)Field(payment,"VRow");
                var dateRepo=paymentRow.Properties.RowEdit as DevExpress.XtraEditors.Repository.RepositoryItemDateEdit;
                Check(dateRepo!=null && dateRepo.DisplayFormat.FormatString=="mmm" && paymentRow.Properties.DisplayFormat.FormatString=="mmm","Funding month uses a real date editor with mmm display");
                dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);
                IWorkbook book=(IWorkbook)model.WB;
                var originalDate=book.Worksheets["Funding Assumptions"].Cells["E137"].Value;
                var dates=(System.Collections.Generic.HashSet<DateTime>)app.GetType("Abovo.FundingPaymentDateSupport").GetMethod("AllowedDates").Invoke(null,new object[]{book});
                var allowedDate=dates.OrderBy(x=>x).First();
                Call(dit,"ActivateEditorPosition",payment);Application.DoEvents();
                var paymentGrid=(VGridControl)Field(payment,"Host");
                Console.WriteLine("PAYMENT focus="+(paymentGrid.FocusedRow==null?"null":paymentGrid.FocusedRow.Properties.Caption)+" editor="+(paymentGrid.ActiveEditor==null?"null":paymentGrid.ActiveEditor.GetType().Name));
                Check(paymentGrid.ActiveEditor is DateEdit,"Payment cell opens native DateEdit");
                paymentGrid.ActiveEditor.EditValue=allowedDate;
                Check(paymentGrid.PostEditor(),"Native payment month date posted");paymentGrid.CloseEditor();Application.DoEvents();
                Check(book.Worksheets["Funding Assumptions"].Cells["E137"].Value.DateTimeValue==allowedDate,"Real date serial written, not month text");
                dynamic undo=model.ChangeManager.Undo();
                Check(!undo.BError && book.Worksheets["Funding Assumptions"].Cells["E137"].Value==originalDate,"Funding month undo retains original full date");
                dynamic redo=model.ChangeManager.Redo();Check(!redo.BError,"Funding month redo");
                var writer=app.GetType("Abovo.ModelChangeManagerV2").GetMethod("WriteTypedValue",F);
                foreach(object bad in new object[]{allowedDate.AddDays(-1),allowedDate.AddYears(3),"not a date"}) {
                    bool rejected=false;try{writer.Invoke(null,new object[]{book.Worksheets["Funding Assumptions"].Cells["E137"],bad,"D"});}catch(TargetInvocationException){rejected=true;}
                    Check(rejected && book.Worksheets["Funding Assumptions"].Cells["E137"].Value.DateTimeValue==allowedDate,"Typed/pasted invalid payment date rejected without mutation");
                }
                model.ChangeManager.Undo();
                var paymentBefore=book.Worksheets["Funding Assumptions"].Cells["E137"].Value;
                var secondBefore=book.Worksheets["Funding Assumptions"].Cells["F137"].Value;
                var changeType=app.GetType("Abovo.DataChangeEvent");
                Array batch=Array.CreateInstance(changeType,2);
                for(int i=0;i<2;i++) {
                    dynamic entry=Activator.CreateInstance(changeType);
                    entry.ModelID=0;entry.WSName="Funding Assumptions";entry.CellAddress=i==0?"E137":"F137";
                    entry.DataFormat="D";entry.ChangedValue=i==0?allowedDate:allowedDate.AddDays(-1);
                    entry.Description="Payment month paste regression";entry.TimeStamp=DateTime.Now;batch.SetValue(entry,i);
                }
                dynamic failedBatch=Call((object)model.ChangeManager,"ProcessChanges",batch,"Payment month paste regression");
                Check(failedBatch.BError && book.Worksheets["Funding Assumptions"].Cells["E137"].Value==paymentBefore && book.Worksheets["Funding Assumptions"].Cells["F137"].Value==secondBefore,"Invalid date in multi-cell paste rolls back earlier valid date");
                writer.Invoke(null,new object[]{book.Worksheets["Funding Assumptions"].Cells["E137"],allowedDate,"S"});
                Check(book.Worksheets["Funding Assumptions"].Cells["E137"].Value.IsDateTime,"Legacy text definition cannot turn funding dates into text");
                writer.Invoke(null,new object[]{book.Worksheets["Funding Assumptions"].Cells["E137"],"","D"});
                Check(book.Worksheets["Funding Assumptions"].Cells["E137"].Value.IsEmpty,"Excel-allowed blank payment month clears the cell");
                book.Worksheets["Funding Assumptions"].Cells["E137"].Value=paymentBefore; // private test-copy restoration only
                var origin=points.First(p=>Field(p,"VRow")!=null && Field(p,"VHeader")==null && (int)Field(p,"Y")>20);
                Call(dit,"ActivateEditorPosition",origin);Application.DoEvents();
                var grid=(VGridControl)Field(origin,"Host");var row=grid.FocusedRow;int record=grid.FocusedRecord;
                tabs.SelectedTabPage.AutoScrollPosition=new Point(0,450);Application.DoEvents();
                var pageScroll=tabs.SelectedTabPage.AutoScrollPosition;
                Console.WriteLine("PAGE type="+tabs.SelectedTabPage.GetType().FullName);
                int top=grid.TopVisibleRowIndex,left=grid.LeftVisibleRecord;
                Console.WriteLine("PAGE before display="+tabs.SelectedTabPage.DisplayRectangle+" grid="+grid.Bounds);
                Call(dit,"RefreshData",false);
                Console.WriteLine("PAGE immediate "+tabs.SelectedTabPage.AutoScrollPosition+" display="+tabs.SelectedTabPage.DisplayRectangle+" grid="+grid.Bounds);
                Application.DoEvents();
                Console.WriteLine("FOCUS before row="+row.Properties.Caption+" record="+record+" top="+top+" left="+left+"; after row="+(grid.FocusedRow==null?"null":grid.FocusedRow.Properties.Caption)+" record="+grid.FocusedRecord+" top="+grid.TopVisibleRowIndex+" left="+grid.LeftVisibleRecord);
                Check(Object.ReferenceEquals(grid.FocusedRow,row) && grid.FocusedRecord==record && grid.TopVisibleRowIndex==top && grid.LeftVisibleRecord==left,"Refresh retains Funding row, record and both scroll axes");
                Console.WriteLine("PAGE SCROLL "+pageScroll+" -> "+tabs.SelectedTabPage.AutoScrollPosition);
                Check(tabs.SelectedTabPage.AutoScrollPosition==pageScroll,"Refresh retains surrounding page scroll");
                foreach(Keys key in new[]{Keys.Tab,Keys.Shift|Keys.Tab,Keys.Right,Keys.Left,Keys.Up,Keys.Down})Exercise(dit,origin,key);
                var first=points.OrderBy(p=>(int)Field(p,"Y")).ThenBy(p=>(int)Field(p,"X")).First();
                Exercise(dit,first,Keys.Up);
                // Perform a real workbook-backed text edit using the native editor.
                var textPoint=points.First(p=>Field(p,"VRow")!=null && ((EditorRow)Field(p,"VRow")).Properties.Caption.Replace("\n"," ").Contains("Description") && Field(p,"VHeader")==null);
                Call(dit,"ActivateEditorPosition",textPoint);Application.DoEvents();
                var textGrid=(VGridControl)Field(textPoint,"Host");textGrid.ActiveEditor.EditValue="Navigation regression loan";
                Check(textGrid.PostEditor(),"Native Funding text commit accepted");Application.DoEvents();
                Check(Convert.ToString(textGrid.GetCellValue((BaseRow)Field(textPoint,"VRow"),(int)Field(textPoint,"Record")))=="Navigation regression loan","Posted text visible after calculation/refresh");
                AssertFocused(textPoint);Exercise(dit,textPoint,Keys.Right);
                var textCell=book.Worksheets["Funding Assumptions"].Cells["E47"];
                var beforeEnter=textCell.Value;
                PressEnter(dit,textPoint,ExpectedEnter(dit,textPoint,false,false),false,false,"Single Enter commit");
                Check(textCell.Value.TextValue=="Single Enter commit"&&model.IsDirty,"Enter posts native text and marks dirty");
                dynamic enterUndo=model.ChangeManager.Undo();
                Check(!enterUndo.BError&&textCell.Value==beforeEnter,"One Undo reverses the Enter edit (no duplicate history write)");
                model.ChangeManager.Redo();
                Check(textCell.Value.TextValue=="Single Enter commit","Enter edit supports Redo");
                // Check all current nodes can only route to enumerated permitted cells.
                var find=dit.GetType().GetMethod("FindNavigationTarget",F);
                var routes=Call(dit,"NavigationPositions");var allowed=Positions(dit);
                foreach(var point in points)foreach(Keys key in new[]{Keys.Tab,Keys.Right,Keys.Up,Keys.Down}){
                    var target=find.Invoke(null,new object[]{point,routes,key});
                    if(target!=null && !allowed.Any(p=>Same(p,target)))throw new Exception("Navigation escaped permitted cells");
                }
                Check(true,"All Funding routes skip unavailable/locked cells");
                Header(dit,points.First(p=>Field(p,"VHeader")!=null));
                var lastHeader=points.Last(p=>Field(p,"VHeader")!=null);
                Call(dit,"ActivateEditorPosition",lastHeader);Application.DoEvents();
                var lastEditor=(BaseEdit)Field(Field(lastHeader,"VHeader"),"_ActiveEditor");
                Check(lastEditor!=null && tabs.SelectedTabPage.RectangleToScreen(tabs.SelectedTabPage.ClientRectangle).IntersectsWith(lastEditor.RectangleToScreen(lastEditor.ClientRectangle)),"Keyboard can reach and reveal lower header editor");
                Call(dit,"ActivateEditorPosition",textPoint);Application.DoEvents();
                var validating=textGrid.ActiveEditor;
                System.ComponentModel.CancelEventHandler reject=(s,e)=>e.Cancel=true;
                validating.Validating+=reject;
                Check((bool)Call(dit,"NavigateGrid",textGrid,Keys.Right,null,null),"Rejected validation consumes navigation");
                AssertFocused(textPoint);
                validating.Validating-=reject;
                Call(dit,"ActivateEditorPosition",textPoint);Application.DoEvents();
                textGrid.ActiveEditor.EditValue="Rejected save input";
                DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventHandler rejectCell=(s,e)=>{e.Valid=false;e.ErrorText="Test validation rejection";};
                textGrid.ValidatingEditor+=rejectCell;
                Check(!(bool)Call(dit,"CommitEditorsForSave"),"Save refuses an editor with rejected native grid validation");
                textGrid.ValidatingEditor-=rejectCell;
                textGrid.ActiveEditor.EditValue="Navigation regression loan";
                var toolbar=(WindowsUIButtonPanel)Field(dit,"WindowsUIButtonPanelActions");
                var saves=toolbar.Buttons.OfType<WindowsUIButton>().Where(b=>Convert.ToString(b.Tag)=="SaveBP" || Convert.ToString(b.Tag)=="SaveBPAs").ToArray();
                Check(saves.Length==2 && saves.All(b=>b.ImageOptions.Image!=null),"DIT exposes both save buttons with File Instance bitmap icons");
                using(var bitmap=new Bitmap(toolbar.Width,toolbar.Height)){toolbar.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(args[2],"dit-toolbar.png"));}
                Call(dit,"InitialiseSaveActions");
                Check(toolbar.Buttons.OfType<WindowsUIButton>().Count(b=>Convert.ToString(b.Tag)=="SaveBP" || Convert.ToString(b.Tag)=="SaveBPAs")==2,"Save buttons are not duplicated");
                Call(dit,"ActivateEditorPosition",textPoint);Application.DoEvents();
                textGrid.ActiveEditor.EditValue="Saved directly from DIT";
                var fullSaveButton=saves.First(b=>Convert.ToString(b.Tag)=="SaveBP");
                Call(dit,"WindowsUIButtonPanelSaveClose_ButtonClick",toolbar,new ButtonEventArgs(fullSaveButton));
                Check(!model.IsDirty && textGrid.ActiveEditor==null,"Save commits active cell and clears dirty state");
                using(var saved=new Workbook()){
                    saved.Options.CalculationMode=WorkbookCalculationMode.Manual;
                    saved.LoadDocument(copy);
                    Check(saved.Worksheets["Funding Assumptions"].Cells["E47"].Value.TextValue=="Saved directly from DIT","Native Save persists latest typed value to private XLSB");
                }
                string saveAs=Path.Combine(args[2],"save-as-copy.xlsb");
                Check((bool)model.SaveFileAsTo(saveAs,true) && File.Exists(saveAs),"Shared Save As service saves a separate private XLSB");
                using(var bitmap=new Bitmap(dit.Width,dit.Height)){dit.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(args[2],"funding-navigation.png"));}
                // Economic CPI/RPI has two distinct grids: verify boundary movement.
                Call(form,"ShowInterface",0,32,false,"None",null,-1);Application.DoEvents();
                dit=(Control)Field(form,"ActiveInterface");
                tabs=(XtraTabControl)Field(dit,"XtraTabControlNewGIT");tabs.SelectedTabPageIndex=0;Application.DoEvents();
                var ratePoints=Positions(dit);
                var hosts=ratePoints.Select(p=>(Control)Field(p,"Host")).Distinct().OrderBy(c=>c.PointToScreen(Point.Empty).Y).ThenBy(c=>c.PointToScreen(Point.Empty).X).ToList();
                Console.WriteLine("METRIC CPI/RPI grids="+hosts.Count);
                Check(hosts.Count>1,"CPI/RPI provides adjacent grids");
                EnterContract(dit);
                var boundary=ratePoints.Where(p=>Field(p,"Host")==hosts[0]).OrderByDescending(p=>(int)Field(p,"Y")).First();
                find=dit.GetType().GetMethod("FindNavigationTarget",F);
                var below=find.Invoke(null,new object[]{boundary,Call(dit,"NavigationPositions"),Keys.Down});
                Check(below!=null && Field(below,"Host")==hosts[1],"Down at last row reaches next grid");
                Exercise(dit,boundary,Keys.Down);Exercise(dit,below,Keys.Up);
                var lastCell=ratePoints.Where(p=>Field(p,"Host")==hosts[0]).OrderByDescending(p=>(int)Field(p,"Y")).ThenByDescending(p=>(int)Field(p,"X")).First();
                var tabNext=find.Invoke(null,new object[]{lastCell,Call(dit,"NavigationPositions"),Keys.Tab});
                var firstNext=ratePoints.Where(p=>Field(p,"Host")==hosts[1]).OrderBy(p=>(int)Field(p,"Y")).ThenBy(p=>(int)Field(p,"X")).First();
                Check(Same(tabNext,firstNext),"Tab at grid end reaches first editable cell in next grid");
                Exercise(dit,lastCell,Keys.Tab);
                // Existing normal/banded Rent editors must still navigate.
                Call(form,"ShowInterface",0,2,false,"None",null,-1);Application.DoEvents();
                dit=(Control)Field(form,"ActiveInterface");
                tabs=(XtraTabControl)Field(dit,"XtraTabControlNewGIT");tabs.SelectedTabPageIndex=2;Application.DoEvents();
                var rentPoints=Positions(dit);
                Console.WriteLine("METRIC Rent voids permitted cells="+rentPoints.Count);
                Header(dit,rentPoints.First(p=>Field(p,"Header")!=null));
                var cell=rentPoints.First(p=>Field(p,"Header")==null && Field(p,"View")!=null);
                Exercise(dit,cell,Keys.Tab);Exercise(dit,cell,Keys.Down);
                EnterContract(dit);
                // Global assumptions provides actual standalone workbook inputs.
                Call(form,"ShowInterface",0,0,false,"None",null,-1);Application.DoEvents();
                dit=(Control)Field(form,"ActiveInterface");
                var globalInputs=EnterOrder(dit,false);
                var single=globalInputs.First(p=>Field(p,"Standalone")!=null && Field(p,"Standalone") is TextEdit && !(Field(p,"Standalone") is PopupBaseEdit));
                Check(globalInputs.Any(p=>Field(p,"Standalone")!=null),"Global standalone inputs included in Enter traversal");
                var singleEditor=(BaseEdit)Field(single,"Standalone");var singleTag=singleEditor.Tag;
                var tabPreview=new PreviewKeyDownEventArgs(Keys.Tab);
                Call(dit,"NavigationEditorPreviewKeyDown",singleEditor,tabPreview);
                Check(!tabPreview.IsInputKey&&!(bool)Field(dit,"EnterNavigationVertical"),"Standalone Tab retains native dispatch and selects horizontal Enter");
                singleEditor.Properties.ReadOnly=true;
                Check(!EnterOrder(dit,false).Any(p=>Field(p,"Standalone")==singleEditor),"Read-only standalone input excluded");
                singleEditor.Properties.ReadOnly=false;singleEditor.Enabled=false;
                Check(!EnterOrder(dit,false).Any(p=>Field(p,"Standalone")==singleEditor),"Disabled standalone input excluded");
                singleEditor.Enabled=true;
                var singleSheet=(Worksheet)Field(singleTag,"TargetWorksheet");var singleCell=singleSheet.Cells[Convert.ToString(Field(singleTag,"TargetCell"))];
                var oldSingle=singleCell.Value;
                Call(dit,"ActivateEditorPosition",single);Application.DoEvents();
                System.ComponentModel.CancelEventHandler rejectSingle=(s,e)=>e.Cancel=true;
                singleEditor.Validating+=rejectSingle;singleEditor.IsModified=true;
                var invalidSingle=new KeyEventArgs(Keys.Enter);typeof(Control).GetMethod("OnKeyDown",F).Invoke(singleEditor,new object[]{invalidSingle});Application.DoEvents();
                Check(invalidSingle.Handled&&singleEditor.ContainsFocus&&singleCell.Value==oldSingle,"Rejected standalone Enter stays put without writing");
                singleEditor.Validating-=rejectSingle;singleEditor.IsModified=true;Check(singleEditor.DoValidate(),"Standalone validation can recover");
                var singleValue=Convert.ToString(Field(singleTag,"DataType"))=="S"?(object)"Enter standalone test":0.0275;
                PressEnter(dit,single,ExpectedEnter(dit,single,false,false),false,false,singleValue);
                Check(singleCell.Value!=oldSingle,"Standalone Enter commits through existing workbook handler");
                dynamic singleUndo=model.ChangeManager.Undo();Check(!singleUndo.BError&&singleCell.Value==oldSingle,"Standalone Enter edit has one-step Undo");
                globalInputs=EnterOrder(dit,false);
                PressEnter(dit,globalInputs.Last(),globalInputs.First());PressEnter(dit,globalInputs.First(),globalInputs.Last(),true);
            }
            Console.WriteLine("PASS: Native Funding focus/navigation and DIT save controls. Only private workbook copies saved.");return 0;
        }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}
    }
}
