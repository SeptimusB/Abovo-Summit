using System;
using System.IO;
using System.Linq;
using System.Data;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using System.Security.Cryptography;
using System.Text;
using System.IO.Compression;
using System.Xml.Linq;
using DevExpress.Spreadsheet;

public static class RecoveryBackupFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static Assembly app;static Type store,history,files;static string output;static int fundingRecords;
 static object Call(object target,string name,params object[] args){return (target as Type??target.GetType()).GetMethods(F).Single(m=>m.Name==name&&m.IsStatic==(target is Type)&&m.GetParameters().Length==args.Length).Invoke(target is Type?null:target,args);}
 static object Get(object target,string name){return target.GetType().GetProperty(name,F).GetValue(target);}
 static string Hash(string path){using(var h=SHA256.Create())return BitConverter.ToString(h.ComputeHash(File.ReadAllBytes(path)));}
 static string WorkbookDigest(IWorkbook w){using(var hash=SHA256.Create()){
  using(var crypto=new CryptoStream(Stream.Null,hash,CryptoStreamMode.Write))using(var writer=new BinaryWriter(crypto,Encoding.UTF8)){
   foreach(var sheet in w.Worksheets){writer.Write(sheet.Name);writer.Write(sheet.IsProtected);
    foreach(var cell in sheet.GetUsedRange().ExistingCells){if(!cell.HasFormula&&cell.Value.IsEmpty)continue;writer.Write(cell.RowIndex);writer.Write(cell.ColumnIndex);writer.Write(cell.HasFormula);writer.Write(cell.HasFormula?cell.FormulaInvariant:cell.Value.Type+":"+cell.Value.ToString());}
    foreach(var name in sheet.DefinedNames.OrderBy(n=>n.Name,StringComparer.Ordinal)){writer.Write(name.Name);writer.Write(name.RefersTo);}
   }foreach(var name in w.DefinedNames.OrderBy(n=>n.Name,StringComparer.Ordinal)){writer.Write(name.Name);writer.Write(name.RefersTo);}
  }return BitConverter.ToString(hash.Hash);
 }}
 static void Check(bool ok,string message){if(!ok)throw new Exception(message);Console.WriteLine("PASS: "+message);}
 static dynamic Change(dynamic model,string sheet,string cell,double value){dynamic change=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));change.ModelID=(int)model.ModelID;change.WSName=sheet;change.CellAddress=cell;change.ChangedValue=value;change.DataFormat="N";change.Description="Recovery test input";return model.ChangeManager.ProcessChange(change);}
 [STAThread] public static int Main(string[] args){try {
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var path=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(path)?Assembly.LoadFrom(path):null;};
  app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));output=args[1];store=app.GetType("Abovo.RecoveryBackupStore");history=app.GetType("Abovo.RecoveryHistoryStore");files=app.GetType("Abovo.FileManager");
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);Call(files,"Initialise",new object[]{null});
  Trace.Listeners.Add(new ConsoleTraceListener());
  if(args.Length>4&&args[3]=="--funding"){
   fundingRecords=int.Parse(args[4]);if(fundingRecords<1||fundingRecords>100)throw new ArgumentOutOfRangeException("fundingRecords");
   int exceptions=0;AppDomain.CurrentDomain.FirstChanceException+=(s,e)=>{if(e.Exception is InvalidOperationException&&exceptions<5&&(e.Exception.StackTrace??"").Contains("Abovo.")){exceptions++;Console.WriteLine("FUNDING FIRST CHANCE: "+e.Exception);}};
   Console.WriteLine("FUNDING_RECOVERY bitness="+(IntPtr.Size*8)+", records="+fundingRecords);Populated(args[2]);return 0;
  }
  if(args.Length>4&&args[3]=="--values"){var left=ReadModelValues(args[2]);GC.Collect();GC.WaitForPendingFinalizers();var right=ReadModelValues(args[4]);int differences=0;foreach(var key in left.Keys.Union(right.Keys)){SavedValue a,b;if(!left.TryGetValue(key,out a)||!right.TryGetValue(key,out b)||!a.Matches(b)){differences++;if(differences<=10)Console.WriteLine("CALC DIFFERENCE: "+key);}}Check(differences==0,"Recalculated native SOCI, SOFP, cashflow, Check Sheet and Development Expenditure agree across "+left.Count+" cells");return 0;}
  if(args.Length>5&&args[3]=="--sheet-compare"){var left=ReadSheetCells(args[2],args[5]);GC.Collect();GC.WaitForPendingFinalizers();var right=ReadSheetCells(args[4],args[5]);int count=0;foreach(var key in left.Keys.Union(right.Keys).OrderBy(k=>k)){string a,b;left.TryGetValue(key,out a);right.TryGetValue(key,out b);if(a!=b){count++;if(count<=12)Console.WriteLine("DIFFERENCE "+key+"\nLEFT "+a+"\nRIGHT "+b);}}Console.WriteLine("DIFFERENT_CELLS="+count);return count==0?0:1;}
  if(args.Length>4&&args[3]=="--verify-digests"){CheckExcelDigestDetails(args[2],args[4]);return 0;}
  if(args.Length>4&&args[3]=="--compare"){string left=ReadDigest(args[2]);GC.Collect();GC.WaitForPendingFinalizers();string right=ReadDigest(args[4]);if(left==right)Check(true,"Exact native workbook digest across Excel SaveCopyAs");else CheckExcelDigestDetails(Path.Combine(output,Path.GetFileName(args[2])+".digest.xml"),Path.Combine(output,Path.GetFileName(args[4])+".digest.xml"));return 0;}
  if(args.Length>3&&args[3]=="--probe"){Probe(args[2]);return 0;}if(args.Length>2){Populated(args[2]);return 0;}Synthetic();return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
 sealed class SavedValue {public bool Numeric;public double Number;public string Text;public SavedValue(CellValue value){Numeric=value.IsNumeric;if(Numeric)Number=value.NumericValue;else Text=value.ToString(System.Globalization.CultureInfo.InvariantCulture);}public bool Matches(SavedValue other){return Numeric?other.Numeric&&Math.Abs(Number-other.Number)<=1e-7:!other.Numeric&&Text==other.Text;}}
 static System.Collections.Generic.Dictionary<string,SavedValue> ReadModelValues(string path){
  string copy=Path.Combine(output,Path.GetFileName(path));File.Copy(path,copy);var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});Check(!result.BError,"Full native model opens "+Path.GetFileName(path));dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)result.IntegerReturn);IWorkbook w=(IWorkbook)model.WB;
  try{Call((object)model,"EnsureSaveCalculationCurrent",new object[]{null});var values=new System.Collections.Generic.Dictionary<string,SavedValue>();foreach(string name in new[]{"Detailed Comp Inc - Trad View","Financial Position - Trad View","Cashflow detailed","Check Sheet","Development Expenditure"})foreach(var cell in w.Worksheets[name].GetUsedRange().ExistingCells)values.Add(name+"!"+cell.GetReferenceA1(),new SavedValue(cell.Value));return values;}finally{Call(files,"CloseModel",(int)model.ModelID);}
 }
 static System.Collections.Generic.Dictionary<string,string> ReadSheetCells(string path,string sheet){using(var w=new Workbook()){w.Options.CalculationMode=WorkbookCalculationMode.Manual;w.Options.Import.ThrowExceptionOnInvalidDocument=true;Check(w.LoadDocument(path),"Read comparison sheet "+Path.GetFileName(path));return w.Worksheets[sheet].GetUsedRange().ExistingCells.Where(c=>c.HasFormula||!c.Value.IsEmpty).ToDictionary(c=>c.GetReferenceA1(),c=>c.HasFormula?"F:"+c.FormulaInvariant:"V:"+c.Value.Type+":"+c.Value.ToString());}}
 static string ReadDigest(string path){using(var w=new Workbook()){w.Options.CalculationMode=WorkbookCalculationMode.Manual;w.Options.Import.ThrowExceptionOnInvalidDocument=true;Check(w.LoadDocument(path),"Native load "+Path.GetFileName(path));
  var details=new XElement("Workbook");foreach(var sheet in w.Worksheets){using(var h=SHA256.Create()){
   int cells=0;using(var cs=new CryptoStream(Stream.Null,h,CryptoStreamMode.Write))using(var writer=new BinaryWriter(cs,Encoding.UTF8)){foreach(var c in sheet.GetUsedRange().ExistingCells){if(!c.HasFormula&&c.Value.IsEmpty)continue;cells++;writer.Write(c.RowIndex);writer.Write(c.ColumnIndex);writer.Write(c.HasFormula);writer.Write(c.HasFormula?c.FormulaInvariant:c.Value.Type+":"+c.Value.ToString());}}
   var item=new XElement("Sheet",new XAttribute("name",sheet.Name),new XAttribute("protected",sheet.IsProtected),new XAttribute("cells",cells),new XAttribute("hash",BitConverter.ToString(h.Hash)));
   foreach(var n in sheet.DefinedNames.OrderBy(n=>n.Name,StringComparer.Ordinal))item.Add(new XElement("Name",new XAttribute("name",n.Name),new XAttribute("hash",TextHash(n.RefersTo))));details.Add(item);
  }}foreach(var n in w.DefinedNames.OrderBy(n=>n.Name,StringComparer.Ordinal))details.Add(new XElement("Name",new XAttribute("name",n.Name),new XAttribute("hash",TextHash(n.RefersTo))));
  new XDocument(details).Save(Path.Combine(output,Path.GetFileName(path)+".digest.xml"));return WorkbookDigest(w);
 }}
 static string TextHash(string text){using(var h=SHA256.Create())return BitConverter.ToString(h.ComputeHash(Encoding.UTF8.GetBytes(text)));}
 static void CheckExcelDigestDetails(string a,string b){
  var left=XDocument.Load(a);var right=XDocument.Load(b);var known=new[]{"_xlfn.ANCHORARRAY","_xlfn.IFERROR","_xlfn.IFNA","_xlfn.PERCENTILE.EXC","_xlfn.SINGLE","_xlfn.SUMIFS"};int removed=0;
  foreach(var name in left.Root.Elements("Name").ToArray()){
   string id=(string)name.Attribute("name");if(known.Contains(id,StringComparer.Ordinal)&&(string)name.Attribute("hash")==TextHash("=#NAME?")&&!right.Root.Elements("Name").Any(n=>(string)n.Attribute("name")==id)){name.Remove();removed++;Console.WriteLine("EXCEL NORMALIZATION: removed unused compatibility placeholder "+id);}
  }
  Check(XNode.DeepEquals(left,right),"All worksheet formulas/constants/local names/order/protection and business names match; only "+removed+" verified function-name placeholders removed by Excel");
 }
 static void Probe(string path){
  AppDomain.CurrentDomain.FirstChanceException+=(s,e)=>{if(e.Exception is ArgumentOutOfRangeException)Console.WriteLine("IMPORT FIRST CHANCE: "+e.Exception);};
  using(var w=new Workbook()){w.Options.CalculationMode=WorkbookCalculationMode.Manual;w.Options.Import.ThrowExceptionOnInvalidDocument=true;bool loaded=w.LoadDocument(path);Console.WriteLine("CORE loaded="+loaded+" sheets="+w.Worksheets.Count+" Global="+w.Worksheets.Contains("Global Assumptions"));}
  GC.Collect();GC.WaitForPendingFinalizers();
  var constructor=app.GetType("Abovo.FileManager+ExcelModel").GetConstructors().Single();dynamic model=constructor.Invoke(new object[]{0,Enum.ToObject(constructor.GetParameters()[1].ParameterType,0)});
  bool controlLoaded=model.ModelSpreadsheetControl.LoadDocument(path);IWorkbook wb=(IWorkbook)model.WB;Console.WriteLine("CONTROL loaded="+controlLoaded+" sheets="+wb.Worksheets.Count+" Global="+wb.Worksheets.Contains("Global Assumptions"));
  model.CloseModel();
 }
 static void Synthetic(){
  var modelType=app.GetType("Abovo.FileManager+ExcelModel");var constructor=modelType.GetConstructors().Single();
  dynamic model=constructor.Invoke(new object[]{0,Enum.ToObject(constructor.GetParameters()[1].ParameterType,0)});
  var models=Array.CreateInstance(modelType,1);models.SetValue(model,0);files.GetField("ExcelModels").SetValue(null,models);files.GetField("ExcelModelCount").SetValue(null,0);
  IWorkbook w=(IWorkbook)model.WB;var sheet=w.Worksheets[0];sheet.Name="Inputs";sheet.Cells["A1"].Value=10;sheet.Cells["B1"].FormulaInvariant="=A1*2";sheet.Cells["C1"].NumberFormat="0.00";
  string source=Path.Combine(output,"source.xlsb");model.ModelSpreadsheetControl.SaveDocument(source,DocumentFormat.Xlsb);model.FileName=source;model.FileInfo=new FileInfo(source);
  dynamic manager=Activator.CreateInstance(app.GetType("Abovo.ModelChangeManagerV2"),new object[]{0});model.ChangeManager=manager;
  Check(Change(model,"Inputs","A1",15).BSuccess,"Typed edit uses the normal ChangeManager");
  sheet.Cells["D1"].FormulaInvariant="=CONCATENATE("+String.Join(",",Enumerable.Repeat("A1",37))+")";
  var protectedSheet=w.Worksheets.Add("Protected");protectedSheet.Cells["A1"].Value="keep";protectedSheet.Protect("fixture-only",WorksheetProtectionPermissions.Default);
  Application.DoEvents();model.IsDirty=true;
  string originalHash=Hash(source),path=w.Path;long revision=(long)Get(model,"CalculationRevision");int count=w.History.Count;bool modified=model.ModelSpreadsheetControl.Modified;
  var mode=w.Options.CalculationMode;var engine=w.Options.CalculationEngineType;bool formulaCheck=model.NeedsFormulaPreflight;
  string backup=(string)Call(store,"Write",model);
  Console.WriteLine("STATE immediate dirty="+model.IsDirty+" modified="+model.ModelSpreadsheetControl.Modified+" revision="+Get(model,"CalculationRevision")+" expected modified="+modified+" revision="+revision);
  Application.DoEvents();
  Console.WriteLine("STATE pumped dirty="+model.IsDirty+" modified="+model.ModelSpreadsheetControl.Modified+" revision="+Get(model,"CalculationRevision"));
  Check((bool)model.IsDirty&&model.ModelSpreadsheetControl.Modified==modified&&(long)Get(model,"CalculationRevision")==revision,"Recovery leaves dirty/native/revision state intact after queued events");
  Check(model.FileName==source&&w.Path==path&&Hash(source)==originalHash,"Original bytes and both model paths are unchanged");
  Check(w.History.Count==count&&manager.CanUndo&&model.NeedsFormulaPreflight==formulaCheck,"Native history, managed Undo and XLSB preflight state preserved");
  Check(w.Options.CalculationMode==mode&&w.Options.CalculationEngineType==engine&&protectedSheet.IsProtected,"Calculation and protection state preserved");
  Check(!w.DocumentProperties.Custom.Names.Contains("Abovo.Summit.RecoverySource")&&!w.DocumentProperties.Custom.Names.Contains("Abovo.Summit.ResultsPending"),"Temporary recovery metadata removed from live workbook");
  using(var read=new Workbook()){read.Options.CalculationMode=WorkbookCalculationMode.Manual;read.LoadDocument(backup);Check(read.Worksheets[0].Cells["A1"].Value.NumericValue==15&&read.Worksheets[0].Cells["D1"].FormulaInvariant==sheet.Cells["D1"].FormulaInvariant,"XLSM contains edited inputs and long CONCATENATE formula");Check(read.DocumentProperties.Custom["Abovo.Summit.ResultsPending"].BooleanValue,"Recovery requires fresh calculated results on Summit reopen");}
  Check((string)Call(store,"ReadSource",backup)==source,"Recovery provenance matches its original");
  File.SetLastWriteTimeUtc(source,DateTime.UtcNow.AddMinutes(-2));File.SetLastWriteTimeUtc(backup,DateTime.UtcNow);
  Check((string)Call(store,"NewerRecovery",source)==backup,"A newer owned backup is offered");
  File.SetLastWriteTimeUtc(source,DateTime.UtcNow.AddMinutes(2));Check(Call(store,"NewerRecovery",source)==null,"An older backup is not offered");File.SetLastWriteTimeUtc(source,DateTime.UtcNow.AddMinutes(-2));
  DataTable recovered=(DataTable)Call(history,"Read",backup,(DataTable)manager.GetHistoryTable());
  Check(recovered.Rows.Count==1&&(string)recovered.Rows[0]["OriginalValue"]=="10"&&(string)recovered.Rows[0]["NewValue"]=="15","Prior-session history records last committed change");
  Check((string)recovered.Rows[0]["Action"]==""&&((string)recovered.Rows[0]["State"]).StartsWith("Recovered:"),"Recovered rows carry no executable Undo/Redo action");
  Check(manager.Undo().BSuccess&&sheet.Cells["A1"].Value.NumericValue==10,"Live Undo still works after recovery saving");
  manager=Activator.CreateInstance(app.GetType("Abovo.ModelChangeManagerV2"),new object[]{0});model.ChangeManager=manager;Call(manager,"RestoreRecoveryHistory",recovered);
  Check(!manager.CanUndo&&!manager.CanRedo&&((DataTable)manager.GetHistoryTable()).Rows.Count==1,"Prior-session display history is not restored into Undo stacks");
  Check(Change(model,"Inputs","A1",25).BSuccess&&manager.CanUndo&&manager.Undo().BSuccess,"New edits retain normal Undo after restoring audit history");
  string goodHash=Hash(backup);bool failed=false;
  using(var locked=new FileStream(backup,FileMode.Open,FileAccess.Read,FileShare.None)){try{Call(store,"Write",model);}catch(TargetInvocationException){failed=true;}}
  Check(failed&&Hash(backup)==goodHash&&model.IsDirty,"Locked backup failure preserves the previous complete recovery and dirty state");
  using(var fail=new ThrowingStream()){failed=false;try{Call(model,"WriteRecoverySnapshot",fail);}catch(TargetInvocationException){failed=true;}Check(failed&&model.IsDirty&&w.Path==path&&!w.DocumentProperties.Custom.Names.Contains("Abovo.Summit.RecoverySource"),"Serialization failure restores metadata and working-file state");}
  model.IsDirty=false;failed=false;try{Call(store,"Write",model);}catch(TargetInvocationException){failed=true;}Check(failed,"Clean model does not create redundant recovery");model.IsDirty=true;
  using((IDisposable)manager.BeginChangeGroup("pending")){failed=false;try{Call(store,"Write",model);}catch(TargetInvocationException){failed=true;}Check(failed,"An in-progress grouped edit cannot be backed up");}
  Scheduler(model,backup);
  string retained=Path.Combine(output,"retained-recovery.xlsm");File.Move(backup,retained);File.WriteAllText(backup,"Not a Summit backup");failed=false;try{Call(store,"Write",model);}catch(TargetInvocationException){failed=true;}Check(failed&&File.ReadAllText(backup)=="Not a Summit backup","Unrelated filename collision is not overwritten");
  using(var options=(Form)Activator.CreateInstance(app.GetType("Abovo.ApplicationOptionsForm"))){options.Show();Application.DoEvents();dynamic tabs=options.Controls[0].Controls[0];tabs.SelectedTabPageIndex=1;Application.DoEvents();dynamic enabled=options.GetType().GetField("BackupEnabled",F).GetValue(options);dynamic interval=options.GetType().GetField("BackupMinutes",F).GetValue(options);Check(!enabled.Checked&&Convert.ToInt32(interval.EditValue)==10,"Recovery is opt-in and interval displays 10 minutes");enabled.Checked=true;Check(interval.Enabled,"Interval enabled with backup checkbox");Check(interval.Properties.MinValue==1&&interval.Properties.MaxValue==120,"Interval limited to 1-120 minutes");using(var bitmap=new System.Drawing.Bitmap(options.Width,options.Height)){options.DrawToBitmap(bitmap,new System.Drawing.Rectangle(0,0,options.Width,options.Height));bitmap.Save(Path.Combine(output,"options.png"));}options.Close();Check(!(bool)app.GetType("Abovo.RecoveryBackupManager").GetProperty("Enabled",F).GetValue(null),"Closing Options without Apply does not enable recovery");}
  model.CloseModel();MetadataGuard();Console.WriteLine("PASS: synthetic recovery tests");
 }
 static string EntryHash(string path,string part){using(var zip=ZipFile.OpenRead(path))using(var stream=zip.GetEntry(part).Open())using(var h=SHA256.Create())return BitConverter.ToString(h.ComputeHash(stream));}
 static void MetadataGuard(){
  string path=Path.Combine(output,"metadata-test.xlsm");using(var w=new Workbook()){w.Worksheets[0].Cells["A1"].DynamicArrayFormulaInvariant="=SEQUENCE(1)";w.SaveDocument(path,DocumentFormat.Xlsm);}
  string sheet=EntryHash(path,"xl/worksheets/sheet1.xml");
  using(var zip=ZipFile.Open(path,ZipArchiveMode.Update)){
   zip.GetEntry("xl/metadata.xml").Delete();
   foreach(string part in new[]{"xl/_rels/workbook.xml.rels","[Content_Types].xml"}){
    XDocument xml;using(var stream=zip.GetEntry(part).Open())xml=XDocument.Load(stream);
    xml.Root.Elements().Where(e=>((string)e.Attribute("Type")??"").EndsWith("/sheetMetadata")||(string)e.Attribute("PartName")=="/xl/metadata.xml").Remove();
    zip.GetEntry(part).Delete();using(var stream=zip.CreateEntry(part).Open())xml.Save(stream);
   }
  }
  var compatibility=app.GetType("Abovo.RecoveryXlsmCompatibility");bool rejected=false;
  try{Call(compatibility,"Prepare",path,false);}catch(TargetInvocationException){rejected=true;}Check(rejected,"Unverified binary metadata is not guessed");
  Call(compatibility,"Prepare",path,true);Check(EntryHash(path,"xl/worksheets/sheet1.xml")==sheet,"Metadata bridge leaves worksheet bytes unchanged");
  using(var read=new Workbook()){read.Options.Import.ThrowExceptionOnInvalidDocument=true;Check(read.LoadDocument(path)&&read.Worksheets[0].Cells["A1"].HasDynamicArrayFormula,"Public-API metadata restores dynamic-array import");}
  string stable=Hash(path);Call(compatibility,"Prepare",path,false);Check(Hash(path)==stable,"Existing metadata is left untouched");
 }
 static void Scheduler(dynamic model,string backup){
  var type=app.GetType("Abovo.RecoveryBackupManager");Call(type,"Configure",false,10,false);Call(type,"Track",(object)model);
  var plans=(System.Collections.IDictionary)type.GetField("Plans",F).GetValue(null);object state=plans[(object)model];
  Action due=()=>{state.GetType().GetField("DueUtc",F).SetValue(state,DateTime.UtcNow.AddMinutes(-1));type.GetField("LastInputUtc",F).SetValue(null,DateTime.UtcNow.AddSeconds(-30));};
  Action tick=()=>Call(type,"Tick",null,EventArgs.Empty);
  string old=Hash(backup);due();tick();Check(Hash(backup)==old,"Disabled timer does not save");
  Call(type,"Configure",true,10,false);due();type.GetField("LastInputUtc",F).SetValue(null,DateTime.UtcNow);tick();Check(Hash(backup)==old,"Timer waits for user inactivity");
  due();using((IDisposable)model.ChangeManager.BeginChangeGroup("timer protection")){tick();Check(Hash(backup)==old,"Timer defers during pending grouped edit");}
  due();files.GetField("InternalBIsSaving",F).SetValue(null,true);try{tick();Check(Hash(backup)==old,"Timer defers during normal Save");}finally{files.GetField("InternalBIsSaving",F).SetValue(null,false);}
  Check(Change(model,"Inputs","A1",40).BSuccess,"New committed input before scheduled backup");due();tick();Check(Hash(backup)==old,"No silent scheduled save when no notification window is available");
  using(var owner=new Form()){owner.ShowInTaskbar=false;owner.Opacity=0;owner.Show();Application.DoEvents();
   owner.WindowState=FormWindowState.Minimized;Application.DoEvents();due();tick();Check(Hash(backup)==old,"Timer defers when all notification windows are minimised");
   owner.WindowState=FormWindowState.Normal;Application.DoEvents();
   Check(Object.ReferenceEquals(Call(type,"NoticeOwner",(object)model),owner),"Recovery resolves a live notification owner");
   due();tick();System.Threading.Thread.Sleep(1200);Application.DoEvents();owner.Close();}
  using(var messages=(IDisposable)Call(app.GetType("Abovo.SystemMessageManager"),"Acquire",(int)model.ModelID)){
   var text=((System.Collections.IEnumerable)Call(messages,"SnapshotItems")).Cast<object>().Select(x=>(string)Get(x,"Message")).ToArray();
   Check(text.Any(x=>x.StartsWith("Saving recovery copy:"))&&text.Any(x=>x.StartsWith("Recovery copy saved:")),"Recovery start and completion recorded in System Messages");}
  Check(Hash(backup)!=old&&(long)state.GetType().GetField("Revision",F).GetValue(state)==(long)Get(model,"CalculationRevision"),"Due idle timer saves the committed revision");old=Hash(backup);due();tick();Check(Hash(backup)==old,"Timer skips unchanged revision even while model remains dirty");
  string previous=model.FileName;model.FileName=Path.Combine(output,"renamed.xlsb");due();tick();Check((string)state.GetType().GetField("Path",F).GetValue(state)==(string)model.FileName&&(DateTime)state.GetType().GetField("DueUtc",F).GetValue(state)>DateTime.UtcNow&&!File.Exists((string)Call(store,"RecoveryPath",(string)model.FileName)),"Save As changes backup destination and restarts interval");model.FileName=previous;
  Call(type,"Configure",false,10,false);Call(type,"Forget",(object)model);Check(!plans.Contains((object)model),"Closing model removes its schedule");
 }
 static void Populated(string input){
  string source=Path.Combine(output,"populated.xlsb");File.Copy(input,source);string original=Hash(source);
  var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{source,new FileInfo(source),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});Check(!result.BError,"Open private populated model");
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);IWorkbook w=model.WB;
  Application.DoEvents(); // Drain load events BEFORE the edit/Undo baseline.
  string recordName=null;int expectedColumns=0;
  if(fundingRecords>0){
   object rule=Call((object)model.WorkbookStructureRules,"GetRule","FUNDING_RECORDS");
   recordName=(string)rule.GetType().GetField("RecordCountNamedRange",F).GetValue(rule);
   expectedColumns=w.DefinedNames.GetDefinedName(recordName).Range.ColumnCount+fundingRecords;
   var calc=w.Options.CalculationMode;var engine=w.Options.CalculationEngineType;
   var protection=w.Worksheets.ToDictionary(s=>s.Name,s=>s.IsProtected);
   var timerInsert=Stopwatch.StartNew();dynamic inserted=model.WorkbookStructureRules.AddRecords("FUNDING_RECORDS",fundingRecords);
   Console.WriteLine("FUNDING_TOTAL_MS="+timerInsert.ElapsedMilliseconds);
   Check(!inserted.BError,"Funding insertion succeeds: "+inserted.StringReturn);
   Check(w.DefinedNames.GetDefinedName(recordName).Range.ColumnCount==expectedColumns,"Funding record count increased by requested amount");
   Check(w.Options.CalculationMode==calc&&w.Options.CalculationEngineType==engine&&w.Worksheets.All(s=>s.IsProtected==protection[s.Name]),"Funding restores calculation engine/mode and all worksheet protection states");
   var issues=((System.Collections.IEnumerable)Call((object)model.TransDBSync,"InspectMirrorGeometry",w)).Cast<string>().ToArray();
   Check(issues.Length==0,"All supported Transactional DB mirror geometries agree after Funding insertion: "+String.Join("; ",issues));
  }
  var target=w.DefinedNames.GetDefinedName("CurrStNo").Range[0,0];double before=target.Value.NumericValue;
  Check(Change(model,target.Worksheet.Name,target.GetReferenceA1(),before+1).BSuccess,"Populated committed stock edit");
  string expected=WorkbookDigest(w);var timer=Stopwatch.StartNew();string backup=fundingRecords>0?ScheduledPopulatedSave((object)model):(string)Call(store,"Write",model);Console.WriteLine("RECOVERY_TOTAL_MS="+timer.ElapsedMilliseconds);
  Check(model.IsDirty&&model.FileName==source&&Hash(source)==original&&model.ChangeManager.CanUndo,"Populated backup leaves original, dirty flag and Undo intact");
  File.SetLastWriteTimeUtc(source,DateTime.UtcNow.AddMinutes(-2));Check((string)Call(store,"NewerRecovery",source)==backup,"Populated recovery discovery");
  Call(files,"CloseModel",0);
  model=null;w=null;target=null;GC.Collect();GC.WaitForPendingFinalizers();
  using(var read=new Workbook()){read.Options.CalculationMode=WorkbookCalculationMode.Manual;read.Options.Import.ThrowExceptionOnInvalidDocument=true;Check(read.LoadDocument(backup),"Recovery package loads independently");Check(WorkbookDigest(read)==expected,"All formula text, constant inputs, global/local names, sheet order and protection preserved");}
  GC.Collect();GC.WaitForPendingFinalizers();
  result=open.Invoke(null,new object[]{backup,new FileInfo(backup),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});Check(!result.BError,"Recovered XLSM loads as a full Summit model: "+result.StringReturn);model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)result.IntegerReturn);w=model.WB;
  Check(model.RecoverySaveAsRequired&&model.RecoverySourcePath==source&&model.IsDirty,"Recovered model explicitly requires Save As to XLSB");
  Check((int)Get(model.ChangeManager,"RecoveryHistoryCount")==1&&!model.ChangeManager.CanUndo,"Recovered populated history is present and read-only");
  Check(w.DefinedNames.GetDefinedName("CurrStNo").Range[0,0].Value.NumericValue==before+1&&!model.DeferredSaveResultsPending,"Committed inputs recovered and pending-result gate completed");
  if(recordName!=null){Check(w.DefinedNames.GetDefinedName(recordName).Range.ColumnCount==expectedColumns,"Recovered model retains all inserted Funding columns");var issues=((System.Collections.IEnumerable)Call((object)model.TransDBSync,"InspectMirrorGeometry",w)).Cast<string>().ToArray();Check(issues.Length==0,"Recovered model retains matching Transactional DB geometry");}
  string normal=Path.Combine(output,"recovered-normal.xlsb");Check(model.SaveFileAsTo(normal,true),"Recovered model saves as normal XLSB through compatibility policy");
  Check(String.IsNullOrEmpty((string)model.RecoverySourcePath)&&!model.RecoverySaveAsRequired&&!model.IsDirty&&Hash(source)==original,"Save As clears recovery guidance without overwriting the original");
  Call(files,"CloseModel",(int)model.ModelID);Console.WriteLine("BACKUP="+backup);Console.WriteLine("NORMAL="+normal);
 }
 static string ScheduledPopulatedSave(dynamic model){
  var manager=app.GetType("Abovo.RecoveryBackupManager");var safety=app.GetType("Abovo.ModelSafetyManager");
  string backup=(string)Call(store,"RecoveryPath",(string)model.FileName);
  Call(manager,"Configure",true,1,false);Call(manager,"Track",(object)model);
  var plans=(System.Collections.IDictionary)manager.GetField("Plans",F).GetValue(null);object state=plans[(object)model];
  Action due=()=>{state.GetType().GetField("DueUtc",F).SetValue(state,DateTime.UtcNow.AddMinutes(-1));manager.GetField("LastInputUtc",F).SetValue(null,DateTime.UtcNow.AddSeconds(-30));};
  Action tick=()=>Call(manager,"Tick",null,EventArgs.Empty);
  try{
   using(var owner=new Form()){owner.ShowInTaskbar=false;owner.Opacity=0;owner.Show();Application.DoEvents();
    Call(safety,"BeginBulkWorkbookMutation",(int)model.ModelID);
    try{due();tick();Check(!File.Exists(backup),"Due recovery is deferred during structural mutation");}finally{Call(safety,"EndBulkWorkbookMutation",(int)model.ModelID);}
    due();files.GetField("InternalBIsSaving",F).SetValue(null,true);try{tick();Check(!File.Exists(backup),"Due recovery is deferred during normal saving");}finally{files.GetField("InternalBIsSaving",F).SetValue(null,false);}
    due();manager.GetField("LastInputUtc",F).SetValue(null,DateTime.UtcNow);tick();Check(!File.Exists(backup),"Due recovery waits for idle after user activity");
    due();var timer=Stopwatch.StartNew();tick();Console.WriteLine("SCHEDULED_RECOVERY_MS="+timer.ElapsedMilliseconds);
    Check(File.Exists(backup),"Due idle scheduler writes recovery after Funding has finished");
    string savedHash=Hash(backup);due();tick();Check(Hash(backup)==savedHash,"No duplicate recovery for unchanged revision");
    System.Threading.Thread.Sleep(1200);Application.DoEvents();owner.Close(); // Atomic save has already returned.
   }
   using(var lease=(IDisposable)Call(app.GetType("Abovo.SystemMessageManager"),"Acquire",(int)model.ModelID)){
    var messages=((System.Collections.IEnumerable)Call(lease,"SnapshotItems")).Cast<object>().Select(x=>(string)Get(x,"Message")).ToArray();
    Check(messages.Any(x=>x.StartsWith("Saving recovery copy:"))&&messages.Any(x=>x.StartsWith("Recovery copy saved:")),"Scheduled recovery start/completion messages recorded");
   }
   return backup;
  }finally{Call(manager,"Configure",false,10,false);Call(manager,"Forget",(object)model);}
 }
 sealed class ThrowingStream:MemoryStream{public override void Write(byte[] buffer,int offset,int count){throw new IOException("Injected serializer failure");}}
}
