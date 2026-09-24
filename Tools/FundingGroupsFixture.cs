using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
class FundingGroupsFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static int count;
 static void Check(bool ok,string message){if(!ok)throw new Exception(message);count++;Console.WriteLine("PASS "+message);}
 static object Invoke(Type t,string name,params object[] args){return t.GetMethod(name,F).Invoke(null,args);}
 static int Active(Type groups,IWorkbook wb){var map=(IDictionary)Invoke(groups,"TintIndex",wb);return map.Values.Cast<object>().Count(v=>(bool)v.GetType().GetProperty("Active").GetValue(v,null));}
 [STAThread] static int Main(string[] args){try{
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  if(args.Length>4&&args[4]=="inspect")using(var r=new Workbook()){
   r.Options.CalculationMode=WorkbookCalculationMode.Manual;Console.WriteLine("LOAD="+r.LoadDocument(args[2],DocumentFormat.Xlsm)+" sheets="+r.Worksheets.Count);var gt=app.GetType("Abovo.FundingScheduleGroups");
   Console.WriteLine("META "+Invoke(gt,"Read",r));var index=(IDictionary)Invoke(gt,"TintIndex",r);Console.WriteLine("TINTS "+index.Count+" active="+Active(gt,r));
   foreach(var n in r.DefinedNames.Where(n=>n.Name.StartsWith("_SummitSchedule_"))){Console.WriteLine(n.Name+"="+n.RefersTo+" hidden="+n.Hidden);if(n.Range!=null){var c=n.Range[0,0];Console.WriteLine(" CELL "+c.GetReferenceA1()+" type="+c.Value.Type+" date="+c.Value.IsDateTime+" num="+c.Value.IsNumeric+" val="+c.Value+" DateValue="+c.Value.DateTimeValue);}}
   return 0;
  }
  Invoke(app.GetType("Abovo.AbovoAppCls"),"Initialise");var files=app.GetType("Abovo.FileManager");Invoke(files,"Initialise",new object[]{null});
  var copy=Path.Combine(args[1],"private-groups.xlsb");File.Copy(args[2],copy);var open=files.GetMethod("OpenModel");dynamic opened=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});if(opened.BError)throw new Exception(opened.StringReturn);
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)opened.IntegerReturn);IWorkbook wb=model.WB;
  var groups=app.GetType("Abovo.FundingScheduleGroups");var writer=app.GetType("Abovo.FundingScheduleWriter");var type=app.GetType("Abovo.FundingScheduleTarget");
  var targets=((IEnumerable)writer.GetField("Targets").GetValue(null)).Cast<object>().ToArray();
  var selected=Array.CreateInstance(type,2);selected.SetValue(targets[0],0);selected.SetValue(targets[3],1);
  var facilities=((IEnumerable)Invoke(groups,"Facilities",wb,false)).Cast<object>().ToList();dynamic facility=facilities[1];
  if(args.Length>4&&args[4]=="amount"){
   var wsAmount=wb.Worksheets["Funding Assumptions"];var target=targets[3];
   var one=Array.CreateInstance(type,1);one.SetValue(target,0);
   string nr=(string)target.GetType().GetProperty("DateRange").GetValue(target,null);
   var dateRange=wb.DefinedNames.GetDefinedName(nr).Range;
   int row=dateRange.TopRowIndex+((List<int>)Invoke(writer,"BlankRows",wb,target))[0];
   var twoDates=new List<DateTime>{new DateTime(2028,6,30),new DateTime(2028,7,31)};
   Invoke(writer,"Apply",(int)model.ModelID,target,twoDates);wsAmount.Calculate();
   foreach(dynamic f in facilities){var c=wsAmount.Cells[row,(int)f.ColumnIndex];Console.WriteLine("ADMISSION "+c.GetReferenceA1()+" locked="+c.Protection.Locked+" pattern="+c.Fill.PatternType+" format="+c.NumberFormat);}
   object eligible=facilities.FirstOrDefault(f=>{int col=(int)f.GetType().GetProperty("ColumnIndex").GetValue(f,null);var c=wsAmount.Cells[row,col];return !c.Protection.Locked&&!c.HasFormula&&c.Fill.PatternType==PatternType.Solid;});
   Check(eligible!=null,"Demo provides an eligible repayment facility after dates are set");
   var targetValues=wb.DefinedNames.GetDefinedName((string)target.GetType().GetProperty("ValueRange").GetValue(target,null)).Range;
   var blockedByPattern=facilities.First(f=>{int col=(int)f.GetType().GetProperty("ColumnIndex").GetValue(f,null);var c=wsAmount.Cells[row,col];return col<=targetValues.RightColumnIndex&&c.Fill.PatternType!=PatternType.Solid;});
   int amountCol=(int)eligible.GetType().GetProperty("ColumnIndex").GetValue(eligible,null);
   var expectedPattern=wsAmount.Cells[row,amountCol].Fill.PatternType;var expectedColour=wsAmount.Cells[row,amountCol].Fill.BackgroundColor;
   Check(model.ChangeManager.Undo().BSuccess,"Undo admission probe");
   int partCount=wb.CustomXmlParts.Count;
   Invoke(writer,"ApplyBatchWithAmount",(int)model.ModelID,one,twoDates,eligible,Color.Coral,"Fixed figure regression",(decimal?)123.456m);
   Check(wsAmount.Cells[row,amountCol].Value.NumericValue==123.456&&wsAmount.Cells[row+1,amountCol].Value.NumericValue==123.456,"Fixed decimal applied to both selected-loan targets");
   Check(wsAmount.Cells[row,amountCol].Fill.PatternType==expectedPattern&&wsAmount.Cells[row,amountCol].Fill.BackgroundColor==expectedColour,"Physical fill unchanged from normal date entry");
   Check(facilities.Where(f=>(int)f.GetType().GetProperty("ColumnIndex").GetValue(f,null)!=amountCol).All(f=>wsAmount.Cells[row,(int)f.GetType().GetProperty("ColumnIndex").GetValue(f,null)].Value.IsEmpty),"Other loans remain blank");
   Check(model.ChangeManager.Undo().BSuccess&&wsAmount.Cells[row,dateRange.LeftColumnIndex].Value.IsEmpty&&wsAmount.Cells[row,amountCol].Value.IsEmpty&&Active(groups,wb)==0,"One Undo removes figures and dates/tints");
   Check(model.ChangeManager.Redo().BSuccess&&wsAmount.Cells[row,amountCol].Value.NumericValue==123.456&&Active(groups,wb)==4,"One Redo restores figures and dates/tints");
   var path=Path.Combine(args[1],"fixed-figures.xlsb");Check((bool)model.SaveFileAsTo(path,true),"Save scheduled fixed figures to disposable copy");
   using(var reload=new Workbook()){reload.Options.CalculationMode=WorkbookCalculationMode.Manual;Check(reload.LoadDocument(path)&&reload.Worksheets[wsAmount.Name].Cells[row,amountCol].Value.NumericValue==123.456&&Active(groups,reload)==4,"Saved figures and colours reopen together");}
   Check(model.ChangeManager.Undo().BSuccess,"Undo before rejection tests");
   var lockedFacility=blockedByPattern;
   int namesBefore=wb.DefinedNames.Count,partsBefore=wb.CustomXmlParts.Count;bool rejected=false;
   try{Invoke(writer,"ApplyBatchWithAmount",(int)model.ModelID,one,twoDates,lockedFacility,Color.Coral,"locked target",(decimal?)25m);}catch(TargetInvocationException){rejected=true;}
   Check(rejected&&wsAmount.Cells[row,dateRange.LeftColumnIndex].Value.IsEmpty&&wsAmount.Cells[row,amountCol].Value.IsEmpty,"Locked target rejects and rolls back all preceding dates");
   Check(wb.DefinedNames.Count==namesBefore&&wb.CustomXmlParts.Count==partsBefore&&Active(groups,wb)==0,"Rejected schedule leaves no active metadata or new anchors");
   var percentage=targets.First(t=>(bool)Invoke(writer,"IsPercentage",wb,t));
   var mixed=Array.CreateInstance(type,2);mixed.SetValue(target,0);mixed.SetValue(percentage,1);rejected=false;
   try{Invoke(writer,"ApplyBatchWithAmount",(int)model.ModelID,mixed,twoDates,eligible,Color.Coral,"mixed units",(decimal?)5m);}catch(TargetInvocationException){rejected=true;}
   Check(rejected&&wb.DefinedNames.Count==namesBefore,"Mixed amount/percentage request rejected before any mutation");
   bool testedPercent=false;
   foreach(var pt in targets.Where(t=>(bool)Invoke(writer,"IsPercentage",wb,t))){
    var pr=wb.DefinedNames.GetDefinedName((string)pt.GetType().GetProperty("DateRange").GetValue(pt,null)).Range;
    int prow=pr.TopRowIndex+((List<int>)Invoke(writer,"BlankRows",wb,pt))[0];
    Invoke(writer,"Apply",(int)model.ModelID,pt,twoDates);wsAmount.Calculate();
    object pf=facilities.FirstOrDefault(f=>{int col=(int)f.GetType().GetProperty("ColumnIndex").GetValue(f,null);var c=wsAmount.Cells[prow,col];return !c.Protection.Locked&&!c.HasFormula&&c.Fill.PatternType==PatternType.Solid;});
    Check(model.ChangeManager.Undo().BSuccess,"Undo percentage admission probe");
    if(pf==null)continue;
    one.SetValue(pt,0);int pc=(int)pf.GetType().GetProperty("ColumnIndex").GetValue(pf,null);
    Invoke(writer,"ApplyBatchWithAmount",(int)model.ModelID,one,twoDates,pf,Color.Coral,"percentage regression",(decimal?)5.125m);
    Check(Math.Abs(wsAmount.Cells[prow,pc].Value.NumericValue-0.05125)<1e-12,"5.125 percent is stored as 0.05125");
    Check(model.ChangeManager.Undo().BSuccess&&wsAmount.Cells[prow,pc].Value.IsEmpty,"Percentage figure and date undo together");testedPercent=true;break;
   }
   Check(testedPercent,"Percentage section tested on a normally editable facility");
   one.SetValue(target,0);var skipped=new List<string>();int lockedColumn=(int)lockedFacility.GetType().GetProperty("ColumnIndex").GetValue(lockedFacility,null);
   Invoke(writer,"ApplyAvailableAmounts",(int)model.ModelID,one,twoDates,lockedFacility,Color.Coral,"soft warning regression",(decimal?)25m,skipped);
   Check(skipped.Count==2&&!wsAmount.Cells[row,dateRange.LeftColumnIndex].Value.IsEmpty&&wsAmount.Cells[row,lockedColumn].Value.IsEmpty,"Soft policy keeps dates and reports both locked amount targets without writing them");
   Check(model.ChangeManager.Undo().BSuccess&&wsAmount.Cells[row,dateRange.LeftColumnIndex].Value.IsEmpty,"Soft-skip schedule dates undo as one command");
   skipped.Clear();Invoke(writer,"ApplyAvailableAmounts",(int)model.ModelID,one,twoDates,eligible,Color.Coral,"admitted amount regression",(decimal?)123.456m,skipped);
   Check(skipped.Count==0&&wsAmount.Cells[row,amountCol].Value.NumericValue==123.456,"Soft policy still fills all available amount targets");
   Check(model.ChangeManager.Undo().BSuccess&&wsAmount.Cells[row,amountCol].Value.IsEmpty,"Admitted amounts and dates undo together");
   Console.WriteLine("PASS "+count+" assertions");return 0;
  }
  int parts=wb.CustomXmlParts.Count,names=wb.DefinedNames.Count;
  Check(Active(groups,wb)==0&&wb.CustomXmlParts.Count==parts&&wb.DefinedNames.Count==names,"Opening/reading schedule metadata does not migrate the workbook");
  wb.CustomXmlParts.Add("<Sentinel xmlns='urn:abovo:test:unrelated'>preserve</Sentinel>");parts++;
  var ws=wb.Worksheets["Funding Assumptions"];
  var date=new DateTime(2028,4,28);var dates=new List<DateTime>{date,date};
  // A date legitimately changes Excel's conditional fill pattern. Compare
  // coloured scheduling with ordinary typed date entry, not stale initial fills.
  foreach(var t in selected)Invoke(writer,"Apply",(int)model.ModelID,t,dates);
  var expectedFills=ws.GetUsedRange().ExistingCells.ToDictionary(c=>c.GetReferenceA1(),c=>Tuple.Create(c.Fill.PatternType,c.Fill.BackgroundColor,c.Fill.PatternColor));
  Check(model.ChangeManager.Undo().BSuccess&&model.ChangeManager.Undo().BSuccess,"Build and undo uncoloured date-entry reference");
  var before=ws.GetUsedRange().ExistingCells.ToDictionary(c=>c.GetReferenceA1(),c=>Tuple.Create(c.FormulaInvariant,c.Value,c.Protection.Locked,c.Fill.PatternType,c.Fill.BackgroundColor,c.Fill.PatternColor));
  int extra=(int)Invoke(writer,"ApplyBatch",(int)model.ModelID,selected,dates,facility,Color.Coral,"Fixture monthly two occurrences");
  Check(extra==0,"Two sections post without unnecessary row expansion");
  Check(wb.CustomXmlParts.Count==parts+1&&wb.DefinedNames.Count==names+5,"One XML group and five hidden movement anchors added");
  Check(Active(groups,wb)==8,"Both duplicate dates in both sections tint defining date and selected facility only");
  var scheduleNames=wb.DefinedNames.Where(n=>n.Name.StartsWith("_SummitSchedule_")).ToList();Check(scheduleNames.All(n=>n.Hidden),"Schedule anchors hidden from normal Excel Name Manager");
  int changed=0;foreach(var c in ws.GetUsedRange().ExistingCells){Tuple<string,CellValue,bool,PatternType,Color,Color> old;if(!before.TryGetValue(c.GetReferenceA1(),out old))continue;
   var fill=expectedFills[c.GetReferenceA1()];
   CheckCell(c.FormulaInvariant==old.Item1&&c.Protection.Locked==old.Item3&&c.Fill.PatternType==fill.Item1&&c.Fill.BackgroundColor==fill.Item2&&c.Fill.PatternColor==fill.Item3,"Workbook formula/lock/fill differs from normal date entry at "+c.GetReferenceA1());
   if(!c.Value.Equals(old.Item2)){changed++;CheckCell(c.Value.DateTimeValue==date,"Unexpected amount changed");}}
  Check(changed==4,"Only four defining dates changed; formulas/locks preserved and fills exactly match ordinary uncoloured date entry");
  Check(model.ChangeManager.Undo().BSuccess&&Active(groups,wb)==0,"One undo removes both sections' dates and hides their tints");
  Check(model.ChangeManager.Redo().BSuccess&&Active(groups,wb)==8,"One redo restores both sections and tints");
  var invalid=Array.CreateInstance(type,2);invalid.SetValue(targets[0],0);invalid.SetValue(targets.Last(),1);bool refused=false;
  try{Invoke(writer,"ApplyBatch",(int)model.ModelID,invalid,dates,facility,Color.Coral,"invalid");}catch(TargetInvocationException){refused=true;}
  Check(refused&&Active(groups,wb)==8&&wb.DefinedNames.Count==names+5,"Incompatible multi-section request fails before modifying dates or metadata");
  if(args.Length>4&&args[4]=="expand"){
   var reversed=Array.CreateInstance(type,2);reversed.SetValue(targets[3],0);reversed.SetValue(targets[0],1);
   int requested=selected.Cast<object>().Max(t=>((List<int>)Invoke(writer,"BlankRows",wb,t)).Count)+2;
   var oldSizes=selected.Cast<object>().Select(t=>wb.DefinedNames.GetDefinedName((string)t.GetType().GetProperty("DateRange").GetValue(t,null)).Range.RowCount).ToArray();
   int expected=selected.Cast<object>().Sum(t=>requested-((List<int>)Invoke(writer,"BlankRows",wb,t)).Count+5);
   var longer=Enumerable.Range(0,requested).Select(i=>new DateTime(2029,1,1).AddMonths(i)).ToList();
   Check((int)Invoke(writer,"ApplyBatch",(int)model.ModelID,reversed,longer,facility,Color.MediumSeaGreen,"Multi-expansion regression")==expected,"Multi-section expansion adds each shortage plus five");
   Check(selected.Cast<object>().All(t=>((List<int>)Invoke(writer,"BlankRows",wb,t)).Count==5),"Both expanded sections retain five safe spare rows");
   Check(Active(groups,wb)==8+requested*4,"Re-resolved addresses and previous schedule anchors survive lower-then-upper expansion");
   Check(model.ChangeManager.Undo().BSuccess&&Active(groups,wb)==8,"One Undo clears both newly expanded schedules, preserving earlier group");
   Check(model.ChangeManager.Redo().BSuccess&&Active(groups,wb)==8+requested*4,"One Redo restores both expanded schedules and tints");
   Check(!(bool)Invoke(app.GetType("Abovo.ModelSafetyManager"),"IsBulkWorkbookMutationInProgress",(int)model.ModelID),"Bulk recovery-save exclusion is released after the operation");
   files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});Console.WriteLine("PASS "+count+" assertions");return 0;
  }
  string saved=Path.Combine(args[1],"funding-groups.xlsb");Check((bool)model.SaveFileAsTo(saved,true),"Save groups to new XLSB");
  string recovery=Path.Combine(args[1],"funding-groups-recovery.xlsm");using(var output=File.Create(recovery))((object)model).GetType().GetMethod("WriteRecoverySnapshot",F).Invoke(model,new object[]{output});
  Invoke(app.GetType("Abovo.RecoveryXlsmCompatibility"),"Prepare",recovery,((object)model).GetType().GetField("RecoveryHasVerifiedBinaryMetadata",F).GetValue(model));
  foreach(var path in new[]{saved,recovery})using(var r=new Workbook()){
   r.Options.CalculationMode=WorkbookCalculationMode.Manual;Check(r.LoadDocument(path,Path.GetExtension(path)==".xlsm"?DocumentFormat.Xlsm:DocumentFormat.Xlsb)&&Active(groups,r)==8,"Reopen retains active groups: "+Path.GetExtension(path));
   Check(r.CustomXmlParts.Any(p=>p.CustomXmlPartDocument.OuterXml.Contains("urn:abovo:test:unrelated")),"Unrelated XML preserved: "+Path.GetExtension(path));
  }
  {
   var r=wb;var sheet=r.Worksheets["Funding Assumptions"];bool wasProtected=sheet.IsProtected;
   if(wasProtected)Invoke(app.GetType("Abovo.WSSecurity"),"UNProtectWS",(int)model.ModelID,sheet.Name);
   var d=r.DefinedNames.First(n=>n.Name.EndsWith("_Date1"));int row=d.Range.TopRowIndex,col=(int)facility.ColumnIndex;
   sheet.Rows.Insert(row);sheet.Columns.Insert(col);Check(Active(groups,r)==8,"Native inserted row and loan column move every association");
   Check(d.Range.TopRowIndex==row+1,"Date anchor follows inserted row");
   var moved=(IDictionary)Invoke(groups,"TintIndex",r);Check(moved.Contains((row+1)+":"+(col+1)),"Amount tint follows selected loan column");
   sheet.Rows.Remove(row+1);Check(Active(groups,r)==6,"Deleting one scheduled row drops only its own date/amount tint");
   Check((bool)Invoke(groups,"IsPresentationAnchor",r,d),"Deleted owned UI anchor is distinguishable from a broken model name");
   var bad=r.DefinedNames.Add("_SummitSchedule_unowned","=#REF!");bad.Hidden=true;Check(!(bool)Invoke(groups,"IsPresentationAnchor",r,bad),"Unowned broken name is not exempt from integrity");
   if(wasProtected)Invoke(app.GetType("Abovo.WSSecurity"),"ProtectWS",(int)model.ModelID,sheet.Name,Type.Missing);
  }
  files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});Console.WriteLine("PASS "+count+" assertions");return 0;
 }catch(Exception e){Console.WriteLine(e);return 1;}}
 static void CheckCell(bool ok,string message){if(!ok)throw new Exception(message);}
}
