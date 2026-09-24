// Actual-model regression. All arrangements use a private copy; Save As writes a new test artifact only.
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.Spreadsheet;

class FundingContiguous297Fixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static int checks; static Assembly app; static Type writer,groups,targetType; static dynamic model; static IWorkbook wb;
 static void Check(bool value,string message){if(!value)throw new Exception(message);Console.WriteLine("PASS "+(++checks)+" "+message);}
 static object Invoke(Type t,string name,params object[] args){return t.GetMethod(name,F).Invoke(null,args);}
 static string Name(object t){return (string)t.GetType().GetProperty("DateRange").GetValue(t,null);}
 static Tuple<CellRange,CellRange> Ranges(object t){return (Tuple<CellRange,CellRange>)Invoke(writer,"Ranges",wb,t);}
 static List<int> Blank(object t){return (List<int>)Invoke(writer,"BlankRows",wb,t);}
 static dynamic Plan(object t,int count){return Invoke(writer,"Placement",wb,t,count);}
 static Array Targets(params object[] values){var array=Array.CreateInstance(targetType,values.Length);for(int i=0;i<values.Length;i++)array.SetValue(values[i],i);return array;}
 static int Active(){return ((IDictionary)Invoke(groups,"TintIndex",wb)).Values.Cast<object>().Count(v=>(bool)v.GetType().GetProperty("Active").GetValue(v,null));}
 static long Revision(){var type=((object)model).GetType();var field=type.GetField("UserChangeRevision",F);return Convert.ToInt64(field!=null?field.GetValue(model):type.GetProperty("UserChangeRevision",F).GetValue(model,null));}
 static void Arrange(Action action){
  var sheet=wb.Worksheets["Funding Assumptions"];bool protection=sheet.IsProtected;wb.BeginUpdate();
  try{if(protection)Invoke(app.GetType("Abovo.WSSecurity"),"UNProtectWS",(int)model.ModelID,sheet.Name);action();}
  finally{if(protection)Invoke(app.GetType("Abovo.WSSecurity"),"ProtectWS",(int)model.ModelID,sheet.Name,Type.Missing);wb.EndUpdate();}
 }
 static List<int> Occupy(object target){
  var pair=Ranges(target);
  Arrange(()=>{
   for(int i=0;i<pair.Item1.RowCount;i++){
    var date=pair.Item1[i,0];if(!date.HasFormula&&!date.Protection.Locked)date.ClearContents();
    for(int j=0;j<pair.Item2.ColumnCount;j++){var cell=pair.Item2[i,j];if(!cell.HasFormula)cell.ClearContents();}
   }
  });
  var candidates=Blank(target);
  Arrange(()=>{foreach(int offset in candidates)pair.Item1[offset,0].Value=new DateTime(2030,1,1).AddDays(offset);});
  return candidates;
 }
 static void Free(object target,IEnumerable<int> offsets){var range=Ranges(target).Item1;Arrange(()=>{foreach(int offset in offsets)range[offset,0].ClearContents();});}
 static int FirstRun(List<int> rows,int needed){for(int i=0;i<=rows.Count-needed;i++)if(rows[i+needed-1]==rows[i]+needed-1)return rows[i];throw new Exception("Private fixture lacks a sufficiently long input area");}
 static string InputState(object target){var p=Ranges(target);return String.Join("|",p.Item1.ExistingCells.Concat(p.Item2.ExistingCells).Select(c=>c.GetReferenceA1()+":"+c.FormulaInvariant+":"+c.Value+":"+c.Protection.Locked).ToArray());}
 static string StructureState(object target){var p=Ranges(target);return String.Join("|",p.Item1.ExistingCells.Concat(p.Item2.ExistingCells).Select(c=>c.GetReferenceA1()+":"+c.FormulaInvariant+":"+c.Protection.Locked).ToArray());}
 static dynamic Apply(object[] sections,List<DateTime> dates,object facility){return Invoke(writer,"ApplyAvailableAmountsWithLocations",(int)model.ModelID,Targets(sections),dates,facility,Color.Coral,"Contiguous regression",null,new List<string>());}
 static List<DateTime> Dates(int count){return Enumerable.Range(0,count).Select(i=>new DateTime(2040,1,24).AddMonths(i/2)).ToList();}
 static void AssertDates(object target,int offset,List<DateTime> dates){var r=Ranges(target).Item1;Check(dates.Select((d,i)=>r[offset+i,0].Value.DateTimeValue==d).All(x=>x),Name(target)+" dates occupy one consecutive block, including duplicates");}
 [STAThread]static int Main(string[] args){try{
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));Invoke(app.GetType("Abovo.AbovoAppCls"),"Initialise");
  var files=app.GetType("Abovo.FileManager");Invoke(files,"Initialise",new object[]{null});
  string copy=Path.Combine(args[1],"private-contiguous.xlsb");File.Copy(args[2],copy);var open=files.GetMethod("OpenModel");
  dynamic opened=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});Check(!opened.BError,"Private Demo opens");
  model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)opened.IntegerReturn);wb=model.WB;
  writer=app.GetType("Abovo.FundingScheduleWriter");groups=app.GetType("Abovo.FundingScheduleGroups");targetType=app.GetType("Abovo.FundingScheduleTarget");
  var targets=((IEnumerable)writer.GetField("Targets").GetValue(null)).Cast<object>().ToArray();
  object repay=targets.First(t=>Name(t)=="IR_Fund_Repay"),increase=targets.First(t=>Name(t)=="IR_Fund_Fac_Inc");
  object facility=((IEnumerable)Invoke(groups,"Facilities",wb,false)).Cast<object>().First();
  int facilityColumn=(int)facility.GetType().GetProperty("ColumnIndex").GetValue(facility,null);
  bool protectedBefore=wb.Worksheets["Funding Assumptions"].IsProtected;

  // Earlier holes do not scatter a schedule. Zero and formula amounts in any
  // other loan column still reserve the corresponding shared defining row.
  var candidates=Occupy(repay);int start=FirstRun(candidates,10);
  Free(repay,new[]{start,start+1,start+3,start+4,start+5,start+7,start+8,start+9});
  dynamic plan=Plan(repay,3);Check(plan.StartOffset==start+3&&plan.RowsToAdd==0,"First-fit uses earliest exact three-row block, not two earlier holes");
  var pair=Ranges(repay);var orphan=pair.Item2[start+3,pair.Item2.ColumnCount-1];
  Arrange(()=>orphan.Value=0);plan=Plan(repay,3);Check(plan.StartOffset==start+7&&plan.RowsToAdd==0,"Zero amount in another loan reserves its date row");
  Arrange(()=>orphan.FormulaInvariant="=0");plan=Plan(repay,3);Check(plan.StartOffset==start+7,"Formula-owned amount reserves its date row");
  Arrange(()=>orphan.ClearContents());
  var guard=pair.Item1[start+3,0];bool oldLock=guard.Protection.Locked;var oldPattern=guard.Fill.PatternType;
  Arrange(()=>guard.Protection.Locked=true);Check((int)Plan(repay,3).StartOffset==start+7,"Locked defining dates cannot join a block");
  Arrange(()=>{guard.Protection.Locked=oldLock;guard.Fill.PatternType=PatternType.DarkGray;});Check((int)Plan(repay,3).StartOffset==start+7,"Fill-restricted defining dates cannot join a block");
  Arrange(()=>guard.Fill.PatternType=oldPattern);
  string original=InputState(repay),structure=StructureState(repay);int oldCount=pair.Item1.RowCount;
  var three=Dates(3);dynamic result=Apply(new[]{repay},three,facility);
  Check(result.AddedRows==0&&Ranges(repay).Item1.RowCount==oldCount,"Exact-fit placement does not expand");
  AssertDates(repay,start+3,three);
  Check(result.FirstDateRows[Name(repay)]==pair.Item1.TopRowIndex+start+3,"Writer reports the actual first target row for UI focus");
  Check(pair.Item1[start,0].Value.IsEmpty&&pair.Item1[start+1,0].Value.IsEmpty&&pair.Item1[start+7,0].Value.IsEmpty,"Unused isolated holes and later runs remain untouched");
  Check(StructureState(repay)==structure&&Active()==6,"Formulas/locks preserved; tints cover dates and selected facility only");
  Check(model.ChangeManager.Undo().BSuccess&&InputState(repay)==original&&Active()==0,"One Undo restores exact pre-schedule inputs and removes active tints");
  Check(model.ChangeManager.Redo().BSuccess&&Active()==6,"One Redo restores the contiguous schedule and tints");
  AssertDates(repay,start+3,three);Check(model.ChangeManager.Undo().BSuccess,"Reset first-fit test through normal Undo");

  // The total number of scattered blanks is deliberately sufficient, but no
  // individual run is. Extend the terminal run, leaving five blank rows after it.
  candidates=Occupy(repay);pair=Ranges(repay);oldCount=pair.Item1.RowCount;start=FirstRun(candidates,7);
  Check(candidates.Contains(oldCount-1)&&candidates.Contains(oldCount-2),"Master provides two editable terminal template rows");
  Free(repay,new[]{start,start+2,start+4,oldCount-2,oldCount-1}.Distinct());
  var four=Dates(4);plan=Plan(repay,4);Check(Blank(repay).Count>=4&&plan.StartOffset==oldCount-2&&plan.RowsToAdd==7,"Insufficient contiguous capacity extends two-row tail by shortage plus five");
  int oldBottom=pair.Item1.BottomRowIndex;var ws=pair.Item1.Worksheet;
  var constants=ws.GetUsedRange().ExistingCells.Where(c=>!c.HasFormula&&!c.Value.IsEmpty).Select(c=>Tuple.Create(c.RowIndex,c.ColumnIndex,c.Value)).ToArray();
  result=Apply(new[]{repay},four,facility);Check(result.AddedRows==7&&Ranges(repay).Item1.RowCount==oldCount+7,"Actual expansion matches the placement plan");
  AssertDates(repay,oldCount-2,four);
  Check(Enumerable.Range(oldCount+2,5).All(i=>Blank(repay).Contains(i)),"Exactly the five new spare tail rows remain safe and blank");
  Check(constants.All(c=>ws.Cells[c.Item1>oldBottom?c.Item1+7:c.Item1,c.Item2].Value.Equals(c.Item3)),"All pre-existing Funding constants survive the physical insertion");
  Check(model.ChangeManager.Undo().BSuccess&&Ranges(repay).Item1.RowCount==oldCount+7&&Active()==0,"Undo removes schedule but retains added blank capacity, as documented");

  // Lower-then-upper insertions must re-resolve both named ranges and return
  // final worksheet addresses, not the addresses before the upper insertion.
  candidates=Occupy(repay);int repayCount=Ranges(repay).Item1.RowCount;Free(repay,new[]{repayCount-1});
  candidates=Occupy(increase);int increaseCount=Ranges(increase).Item1.RowCount;Free(increase,new[]{increaseCount-2,increaseCount-1});
  int repayTop=Ranges(repay).Item1.TopRowIndex;result=Apply(new[]{repay,increase},four,facility);
  Check(result.AddedRows==15,"Two selected sections expand by their own contiguous shortages plus five");
  AssertDates(repay,repayCount-1,four);AssertDates(increase,increaseCount-2,four);
  Check(result.FirstDateRows[Name(repay)]==repayTop+7+repayCount-1&&result.FirstDateRows[Name(increase)]==Ranges(increase).Item1.TopRowIndex+increaseCount-2,"Returned focus locations follow lower-then-upper structural shifts");
  Check(Active()==16,"Both sections tint all eight distinct scheduled rows");
  Check(model.ChangeManager.Undo().BSuccess&&Active()==0,"One Undo removes both schedules together");
  Check(model.ChangeManager.Redo().BSuccess&&Active()==16,"One Redo restores both schedules together");
  Check(model.ChangeManager.Undo().BSuccess,"Reset multi-section schedule through normal Undo");

  // An amount rejection occurs after dates are admitted: date values and XML
  // anchors must roll back, without changing the selected cell's lock policy.
  candidates=Occupy(repay);start=FirstRun(candidates,4);Free(repay,Enumerable.Range(start,4));pair=Ranges(repay);
  var blocked=pair.Item1.Worksheet.Cells[pair.Item1.TopRowIndex+start,facilityColumn];bool blockedLock=blocked.Protection.Locked;
  Arrange(()=>blocked.Protection.Locked=true);original=InputState(repay);int names=wb.DefinedNames.Count,parts=wb.CustomXmlParts.Count;long revision=Revision();
  bool rejected=false;try{Invoke(writer,"ApplyBatchWithAmount",(int)model.ModelID,Targets(repay),four,facility,Color.Coral,"Rejected contiguous figures",(decimal?)25m);}catch(TargetInvocationException){rejected=true;}
  Check(rejected&&InputState(repay)==original,"Failed amount write restores every preceding date and existing input");
  Check(wb.DefinedNames.Count==names&&wb.CustomXmlParts.Count==parts&&Active()==0,"Failed schedule restores metadata and hidden anchors");
  Check(Revision()==revision&&blocked.Protection.Locked,"Failed schedule does not log a user revision or bypass the lock");
  Arrange(()=>blocked.Protection.Locked=blockedLock);
  Check(wb.Worksheets["Funding Assumptions"].IsProtected==protectedBefore,"Entry worksheet protection restored across all operations");
  Check(!(bool)Invoke(app.GetType("Abovo.ModelSafetyManager"),"IsBulkWorkbookMutationInProgress",(int)model.ModelID),"Bulk mutation gate released after rejection");
  result=Apply(new[]{repay},four,facility);int savedOffset=(int)result.FirstDateRows[Name(repay)]-Ranges(repay).Item1.TopRowIndex;
  string saved=Path.Combine(args[1],"contiguous-schedule.xlsb");Check((bool)model.SaveFileAsTo(saved,true),"Save contiguous schedule to a new disposable artifact");
  using(var reload=new Workbook()){
   reload.Options.CalculationMode=WorkbookCalculationMode.Manual;Check(reload.LoadDocument(saved),"Native reload opens scheduled artifact");
   var savedDates=reload.DefinedNames.GetDefinedName(Name(repay)).Range;
   Check(four.Select((d,i)=>savedDates[savedOffset+i,0].Value.DateTimeValue==d).All(x=>x),"Reload retains contiguous duplicate-date block");
   Check(((IDictionary)Invoke(groups,"TintIndex",reload)).Values.Cast<object>().Count(v=>(bool)v.GetType().GetProperty("Active").GetValue(v,null))==8,"Reload retains exactly the active date/facility tints");
  }
  model.ModelSpreadsheetControl.Dispose();Console.WriteLine("PASS "+checks+" assertions; original workbook unchanged; save artifact="+saved);return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
}
