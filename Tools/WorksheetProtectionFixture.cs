using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using DevExpress.Spreadsheet;

// Generated test credentials exist only in this process. Never log or persist them.
public static class WorksheetProtectionFixture {
    static MethodInfo protect;
    static void Check(bool ok,string message) {
        if(!ok) throw new InvalidOperationException(message);
        Console.WriteLine("PASS: "+message);
    }
    static void Protect(IWorkbook book,Worksheet sheet,string password,WorksheetProtectionPermissions permissions) {
        protect.Invoke(null,new object[]{book,sheet,password,permissions});
    }
    static void Release(object value) {
        if(value!=null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
    }
    [STAThread] public static int Main(string[] args) {
        try {
            AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=> {
                var path=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");
                return File.Exists(path)?Assembly.LoadFrom(path):null;
            };
            var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
            protect=app.GetType("Abovo.WSSecurity").GetMethod("ProtectWorksheetForEditing",BindingFlags.NonPublic|BindingFlags.Static);
            Check(protect!=null,"Production worksheet-protection policy found");
            string password=Guid.NewGuid().ToString("N");
            string wrong=Guid.NewGuid().ToString("N");
            string native=Path.Combine(args[1],"native-protection.xlsb");
            string excelCopy=Path.Combine(args[1],"excel-protection-roundtrip.xlsb");
            using(var book=new Workbook()) {
                book.Options.CalculationMode=WorkbookCalculationMode.Manual;
                var sheet=book.Worksheets[0];sheet.Name="Protected";
                sheet.Cells["A1"].Value=7;sheet.Cells["A1"].Protection.Locked=false;
                sheet.Cells["B1"].Formula="=A1*2";sheet.Cells["B1"].Protection.Locked=true;
                sheet.Cells["B1"].NumberFormat="0.00";
                book.DefinedNames.Add("InputValue","'Protected'!$A$1");
                var unprotected=book.Worksheets.Add("Unprotected");
                unprotected.Cells["A1"].Value="Control";
                var options=book.Options.Protection;
                int spinCount=options.SpinCount;
                var custom=WorksheetProtectionPermissions.SelectUnlockedCells;
                foreach(bool setting in new[]{true,false}) {
                    options.UseStrongPasswordVerifier=setting;
                    Protect(book,sheet,password,custom);
                    Check(sheet.IsProtected && sheet.GetProtectionPermissions()==custom,"Explicit worksheet permissions preserved");
                    Check(options.UseStrongPasswordVerifier==setting && options.SpinCount==spinCount,"Verifier policy and spin count restored");
                    try {sheet.Unprotect(wrong);} catch(ArgumentException) { }
                    Check(sheet.IsProtected,"Incorrect password does not remove protection");
                    sheet.Unprotect(password);Check(!sheet.IsProtected,"Same password still unlocks sheet");
                    bool failed=false;
                    try {Protect(book,null,password,custom);} catch(TargetInvocationException) {failed=true;}
                    Check(failed && options.UseStrongPasswordVerifier==setting && options.SpinCount==spinCount,"Policy restored when protection fails");
                }
                // Isolated like-for-like timings, not a full model performance claim.
                options.UseStrongPasswordVerifier=true;
                var watch=Stopwatch.StartNew();
                for(int i=0;i<8;i++){sheet.Protect(password,WorksheetProtectionPermissions.Default);sheet.Unprotect(password);}
                watch.Stop();long strong=watch.ElapsedMilliseconds;watch.Restart();
                for(int i=0;i<8;i++){Protect(book,sheet,password,WorksheetProtectionPermissions.Default);sheet.Unprotect(password);}
                watch.Stop();
                Console.WriteLine("TIMING: 8 protect/unprotect cycles; strong="+strong+" ms, editingPolicy="+watch.ElapsedMilliseconds+" ms, spinCount="+spinCount);
                Protect(book,sheet,password,WorksheetProtectionPermissions.Default);
                Check(!unprotected.IsProtected && !sheet.Cells["A1"].Protection.Locked && sheet.Cells["B1"].Protection.Locked,"Entry protection and cell locks unchanged");
                book.SaveDocument(native,DocumentFormat.Xlsb);
            }
            using(var book=new Workbook()) {
                book.Options.CalculationMode=WorkbookCalculationMode.Manual;book.LoadDocument(native);
                Check(book.Worksheets[0].IsProtected && !book.Worksheets[1].IsProtected,"XLSB retains protected and unprotected sheets");
                book.Worksheets[0].Unprotect(password);
                Check(!book.Worksheets[0].IsProtected,"XLSB reload accepts unchanged password");
            }
            dynamic excel=null,books=null,bookExcel=null,sheetExcel=null,sheets=null,a1=null,b1=null;
            try {
                excel=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excel.Visible=false;excel.DisplayAlerts=false;excel.EnableEvents=false;excel.AutomationSecurity=3;excel.AskToUpdateLinks=false;
                books=excel.Workbooks;bookExcel=books.Open(native,0,true);sheets=bookExcel.Worksheets;sheetExcel=sheets["Protected"];
                Check((bool)sheetExcel.ProtectContents,"Microsoft Excel recognises Summit worksheet protection");
                a1=sheetExcel.Range["A1"];b1=sheetExcel.Range["B1"];
                Check(!(bool)a1.Locked && (bool)b1.Locked,"Excel retains editable inputs and locked formulas");
                sheetExcel.Unprotect(password);
                Check(!(bool)sheetExcel.ProtectContents,"Microsoft Excel unlocks with the same password");
                a1.Value2=8;
                sheetExcel.Protect(password);
                Check((bool)sheetExcel.ProtectContents,"Excel can re-protect the Summit-produced sheet");
                bookExcel.SaveCopyAs(excelCopy);
            } finally {
                try {
                    Release((object)b1);Release((object)a1);Release((object)sheetExcel);Release((object)sheets);
                    if((object)bookExcel!=null){bookExcel.Close(false);Release((object)bookExcel);}
                    Release((object)books);
                } finally {
                    if((object)excel!=null){excel.Quit();Release((object)excel);}
                }
            }
            using(var book=new Workbook()) {
                book.Options.CalculationMode=WorkbookCalculationMode.Manual;book.LoadDocument(excelCopy);
                var sheet=book.Worksheets["Protected"];
                Check(sheet.IsProtected && !book.Worksheets["Unprotected"].IsProtected,"Excel round-trip retains entry protection state");
                sheet.Unprotect(password);Check(!sheet.IsProtected,"Summit reads protection re-applied by Excel");
                Check(sheet.Cells["A1"].Value.NumericValue==8 && sheet.Cells["B1"].FormulaInvariant=="=A1*2","Round-trip retains input edits and formulas");
                var named=book.DefinedNames.GetDefinedName("InputValue").Range;
                Check(named.Worksheet==sheet && named.TopRowIndex==0 && named.LeftColumnIndex==0 && named.RowCount==1 && named.ColumnCount==1,"Round-trip retains defined name");
                Check(!sheet.Cells["A1"].Protection.Locked && sheet.Cells["B1"].Protection.Locked && sheet.Cells["B1"].NumberFormat=="0.00","Round-trip retains cell locks and number format");
            }
            Console.WriteLine("PASS: Worksheet protection round-trip. No source workbook used or changed.");
            return 0;
        } catch(Exception error) {Console.Error.WriteLine(error);return 1;}
    }
}
