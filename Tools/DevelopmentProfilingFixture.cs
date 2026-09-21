using System;
using System.IO;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.XtraGrid.Views.Grid;
using System.Text.RegularExpressions;

public static class DevelopmentProfilingFixture {
    [STAThread] public static int Main(string[] args) {
        try {
            AppDomain.CurrentDomain.AssemblyResolve += (s,e) => {
                string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");
                return File.Exists(p)?Assembly.LoadFrom(p):null;
            };
            Application.EnableVisualStyles();
            var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
            app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
            var files=app.GetType("Abovo.FileManager");
            files.GetMethod("Initialise").Invoke(null,new object[]{null});
            var open=files.GetMethod("OpenModel");
            dynamic result=open.Invoke(null,new object[]{args[1],new FileInfo(args[1]),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
            if(result.BError)throw new Exception("Open failed");
            dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);
            IWorkbook book=(IWorkbook)model.WB;
            foreach(string name in new[]{"HouseTypeInID","Rep_DevBP_12a","Rep_DevBP_13","Rep_DevBP_16"}) {
                var range=book.DefinedNames.GetDefinedName(name).Range;
                int values=0, formulas=0;
                for(int r=0;r<range.RowCount;r++)for(int c=0;c<range.ColumnCount;c++) {
                    if(!range[r,c].Value.IsEmpty) {
                        values++;
                        if(name!="HouseTypeInID" && values<=5) Console.WriteLine("  Source "+range[r,c].GetReferenceA1()+" = "+range[r,c].DisplayText);
                    }
                    if(!String.IsNullOrEmpty(range[r,c].Formula)) formulas++;
                }
                Console.WriteLine(name+": "+range.GetReferenceA1()+", nonempty="+values+", formulas="+formulas);
                if(name=="Rep_DevBP_12a") foreach(int r in new[]{0,20,21,22,58,59,60,83,84,94}) {
                    string line="Axis row "+(range.TopRowIndex+r+1)+":";
                    for(int c=0;c<6;c++)line+=" | "+range.Worksheet.Cells[range.TopRowIndex+r,c].DisplayText;
                    Console.WriteLine(line);
                }
            }
            using(Form group=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{0,0,"Normal"})) {
                group.Opacity=0; group.ShowInTaskbar=false; group.ClientSize=new Size(1800,1000);
                group.Show(); Application.DoEvents();
                Call(group,"ShowInterface",0,28,false,"None",null,-1);
                Control dit=(Control)Field(group,"ActiveInterface");
                for(int i=1;i<3;i++) Call(dit,"BuildSection",i,false,false);
                Application.DoEvents();
                int checkedViews=0;
                foreach(GridView view in (Array)Field(dit,"UsedGridVIEWS")) {
                    if(view==null || view.GridControl==null || view.Columns.Count==0)continue;
                    view.GridControl.ForceInitialize();
                    checkedViews++;
                    Console.WriteLine("GRID: rows="+view.DataRowCount+", cols="+view.Columns.Count);
                    for(int r=0;r<Math.Min(3,view.DataRowCount);r++)
                        Console.WriteLine("Row "+r+": "+view.GetRowCellDisplayText(r,view.Columns[0])+" | "+view.GetRowCellDisplayText(r,view.Columns[1]));
                    var range=book.DefinedNames.GetDefinedName("HouseTypeInID").Range;
                    for(int r=0;r<view.DataRowCount;r++)
                        if(view.GetRowCellDisplayText(r,view.Columns[0])!=range[0,r].DisplayText)throw new Exception("Grid description differs from workbook");
                    if(view.Columns[0].Caption.Trim()!="House Type")throw new Exception("Description caption");
                    dynamic data=Field(view.Tag,"DataSet");
                    string profileName=((string)data.Name).EndsWith("-IEDS-UnitProfiling")?"Rep_DevBP_12a":((string)data.Name).EndsWith("-IEDS-BOCProfiling")?"Rep_DevBP_13":"Rep_DevBP_16";
                    var profile=book.DefinedNames.GetDefinedName(profileName).Range;
                    if(view.Columns.Count!=profile.RowCount+1 || view.DataRowCount!=profile.ColumnCount-1)throw new Exception("Incomplete profiling dimensions");
                    for(int c=1;c<view.Columns.Count;c++) {
                        int sr=profile.TopRowIndex+c-1;
                        string period=profile.Worksheet.Cells[sr,2].DisplayText;
                        string year=profile.Worksheet.Cells[sr,3].DisplayText;
                        string expected=String.IsNullOrWhiteSpace(period)?"Yr "+year:"Per "+period+" Yr "+year;
                        if(Regex.Replace(view.Columns[c].Caption,@"\s+"," ").Trim()!=expected)throw new Exception("Heading drift: "+profileName+" field "+c+" expected "+expected);
                        for(int r=0;r<view.DataRowCount;r++) {
                            var source=profile[c-1,r];
                            dynamic point=data.DataRows[r].DataCells[c];
                            if((string)point.SourceSheet!=source.Worksheet.Name || ((string)point.SourceAddress).Replace("$","")!=source.GetReferenceA1().Replace("$",""))throw new Exception("Source address drift");
                            object value=view.GetRowCellValue(r,view.Columns[c]);
                            if(String.IsNullOrEmpty(source.DisplayText)) {
                                if(value!=null && value!=DBNull.Value && Convert.ToString(value)!="")throw new Exception("Blank manufactured as a value");
                            } else if(Convert.ToInt32(value)!=Convert.ToInt32(source.Value.NumericValue))throw new Exception("Source value differs");
                        }
                    }
                    for(int r=0;r<view.DataRowCount;r++) for(int c=1;c<view.Columns.Count;c++) {
                        var value=view.GetRowCellValue(r,view.Columns[c]);
                        if(value!=null && value!=DBNull.Value && Convert.ToString(value)!="")
                            Console.WriteLine("  Grid r"+r+" c"+c+" "+view.Columns[c].Caption.Replace("\n"," ")+" = "+value);
                    }
                }
                if(checkedViews!=3)throw new Exception("Expected three profiling grids");
                using(var image=new Bitmap(dit.Width,dit.Height)) {
                    dit.DrawToBitmap(image,new Rectangle(Point.Empty,dit.Size));
                    image.Save(Path.Combine(args[2],"profiling.png"));
                }
            }
            Console.WriteLine("PASS: all three native grids match descriptions, all 95 mixed period headings, cell addresses and values; no save performed.");
            return 0;
        } catch(Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
    static object Field(object obj,string name) { return obj.GetType().GetField(name,BindingFlags.Instance|BindingFlags.Public|BindingFlags.NonPublic).GetValue(obj); }
    static object Call(object obj,string name,params object[] args) { return obj.GetType().GetMethod(name,BindingFlags.Instance|BindingFlags.Public|BindingFlags.NonPublic).Invoke(obj,args); }
}
