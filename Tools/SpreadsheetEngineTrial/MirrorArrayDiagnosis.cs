using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using DX = DevExpress.Spreadsheet;

internal static partial class EngineBenchmark
{
    static string ArrayReference(int row,int col,int rows=1,int cols=1)
    {return rows==1&&cols==1?Column(col)+(row+1):Address(row,col,rows,cols);}

    static Dictionary<string,string> MirrorArrayDeclarations(DX.Worksheet sheet)
    {
        return sheet.DynamicArrayFormulas.Cast<DX.DynamicArrayFormula>().ToDictionary(
            a=>ArrayReference(a.Range.TopRowIndex,a.Range.LeftColumnIndex),
            a=>ArrayReference(a.Range.TopRowIndex,a.Range.LeftColumnIndex,a.Range.RowCount,a.Range.ColumnCount));
    }

    static void AssertTrialArrayDeclarations(string file,string sheetName,Dictionary<string,string> expected)
    {
        using(var zip=ZipFile.OpenRead(file))
        {
            XNamespace ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var sheet=ReadMirrorXml(zip.GetEntry("xl/workbook.xml")).Descendants(ns+"sheet").Single(s=>(string)s.Attribute("name")==sheetName);
            string id=(string)sheet.Attribute(rel+"id");
            string target=(string)ReadMirrorXml(zip.GetEntry("xl/_rels/workbook.xml.rels")).Root.Elements().Single(r=>(string)r.Attribute("Id")==id).Attribute("Target");
            string part=new Uri(new Uri("http://trial/xl/"),target).AbsolutePath.TrimStart('/');
            var actual=ReadMirrorXml(zip.GetEntry(part)).Descendants(ns+"c").Where(c=>(string)c.Element(ns+"f")?.Attribute("t")=="array")
                .ToDictionary(c=>(string)c.Attribute("r"),c=>(string)c.Element(ns+"f").Attribute("ref"));
            foreach(var entry in expected)
            {
                string range;Check(actual.TryGetValue(entry.Key,out range)&&range==entry.Value,
                    "Candidate rejected: dynamic-array declaration differs at "+sheetName+"!"+entry.Key);
            }
        }
    }

    internal static int RunMirrorExcelArrayOpen(string input,string report)
    {
        input=IoTrial.InTrial(input);report=IoTrial.InTrial(report);Check(!File.Exists(report),"New report required.");
        string hash=IoTrial.FileHash(input);object result;
        using(var reader=new ExcelEngine(false))
        {
            reader.Open(input);result=new {reader.Version,reader.Diagnostics,opened=reader.MirrorArrayOpen()};
        }
        Check(hash==IoTrial.FileHash(input),"Read-only Excel open changed the input.");
        IoTrial.WriteJson(report,new {input,inputHash=hash,sourceUnchanged=true,result,scope="Owned macro-disabled normal read-only Excel open; no calculation, save or VBA execution."});
        Console.WriteLine("EXCEL ARRAY OPEN "+report);return 0;
    }

    internal static int RunMirrorNativeArrayState(string input,string report)
    {
        input=IoTrial.InTrial(input);report=IoTrial.InTrial(report);Check(!File.Exists(report),"New report required.");
        string hash=IoTrial.FileHash(input);
        using(var reader=new DxEngine(true))
        {
            reader.Open(input);
            IoTrial.WriteJson(report,new {input,inputHash=hash,sourceUnchanged=hash==IoTrial.FileHash(input),native=reader.MirrorArrayState(new string[0]),
                scope="Native array interpretation only; no calculation, insertion or save."});
        }
        Console.WriteLine("NATIVE ARRAY STATE "+report);return 0;
    }

    internal static int RunMirrorArrayState(string input,string report)
    {
        input=IoTrial.InTrial(input);report=IoTrial.InTrial(report);Check(!File.Exists(report),"New report required.");
        string hash=IoTrial.FileHash(input);string[] missing;int packaged;
        using(var zip=ZipFile.OpenRead(input))
        {
            XNamespace ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var sheet=ReadMirrorXml(zip.GetEntry("xl/workbook.xml")).Descendants(ns+"sheet").Single(s=>(string)s.Attribute("name")=="Transactional DB");
            string id=(string)sheet.Attribute(rel+"id");
            string target=(string)ReadMirrorXml(zip.GetEntry("xl/_rels/workbook.xml.rels")).Root.Elements().Single(r=>(string)r.Attribute("Id")==id).Attribute("Target");
            string part=new Uri(new Uri("http://trial/xl/"),target).AbsolutePath.TrimStart('/');
            var cells=ReadMirrorXml(zip.GetEntry(part)).Descendants(ns+"c").ToArray();
            packaged=cells.Count(c=>(string)c.Element(ns+"f")?.Attribute("t")=="array");
            missing=cells.Where(c=>c.Attribute("cm")!=null&&c.Element(ns+"f")!=null&&(string)c.Element(ns+"f").Attribute("t")!="array")
                .Select(c=>(string)c.Attribute("r")).ToArray();
        }
        using(var reader=new DxEngine(true))
        {
            reader.Open(input);
            IoTrial.WriteJson(report,new {input,inputHash=hash,sourceUnchanged=hash==IoTrial.FileHash(input),packagedArrayAnchors=packaged,
                metadataFormulaWithoutArray=missing.Length,native=reader.MirrorArrayState(missing),scope="Read-only package/native interpretation; no calculation or workbook save."});
        }
        Console.WriteLine("ARRAY STATE "+report);return 0;
    }
    sealed partial class DxEngine
    {
        internal object MirrorArrayState(string[] missing)
        {
            var sheet=book.Worksheets["Transactional DB"];
            return new {nativeDynamicArrays=sheet.DynamicArrayFormulas.Count,nativeLegacyArrays=sheet.ArrayFormulas.Count,
                metadataOnlyCellsRecognizedAsDynamic=missing.Count(a=>sheet.Cells[a].HasDynamicArrayFormula),
                samples=missing.Take(8).Select(a=>new {address=a,dynamic=sheet.Cells[a].HasDynamicArrayFormula,legacy=sheet.Cells[a].HasArrayFormula}).ToArray()};
        }
    }
    sealed partial class ExcelEngine
    {
        internal object MirrorArrayOpen()
        {
            dynamic sheet=Sheet("Transactional DB"),sheets=null,names=null,xml=null;
            var samples=new System.Collections.Generic.List<object>();
            try
            {
                foreach(string address in new[]{"O1409","Q1409","AA1408","AA1409","AA1448"})
                {
                    dynamic cell=sheet.Range[address];
                    try{samples.Add(new {address,hasSpill=(bool)cell.HasSpill,hasArray=(bool)cell.HasArray,formula=(string)cell.Formula,formula2=(string)cell.Formula2});}
                    finally{Release((object)cell);}
                }
                sheets=book.Worksheets;names=book.Names;xml=book.CustomXMLParts;
                return new {readOnly=(bool)book.ReadOnly,worksheets=(int)sheets.Count,names=(int)names.Count,customXmlParts=(int)xml.Count,hasVba=(bool)book.HasVBProject,samples};
            }
            finally{Release((object)xml);Release((object)names);Release((object)sheets);Release((object)sheet);}
        }
    }
}
