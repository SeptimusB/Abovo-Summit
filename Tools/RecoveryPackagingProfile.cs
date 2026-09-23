// Isolated package-stage benchmark. Supplied workbooks are never saved.
using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Linq;
using System.Reflection;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Collections.Generic;
using DevExpress.Spreadsheet;

public static class RecoveryPackagingProfile {
 const BindingFlags F=BindingFlags.Static|BindingFlags.NonPublic|BindingFlags.Public;
 static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("PASS: "+text);}
 static string Hash(Stream stream){using(var h=SHA256.Create())return BitConverter.ToString(h.ComputeHash(stream));}
 static Dictionary<string,string> Parts(string path){using(var zip=ZipFile.OpenRead(path))return zip.Entries.ToDictionary(e=>e.FullName,e=>{using(var s=e.Open())return Hash(s);});}
 static string FileHash(string path){using(var s=File.OpenRead(path))return Hash(s);}
 static void Part(ZipArchive zip,string name,string text){using(var s=new StreamWriter(zip.CreateEntry(name).Open()))s.Write(text);}
 static void MetadataCases(string output,MethodInfo prepare){
  string ss="http://schemas.openxmlformats.org/spreadsheetml/2006/main",rel="http://schemas.openxmlformats.org/package/2006/relationships",ct="http://schemas.openxmlformats.org/package/2006/content-types";
  foreach(string test in new[]{"no-metadata","existing-metadata","unknown-cm","unknown-vm","unverified-source","dangling-relation","dangling-type","malformed-sheet","dtd-sheet","missing-relations"}){
   string path=Path.Combine(output,"edge-"+test+".xlsm");
   using(var zip=ZipFile.Open(path,ZipArchiveMode.Create)){
    string attributes=test=="no-metadata"?"":test=="unknown-cm"?"cm=\"2\"":test=="unknown-vm"?"cm=\"1\" vm=\"1\"":"cm=\"1\"";
    Part(zip,"xl/worksheets/sheet1.xml",test=="malformed-sheet"?"<worksheet":(test=="dtd-sheet"?"<!DOCTYPE worksheet [<!ENTITY x '1'>]>":"")+"<worksheet xmlns=\""+ss+"\"><sheetData><row r=\"1\"><c r=\"A1\" "+attributes+"><v>1</v></c></row></sheetData></worksheet>");
    if(test!="missing-relations")Part(zip,"xl/_rels/workbook.xml.rels","<Relationships xmlns=\""+rel+"\">"+(test=="dangling-relation"?"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata\" Target=\"metadata.xml\"/>":"")+"</Relationships>");
    Part(zip,"[Content_Types].xml","<Types xmlns=\""+ct+"\">"+(test=="dangling-type"?"<Override PartName=\"/xl/metadata.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml\"/>":"")+"</Types>");
    if(test=="existing-metadata")Part(zip,"xl/metadata.xml","<metadata xmlns=\""+ss+"\"/>");
   }
   string before=FileHash(path);bool rejected=false;
   try{prepare.Invoke(null,new object[]{path,test!="unverified-source"});}catch(TargetInvocationException e){if(!(e.InnerException is InvalidDataException)&&!(e.InnerException is XmlException)&&!(e.InnerException is NullReferenceException))throw;rejected=true;}
   bool noop=test=="no-metadata"||test=="existing-metadata";
   Check(rejected!=noop,test+": expected no-op/rejection");
   Check(FileHash(path)==before,test+": exact file bytes preserved");
   using(var unlocked=new FileStream(path,FileMode.Open,FileAccess.ReadWrite,FileShare.None))Check(unlocked.CanWrite,test+": exclusive handle released");
  }
 }
 static long Scan(string path,ZipArchiveMode mode){long count=0;using(var file=new FileStream(path,FileMode.Open,mode==ZipArchiveMode.Read?FileAccess.Read:FileAccess.ReadWrite,FileShare.None))using(var zip=new ZipArchive(file,mode)){
  foreach(var part in zip.Entries.Where(e=>e.FullName.StartsWith("xl/worksheets/",StringComparison.Ordinal)&&e.FullName.EndsWith(".xml",StringComparison.OrdinalIgnoreCase))){
   using(var s=part.Open())using(var reader=XmlReader.Create(s,new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=536870912}))while(reader.Read()){
    if(reader.NodeType!=XmlNodeType.Element||reader.LocalName!="c"||reader.NamespaceURI!="http://schemas.openxmlformats.org/spreadsheetml/2006/main")continue;
    var cm=reader.GetAttribute("cm");if(reader.GetAttribute("vm")!=null||(cm!=null&&cm!="1"))throw new InvalidDataException("Unsupported profile");if(cm!=null)count++;
   }
  }
 }return count;}
 [STAThread]public static int Main(string[] args){try{
  string bin=args[0],output=args[1],source=args[2];
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(bin,new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  string raw=Path.Combine(output,"raw-export.xlsm");var timer=new Stopwatch();
  if(args[3]=="--raw")File.Copy(source,raw);else{
   using(var w=new Workbook()){
    w.Options.CalculationMode=WorkbookCalculationMode.Manual;w.Options.Import.ThrowExceptionOnInvalidDocument=true;
    timer.Start();Check(w.LoadDocument(source),"Load supplied workbook without saving original");Console.WriteLine("OPEN_MS="+timer.ElapsedMilliseconds);
    w.Options.CalculationMode=WorkbookCalculationMode.Manual;timer.Restart();w.SaveDocument(raw,DocumentFormat.Xlsm);Console.WriteLine("EXPORT_MS="+timer.ElapsedMilliseconds);
   }
   GC.Collect();GC.WaitForPendingFinalizers();
  }
  var original=Parts(raw);Check(!original.ContainsKey("xl/metadata.xml"),"Raw export exercises missing-metadata bridge");
  var app=Assembly.LoadFrom(Path.Combine(bin,"Abovo-summit.exe"));var prepare=app.GetType("Abovo.RecoveryXlsmCompatibility").GetMethod("Prepare",F);
  MetadataCases(output,prepare);
  for(int run=0;run<3;run++){
   foreach(var mode in (run%2==0?new[]{ZipArchiveMode.Update,ZipArchiveMode.Read}:new[]{ZipArchiveMode.Read,ZipArchiveMode.Update})){
    string copy=Path.Combine(output,"scan-"+run+"-"+mode+".xlsm");File.Copy(raw,copy);timer.Restart();long cells=Scan(copy,mode);long ms=timer.ElapsedMilliseconds;
    Console.WriteLine("SCAN run="+run+" mode="+mode+" ms="+ms+" metadataCells="+cells+" bytes="+new FileInfo(copy).Length);
    var after=Parts(copy);Check(after.Count==original.Count&&original.All(p=>after[p.Key]==p.Value),"Scan preserves all uncompressed package parts");
   }
   string prepared=Path.Combine(output,"prepared-"+run+".xlsm");File.Copy(raw,prepared);timer.Restart();prepare.Invoke(null,new object[]{prepared,true});long prepareMs=timer.ElapsedMilliseconds;
   Console.WriteLine("PREPARE run="+run+" ms="+prepareMs+" bytes="+new FileInfo(prepared).Length);
   var actual=Parts(prepared);Check(actual.Count==original.Count+1&&actual.ContainsKey("xl/metadata.xml")&&original.Where(p=>p.Key!="[Content_Types].xml"&&p.Key!="xl/_rels/workbook.xml.rels").All(p=>actual[p.Key]==p.Value),"Bridge preserves every original worksheet, VBA and other payload part");
   string preparedHash=FileHash(prepared);prepare.Invoke(null,new object[]{prepared,true});Check(FileHash(prepared)==preparedHash,"Repeated preparation is an exact-byte no-op");
  }
  return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
}
