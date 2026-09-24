param([string]$Document, [string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo 'bin/Release'
$out=(Resolve-Path -LiteralPath $OutputDirectory).Path
$source=(Resolve-Path -LiteralPath $Document).Path
$hash=(Get-FileHash -LiteralPath $source).Hash
$code=@'
using System;
using System.IO;
using System.Reflection;
class RenderReview {
 [STAThread] static int Main(string[] a) {
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(a[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  return Render(a);
 }
 static int Render(string[] a) {
  using(var doc=new DevExpress.XtraRichEdit.RichEditDocumentServer()) {
   doc.LoadDocument(a[1],DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
   doc.ExportToPdf(a[2]);
  }
  Console.WriteLine("Native PDF render complete");return 0;
 }
}
'@
$compiler=New-Object System.CodeDom.Compiler.CompilerParameters
$compiler.GenerateExecutable=$true
$compiler.OutputAssembly=Join-Path $out 'RenderReview.exe'
$compiler.CompilerOptions='/platform:x86'
foreach($ref in @('System.dll','System.Core.dll','System.Drawing.dll')){[void]$compiler.ReferencedAssemblies.Add($ref)}
foreach($ref in @('DevExpress.RichEdit.v25.2.Core.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll')){[void]$compiler.ReferencedAssemblies.Add((Join-Path $bin $ref))}
$provider=New-Object Microsoft.CSharp.CSharpCodeProvider
try{$compiled=$provider.CompileAssemblyFromSource($compiler,$code);if($compiled.Errors.HasErrors){throw ($compiled.Errors | Out-String)}}finally{$provider.Dispose()}
& $compiler.OutputAssembly $bin $source (Join-Path $out 'review.pdf')
if($LASTEXITCODE -ne 0){throw 'Native render failed'}
if((Get-FileHash -LiteralPath $source).Hash -ne $hash){throw 'Report changed during native render'}
