param([Parameter(Mandatory=$true)][string]$Fixture,[string]$Configuration='Debug',[string]$Workbook='Library/Demo BP v26_0001.xlsb',[string]$Mode='',[switch]$UseApplicationConfig,[switch]$MatchApplicationAddressSpace)
$ErrorActionPreference='Stop'
function Get-LargeAddressAware([string]$Image) {
    $imageStream=[IO.File]::OpenRead($Image)
    $imageReader=[IO.BinaryReader]::new($imageStream)
    try {
        if($imageReader.ReadUInt16() -ne 0x5A4D){throw "Not a PE executable: $Image"}
        $imageStream.Position=0x3c
        $headerOffset=$imageReader.ReadInt32()
        $imageStream.Position=$headerOffset
        if($imageReader.ReadUInt32() -ne 0x4550){throw "Invalid PE signature: $Image"}
        $imageStream.Position=$headerOffset+22
        return (($imageReader.ReadUInt16() -band 0x20) -ne 0)
    } finally {$imageReader.Dispose();$imageStream.Dispose()}
}
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo "bin/$Configuration"
$out=New-Item -ItemType Directory -Path (Join-Path $repo ('obj/ClientReportTests/'+[guid]::NewGuid().ToString('N')))
$source=(Resolve-Path -LiteralPath $Fixture).Path
$book=(Resolve-Path -LiteralPath $Workbook).Path
$hash=(Get-FileHash -LiteralPath $book).Hash
Copy-Item -LiteralPath (Join-Path $repo 'Structure.xml') -Destination $out.FullName
$compiler=New-Object System.CodeDom.Compiler.CompilerParameters
$compiler.GenerateExecutable=$true
$compiler.OutputAssembly=Join-Path $out.FullName ([IO.Path]::GetFileNameWithoutExtension($source)+'.exe')
$compiler.CompilerOptions='/platform:x86'
if($MatchApplicationAddressSpace){
    $applicationLargeAddressAware=Get-LargeAddressAware (Join-Path $bin 'Abovo-summit.exe')
    # The actual Summit AnyCPU/32-bit-preferred host is large-address-aware.
    # A hard x86 fixture otherwise has a smaller virtual address space, causing
    # unrelated native display allocations to fail first on a large workbook.
    if($applicationLargeAddressAware){$compiler.CompilerOptions='/platform:anycpu32bitpreferred'}
}
foreach($ref in @('System.dll','System.Core.dll','System.Drawing.dll','System.Windows.Forms.dll','System.Xml.dll','System.Data.dll','Microsoft.CSharp.dll')){[void]$compiler.ReferencedAssemblies.Add($ref)}
Get-ChildItem -LiteralPath $bin -Filter 'DevExpress*.v25.2*.dll' | ForEach-Object {[void]$compiler.ReferencedAssemblies.Add($_.FullName)}
if($MatchApplicationAddressSpace -and $applicationLargeAddressAware){
    # Legacy Framework CodeDom accepts 32bitpreferred but still emits a non-LAA
    # image. Use the installed Roslyn compiler, then verify the resulting flag.
    $roslyn=Join-Path ${env:ProgramFiles} 'Microsoft Visual Studio/2022/Enterprise/MSBuild/Current/Bin/Roslyn/csc.exe'
    if(-not (Test-Path -LiteralPath $roslyn)){throw 'Address-space matching requires the installed VS2022 Roslyn compiler'}
    $compileArguments=@('/nologo','/target:exe',$compiler.CompilerOptions,('/out:'+$compiler.OutputAssembly))
    $compileArguments+=@($compiler.ReferencedAssemblies | ForEach-Object {'/reference:'+$_})
    $compileArguments+=$source
    $compilerOutput=& $roslyn @compileArguments 2>&1
    if($LASTEXITCODE -ne 0){throw ($compilerOutput | Out-String)}
    if($compilerOutput){Write-Output $compilerOutput}
}else{
    $provider=New-Object Microsoft.CSharp.CSharpCodeProvider
    try{$compiled=$provider.CompileAssemblyFromFile($compiler,$source);if($compiled.Errors.HasErrors){throw ($compiled.Errors | Out-String)}}finally{$provider.Dispose()}
}
if($MatchApplicationAddressSpace){
    $fixtureLargeAddressAware=Get-LargeAddressAware $compiler.OutputAssembly
    Write-Output "ADDRESS_SPACE targetLAA=$applicationLargeAddressAware fixtureLAA=$fixtureLargeAddressAware compiler=$($compiler.CompilerOptions)"
    if($fixtureLargeAddressAware -ne $applicationLargeAddressAware){throw 'Fixture address-space flag differs from the application'}
}
if($UseApplicationConfig){Copy-Item -LiteralPath (Join-Path $bin 'Abovo-summit.exe.config') -Destination ($compiler.OutputAssembly+'.config')}
Write-Output "EVIDENCE=$($out.FullName)"
$run=Start-Process -FilePath $compiler.OutputAssembly -ArgumentList @(('"'+$bin+'"'),('"'+$out.FullName+'"'),('"'+$book+'"'),'False',('"'+$Mode+'"')) -WorkingDirectory $repo -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $out.FullName 'result.log') -RedirectStandardError (Join-Path $out.FullName 'error.log')
$null=$run.Handle
$run.WaitForExit()
Get-Content (Join-Path $out.FullName 'result.log')
Get-Content (Join-Path $out.FullName 'error.log')
if((Get-FileHash -LiteralPath $book).Hash -ne $hash){throw 'Source workbook changed'}
if($run.ExitCode -ne 0){throw "Regression failed: $($run.ExitCode); evidence $($out.FullName)"}
