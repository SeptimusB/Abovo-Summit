param([string]$Configuration='Release')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$project=[xml](Get-Content (Join-Path $repo 'Abovo Business Suite.vbproj') -Raw)
foreach($entry in $project.Project.ItemGroup.Compile){
    if(!$entry.Include -or $entry.Include -like '*SummitDiagnostics.vb'){continue}
    $source=Get-Content (Join-Path $repo $entry.Include) -Raw
    if($source -match '\b(?:Trace|Debug|Console)\.(?:WriteLine|Write|Print)\('){throw "Ungated diagnostic in $($entry.Include)"}
}
Write-Output 'PASS: all compiled application diagnostic outputs use the quiet-build gate'
$helper=Get-Content (Join-Path $repo 'Services/Development and Debug/SummitDiagnostics.vb') -Raw
$probe=@"
Namespace Abovo
    Friend Class CheckSheetWatch
        Friend Shared Benchmark As Boolean
    End Class
    Public Class DiagnosticProbe
        Private Shared Arguments As Integer
        Private Shared Function Argument() As String
            Arguments += 1
            Return "probe"
        End Function
        Public Shared Function Run() As Integer
            Arguments = 0
            SummitDiagnostics.WriteLine(Argument())
            Return Arguments
        End Function
        Public Shared Function RunningClockExists() As Boolean
            Dim timer = SummitDiagnostics.DiagnosticTimer.StartNew()
            timer.Restart()
            Dim field = timer.GetType().GetField("Clock", System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic)
            Return field.GetValue(timer) IsNot Nothing
        End Function
        Public Shared Function TrialOutput(enabled As Boolean) As String
            CheckSheetWatch.Benchmark = enabled
            Dim sink As New System.IO.StringWriter()
            Dim listener As New System.Diagnostics.TextWriterTraceListener(sink)
            System.Diagnostics.Trace.Listeners.Add(listener)
            Try
                SummitDiagnostics.WriteTrialLine("trial marker")
                System.Diagnostics.Trace.Flush()
                Return sink.ToString()
            Finally
                System.Diagnostics.Trace.Listeners.Remove(listener)
                listener.Dispose()
            End Try
        End Function
    End Class
End Namespace
"@
foreach($enabled in @($false,$true)){
    $compiler=New-Object System.CodeDom.Compiler.CompilerParameters
    $compiler.GenerateInMemory=$true
    $compiler.CompilerOptions='/define:SUMMIT_DIAGNOSTICS='+$enabled.ToString()+',TRACE=True'
    [void]$compiler.ReferencedAssemblies.Add('System.dll')
    $provider=New-Object Microsoft.VisualBasic.VBCodeProvider
    try{$result=$provider.CompileAssemblyFromSource($compiler,@($helper,$probe)); if($result.Errors.HasErrors){throw ($result.Errors | Out-String)}}finally{$provider.Dispose()}
    $type=$result.CompiledAssembly.GetType('Abovo.DiagnosticProbe')
    $evaluated=$type.GetMethod('Run').Invoke($null,@())
    $clock=$type.GetMethod('RunningClockExists').Invoke($null,@())
    if($evaluated -ne [int]$enabled -or $clock -ne $enabled){throw 'Conditional arguments or passive clock gate failed'}
    Write-Output "PASS: diagnostics=$enabled; argument evaluations=$evaluated; active timing=$clock"
    foreach($trial in @($false,$true)) {
        $trialOutput=$type.GetMethod('TrialOutput').Invoke($null,@($trial))
        if(($trialOutput -like '*trial marker*') -ne $trial){throw 'Runtime trial output gate failed'}
        Write-Output "PASS: diagnostics=$enabled; explicit runtime trial output=$trial"
    }
}
$assembly=[Reflection.Assembly]::LoadFile((Join-Path $repo ('bin/'+$Configuration+'/Abovo-summit.exe')))
$enabledField=$assembly.GetType('Abovo.SummitDiagnostics').GetField('Enabled')
if($enabledField.GetRawConstantValue()){throw 'Delivery executable still enables diagnostics'}
Write-Output "PASS: $Configuration delivery executable diagnostics disabled"
$integrity=Get-Content (Join-Path $repo 'Services/IdleIntegrityManager.vb') -Raw
if(([regex]::Matches($integrity,'\bStopwatch.StartNew\(\)')).Count -ne 3){throw 'Review functional integrity timers: idle clock and two time-sliced scans must remain real Stopwatches'}
Write-Output 'PASS: functional integrity scheduling and scan time budgets retain real timers'
