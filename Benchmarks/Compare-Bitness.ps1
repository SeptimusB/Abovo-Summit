<#
.SYNOPSIS
Compare Summit Release x86 and x64 on the same XLSB without saving it.
.EXAMPLE
.\Benchmarks\Compare-Bitness.ps1 -WorkbookPath 'C:\Sandbox\Populated BP.xlsb' -Runs 3
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,
    [ValidateRange(1, 20)]
    [int]$Runs = 3,
    [string]$EditRangeName = 'TransRents',
    [string]$OutputDirectory = [IO.Path]::Combine([IO.Path]::GetTempPath(), 'SummitBitnessBenchmarks'),
    [switch]$NoBuild
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$workbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$initialHash = (Get-FileHash -LiteralPath $workbook -Algorithm SHA256).Hash
$runDirectory = Join-Path $OutputDirectory (Get-Date -Format 'yyyyMMdd-HHmmss')
New-Item -ItemType Directory -Path $runDirectory -Force | Out-Null

if (-not $NoBuild) {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (-not (Test-Path -LiteralPath $vswhere)) {
        throw 'Visual Studio vswhere.exe was not found. Build Release x86/x64 in Visual Studio, then use -NoBuild.'
    }
    $msbuild = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' |
        Select-Object -First 1
    if (-not $msbuild) { throw 'MSBuild was not found.' }
    foreach ($architecture in @('x86', 'x64')) {
        & $msbuild (Join-Path $repoRoot 'Abovo Business Suite.sln') /t:Build /p:Configuration=Release "/p:Platform=$architecture" /v:minimal /nologo
        if ($LASTEXITCODE -ne 0) { throw "Release $architecture build failed." }
    }
}

$context = [ordered]@{
    workbook = $workbook
    workbookSha256 = $initialHash
    workbookBytes = (Get-Item -LiteralPath $workbook).Length
    editRangeName = $EditRangeName
    runsPerArchitecture = $Runs
    started = (Get-Date).ToString('o')
    processorCount = [Environment]::ProcessorCount
    gitCommit = (& git -C $repoRoot rev-parse HEAD)
    gitDirty = [bool](& git -C $repoRoot status --porcelain)
}
$context | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $runDirectory 'context.json') -Encoding UTF8

for ($iteration = 1; $iteration -le $Runs; $iteration++) {
    $order = if ($iteration % 2 -eq 1) { @('x86', 'x64') } else { @('x64', 'x86') }
    foreach ($architecture in $order) {
        $exe = Join-Path $repoRoot "bin\$architecture\Release\Abovo-summit.exe"
        if (-not (Test-Path -LiteralPath $exe)) { throw "Missing Release executable: $exe" }
        $result = Join-Path $runDirectory ("run-{0:D2}-{1}.tsv" -f $iteration, $architecture)
        $arguments = '--benchmark-bitness "{0}" "{1}" "{2}"' -f $workbook, $result, $EditRangeName
        $process = Start-Process -FilePath $exe -ArgumentList $arguments -WorkingDirectory $repoRoot -WindowStyle Hidden -PassThru
        if (-not $process.WaitForExit(180000)) {
            Stop-Process -Id $process.Id -Force
            throw "Release $architecture benchmark timed out after 180 seconds."
        }
        if ($process.ExitCode -ne 0) { throw "Release $architecture exited with code $($process.ExitCode)." }
        if (-not (Test-Path -LiteralPath $result)) { throw "No benchmark output from Release $architecture." }
        $phases = Import-Csv -LiteralPath $result -Delimiter "`t"
        $failed = @($phases | Where-Object { $_.outcome -notlike 'ok*' })
        if ($failed.Count -gt 0) {
            throw "Release $architecture reported an incomplete phase: $($failed | ConvertTo-Json -Compress)"
        }
        $currentHash = (Get-FileHash -LiteralPath $workbook -Algorithm SHA256).Hash
        if ($currentHash -ne $initialHash) { throw 'The benchmark workbook changed on disk. Stop and inspect it.' }
        Write-Host ("Run {0}/{1}, {2}: {3}" -f $iteration, $Runs, $architecture, $result)
    }
}

$rows = Get-ChildItem -LiteralPath $runDirectory -Filter 'run-*.tsv' |
    ForEach-Object { Import-Csv -LiteralPath $_.FullName -Delimiter "`t" }
$summary = $rows | Group-Object bitness, phase | ForEach-Object {
    $times = @($_.Group | ForEach-Object { [double]$_.elapsed_ms } | Sort-Object)
    $middle = [int][Math]::Floor($times.Count / 2)
    $median = if ($times.Count % 2 -eq 0) {
        ($times[$middle - 1] + $times[$middle]) / 2
    } else {
        $times[$middle]
    }
    [pscustomobject]@{
        bitness = $_.Group[0].bitness
        phase = $_.Group[0].phase
        runs = $times.Count
        median_ms = $median
        min_ms = $times[0]
        max_ms = $times[-1]
    }
} | Sort-Object phase, bitness
$summary | Export-Csv -LiteralPath (Join-Path $runDirectory 'summary.csv') -NoTypeInformation
$summary | Format-Table -AutoSize
Write-Host "Results: $runDirectory"
