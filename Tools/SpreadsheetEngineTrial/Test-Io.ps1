param(
    [Parameter(Mandatory=$true)][string]$LicencePath,
    [ValidateSet('Blank','Demo','AGL','Stori')][string[]]$Cases = @('Demo'),
    [ValidateRange(3,5)][int]$Runs = 3,
    [ValidateSet('x86','x64')][string]$Architecture = 'x64'
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$exe = Join-Path $PSScriptRoot "bin/$Architecture/Release/net48/SpreadsheetEngineTrial.exe"
if (-not (Test-Path -LiteralPath $exe)) { throw 'Build the isolated Release trial first.' }
if (-not (Test-Path -LiteralPath $LicencePath)) { throw 'Explicit external licence required.' }
$sources = @{
    Blank = Join-Path $repo 'Library/Blank BP v26_0001.xlsb'
    Demo = Join-Path $repo 'Library/Demo BP v26_0001.xlsb'
    AGL = 'C:/Sandbox/Insert Comp/AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb'
    Stori = 'D:/Downloads/Stori BP v25_0704 2026-27 V4 update 100826 CLEAN pass 3.xlsb'
}
$out = Join-Path $repo ('obj/AsposeTrial/' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$previousLicence = $env:SUMMIT_ASPOSE_LICENSE_PATH
$records = @()
try {
    $env:SUMMIT_ASPOSE_LICENSE_PATH = $LicencePath
    foreach ($case in $Cases) {
        $original = $sources[$case]
        $before = (Get-FileHash -LiteralPath $original -Algorithm SHA256).Hash
        $source = Join-Path $out ($case + '-source.xlsb')
        Copy-Item -LiteralPath $original -Destination $source
        try {
            if ((Get-FileHash -LiteralPath $source).Hash -ne $before) { throw 'Copy hash mismatch.' }
            # Fresh process per run; sequential, with alternating engine order.
            # No OS-cache flush: these are warm-cache observations, not cold-start claims.
            for ($run = 1; $run -le $Runs; $run++) {
                $engines = if ($run % 2) { @('aspose','devexpress') } else { @('devexpress','aspose') }
                foreach ($format in @('xlsb','xlsm')) {
                    foreach ($engine in $engines) {
                        $stem = "$case-$engine-$format-$run"
                        $output = Join-Path $out ($stem + '.' + $format)
                        $result = Join-Path $out ($stem + '.json')
                        & $exe io $engine $source $output $result
                        if ($LASTEXITCODE -ne 0) { throw "I/O failed: $stem" }
                        $record = Get-Content -LiteralPath $result -Raw | ConvertFrom-Json
                        $record | Add-Member NoteProperty case $case
                        $record | Add-Member NoteProperty run $run
                        $record | Add-Member NoteProperty format $format
                        $records += $record
                    }
                }
            }
        } finally {
            if ((Get-FileHash -LiteralPath $original).Hash -ne $before) { throw "Original changed: $case" }
            if ((Get-FileHash -LiteralPath $source).Hash -ne $before) { throw "Private input changed: $case" }
        }
    }
} finally {
    $env:SUMMIT_ASPOSE_LICENSE_PATH = $previousLicence
    # Generated diagnostic evidence, never a workbook overwrite.
    $records | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $out 'runs.json') -Encoding UTF8
    Write-Output "TRIAL_OUTPUT=$out"
}
Write-Output 'Raw native I/O only; no financial calculation, Summit compatibility guards, UI rebind or Excel/VBA execution. Preservation is a separate gate.'
