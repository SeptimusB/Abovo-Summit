param([Parameter(Mandatory=$true)][string]$Baseline,[Parameter(Mandatory=$true)][string[]]$Candidates,[Parameter(Mandatory=$true)][string]$Report,[switch]$AllowRoundtripInputs)
$ErrorActionPreference='Stop'
$trialRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../obj/AsposeTrial'))+[IO.Path]::DirectorySeparatorChar
foreach($path in @($Baseline)+$Candidates+@($Report)){
    if(-not [IO.Path]::GetFullPath($path).StartsWith($trialRoot,[StringComparison]::OrdinalIgnoreCase)){throw 'Private trial paths required.'}
    if((Test-Path -LiteralPath $path) -and ((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Reparse points not permitted.'}
    for($parent=[IO.DirectoryInfo]::new([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($path)));$null -ne $parent;$parent=$parent.Parent){
        if($parent.Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Reparse-point parent not permitted.'}
    }
}
if(Test-Path -LiteralPath $Report){throw 'New report required.'}
$reference=Get-Content -LiteralPath $Baseline -Raw|ConvertFrom-Json
if(-not $reference.success){throw 'Baseline run did not complete.'}
function Is-Number($value){return $null -ne $value -and $value -isnot [string] -and $value -isnot [bool] -and $value -is [ValueType] -and $value -isnot [datetime]}
function Cell-Value($probe,[int]$row,[int]$col){
    if($null -eq $probe -or $row -ge $probe.rows -or $col -ge $probe.columns){return $null}
    $v=$probe.values[$row][$col]
    if($v -is [string] -and $v.Length -eq 0){return $null}
    return $v
}
$results=@()
foreach($path in $Candidates){
    $candidate=Get-Content -LiteralPath $path -Raw|ConvertFrom-Json
    if(-not $candidate.success){throw 'Candidate run did not complete.'}
    if($candidate.inputHash -ne $reference.inputHash -and -not $AllowRoundtripInputs){throw 'Inputs differ; use AllowRoundtripInputs only for explicitly paired saved/reopened comparisons.'}
    foreach($sheetProperty in $reference.calculatedProbes.PSObject.Properties){
        $name=$sheetProperty.Name;$a=$sheetProperty.Value;$b=$candidate.calculatedProbes.$name
        if($null -eq $b){throw ('Missing sheet '+$name)}
        $numeric=0;$types=0;$text=0;$baselineErrors=0;$candidateErrors=0;$compared=0;$maxDifference=0.0;$samples=@()
        for($row=0;$row -lt [Math]::Max($a.rows,$b.rows);$row++){
            for($col=0;$col -lt [Math]::Max($a.columns,$b.columns);$col++){
                $x=Cell-Value $a $row $col;$y=Cell-Value $b $row $col
                if($x -is [string] -and $x.StartsWith('#')){$baselineErrors++}
                if($y -is [string] -and $y.StartsWith('#')){$candidateErrors++}
                if($null -eq $x -and $null -eq $y){continue};$compared++
                $different=$false
                if((Is-Number $x) -and (Is-Number $y)){
                    $delta=[Math]::Abs([double]$x-[double]$y)
                    $maxDifference=[Math]::Max($maxDifference,$delta)
                    if($delta -gt [Math]::Max(0.000001,[Math]::Abs([double]$x)*1e-10)){$numeric++;$different=$true}
                }elseif($x -is [string] -and $y -is [string]){
                    if($x -cne $y){$text++;$different=$true}
                }elseif(-not [object]::Equals($x,$y)){$types++;$different=$true}
                if($different -and $samples.Count -lt 5){$samples += [pscustomobject]@{row=$row+1;column=$col+1;excel=$x;candidate=$y}}
            }
        }
        $results += [pscustomobject]@{engine=$candidate.engine;sheet=$name;cellsCompared=$compared;numericDifferences=$numeric;typeOrBlankDifferences=$types;textDifferences=$text;maxAbsoluteNumericDifference=$maxDifference;excelErrors=$baselineErrors;candidateErrors=$candidateErrors;samples=$samples}
    }
}
$data=[pscustomobject]@{baseline=$Baseline;candidates=$Candidates;inputHash=$reference.inputHash;allowRoundtripInputs=[bool]$AllowRoundtripInputs;relativeTolerance=1e-10;absoluteTolerance=0.000001;scope='Five calculated worksheet ranges only; not complete workbook or financial approval';results=$results}
$json=$data|ConvertTo-Json -Depth 8
$stream=[IO.File]::Open($Report,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write)
try{$writer=[IO.StreamWriter]::new($stream);$writer.Write($json);$writer.Flush()}finally{if($writer){$writer.Dispose()}else{$stream.Dispose()}}
$results|Select-Object engine,sheet,cellsCompared,numericDifferences,typeOrBlankDifferences,textDifferences,candidateErrors|Format-Table -AutoSize
