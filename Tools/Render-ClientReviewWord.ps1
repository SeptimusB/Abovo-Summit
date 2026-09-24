param([Parameter(Mandatory=$true)][string]$Document, [Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
# Fallback only when the packaged renderer cannot locate bundled LibreOffice.
# Independent Word instance; read-only input; macros disabled; never save the DOCX.
[void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
$source=(Resolve-Path -LiteralPath $Document).Path
$destination=(Resolve-Path -LiteralPath $OutputDirectory).Path
$hash=(Get-FileHash -LiteralPath $source).Hash
$word=$null;$doc=$null
try {
    $word=New-Object -ComObject Word.Application
    Write-Output 'Word started'
    $word.Visible=$false
    $word.DisplayAlerts=0
    $word.AutomationSecurity=3
    Write-Output 'Opening read-only report'
    $doc=$word.Documents.Open($source,$false,$true,$false)
    Write-Output 'Exporting PDF'
    $doc.ExportAsFixedFormat((Join-Path $destination 'review.pdf'),17)
    Write-Output 'PDF exported'
} finally {
    if($doc){$doc.Close(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}
    if($word){$word.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)}
    if((Get-FileHash -LiteralPath $source).Hash -ne $hash){throw 'Report changed during read-only rendering'}
}
Write-Output (Join-Path $destination 'review.pdf')
