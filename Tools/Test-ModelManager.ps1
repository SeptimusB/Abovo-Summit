param(
    [string]$Repository = 'C:\Repos\Abovo Summit',
    [string]$StoriSource,
    [string]$BlankSource,
    [string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$binaryDirectory = Join-Path $Repository 'bin\Debug'
[void][Reflection.Assembly]::LoadFrom((Join-Path $binaryDirectory 'Abovo-summit.exe'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $binaryDirectory 'DevExpress.Docs.v25.2.dll'))
$testDirectory = Join-Path $Repository ('obj\ModelManagerTests\' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $testDirectory)
$results = New-Object 'System.Collections.Generic.List[string]'
function Assert-True([bool]$condition, [string]$message) {
    if (-not $condition) { throw $message }
    $results.Add('PASS: ' + $message)
}
function Expect-Rejection([scriptblock]$action, [string]$message) {
    $rejected = $false
    try { & $action } catch { $rejected = $true }
    Assert-True $rejected $message
}
function New-Definition([string]$name, [bool]$generic) {
    $definition = New-Object Abovo.ModelManagerDefinition
    $definition.DisplayName = $name
    $definition.IsGeneric = $generic
    if (-not $generic) { $definition.ClientName = 'Stori' }
    return $definition
}
function Get-PartHashes([string]$fileName) {
    $package = [IO.Compression.ZipFile]::OpenRead($fileName)
    $hashes = @{}
    try {
        foreach ($entry in $package.Entries) {
            $stream = $entry.Open()
            $sha = [Security.Cryptography.SHA256]::Create()
            try { $hashes[$entry.FullName] = [BitConverter]::ToString($sha.ComputeHash($stream)) }
            finally { $sha.Dispose(); $stream.Dispose() }
        }
    } finally { $package.Dispose() }
    return $hashes
}
function Assert-Preserved([string]$source, [string]$copy) {
    $before = Get-PartHashes $source
    $after = Get-PartHashes $copy
    $count = 0
    foreach ($name in $before.Keys) {
        if ($name -in @('[Content_Types].xml', 'xl/_rels/workbook.bin.rels')) { continue }
        if (-not $after.ContainsKey($name) -or $before[$name] -ne $after[$name]) { throw ('Original package part changed: ' + $name) }
        $count++
    }
    Assert-True ($count -gt 0) ('All ' + $count + ' unchanged original parts preserved: ' + [IO.Path]::GetFileName($copy))
}

$fixture = Join-Path $testDirectory 'two-parts.xlsb'
$managed = Join-Path $testDirectory 'three-parts.xlsb'
$native = Join-Path $testDirectory 'native-resaved.xlsb'
$definition = New-Definition 'Generic test definition' $true
$book = New-Object DevExpress.Spreadsheet.Workbook
try {
    $book.Options.CalculationMode = [DevExpress.Spreadsheet.WorkbookCalculationMode]::Manual
    [void]$book.CustomXmlParts.Add('<Abovo_Model_Def><ModelType>AbovoBP</ModelType></Abovo_Model_Def>')
    [void]$book.CustomXmlParts.Add('<Existing xmlns="urn:summit:test:unrelated">preserve me</Existing>')
    $book.SaveDocument($fixture, [DevExpress.Spreadsheet.DocumentFormat]::Xlsb)
} finally { $book.Dispose() }
Assert-True ($null -eq [Abovo.ModelManagerStore]::ReadEmbedded($fixture)) 'Missing manager definition is normal'
$badManager = Join-Path $testDirectory 'malformed-manager.xlsb'
$book = New-Object DevExpress.Spreadsheet.Workbook
try { $book.SaveDocument($badManager, [DevExpress.Spreadsheet.DocumentFormat]::Xlsb) } finally { $book.Dispose() }
$archive = [IO.Compression.ZipFile]::Open($badManager, [IO.Compression.ZipArchiveMode]::Update)
try {
    $stream = $archive.CreateEntry('customXml/badManager.xml').Open()
    $writer = [IO.StreamWriter]::new($stream)
    try { $writer.Write('<Abovo_Model_Manager xmlns="urn:abovo:summit:model-manager:1"><Invalid>') }
    finally { $writer.Dispose() }
} finally { $archive.Dispose() }
Expect-Rejection { [Abovo.ModelManagerStore]::ReadEmbedded($badManager) } 'Malformed manager XML rejected'
Assert-True ($null -eq [Abovo.EmbeddedWorkbookStructureReader]::TryRead($badManager, 'AbovoBP')) 'Malformed unrelated definition does not block Structure fallback'
[Abovo.ModelManagerStore]::SaveWorkbookCopy($fixture, $managed, $definition)
Assert-Preserved $fixture $managed
$loaded = [Abovo.ModelManagerStore]::ReadEmbedded($managed)
Assert-True ($loaded.DefinitionId -eq $definition.DefinitionId) 'Embedded definition readback'
$book = New-Object DevExpress.Spreadsheet.Workbook
try {
    $book.Options.CalculationMode = [DevExpress.Spreadsheet.WorkbookCalculationMode]::Manual
    $book.LoadDocument($managed)
    Assert-True ($book.CustomXmlParts.Count -eq 3) 'DevExpress 25.2 reads three independent XML parts'
    $book.SaveDocument($native, [DevExpress.Spreadsheet.DocumentFormat]::Xlsb)
} finally { $book.Dispose() }
Assert-True ([Abovo.ModelManagerStore]::ReadEmbedded($native).DefinitionId -eq $definition.DefinitionId) 'Manager survives native DevExpress XLSB save/reopen'
Assert-True ([Abovo.EmbeddedWorkbookStructureReader]::TryRead($native, 'AbovoBP').Contains('Abovo_Model_Def')) 'Existing Summit structure reader still selects only Structure XML'
$definition.Notes = 'Revision update test'
$updated = Join-Path $testDirectory 'updated-definition.xlsb'
[Abovo.ModelManagerStore]::SaveWorkbookCopy($native, $updated, $definition)
$book = New-Object DevExpress.Spreadsheet.Workbook
try { $book.LoadDocument($updated); Assert-True ($book.CustomXmlParts.Count -eq 3) 'Updating manager XML does not duplicate parts' } finally { $book.Dispose() }
Expect-Rejection { [Abovo.ModelManagerStore]::SaveWorkbookCopy($fixture, $managed, $definition) } 'Existing output is never overwritten'
Expect-Rejection { [Abovo.ModelManagerStore]::SaveWorkbookCopy($fixture, $fixture, $definition) } 'Source cannot be overwritten'
$xmlCopy = Join-Path $testDirectory 'definition.xml'
[Abovo.ModelManagerStore]::ExportDefinition($definition, $xmlCopy)
Assert-True ([Abovo.ModelManagerStore]::LoadDefinition($xmlCopy).Notes -eq $definition.Notes) 'Standalone definition roundtrip'
$revisionOne = [Abovo.ModelManagerStore]::SaveRevision($definition)
$revisionTwo = [Abovo.ModelManagerStore]::SaveRevision($definition)
Assert-True ((Test-Path -LiteralPath $revisionOne) -and (Test-Path -LiteralPath $revisionTwo) -and $revisionOne -ne $revisionTwo) 'Local revisions are immutable distinct files'
Assert-True ([Abovo.ModelManagerStore]::LoadDefinition($revisionTwo).Revision -eq 2) 'Local revision sequence persists'
# Delete only these two generated test revisions; no recursive or wildcard removal.
$revisionRoot = [IO.Path]::GetFullPath([Abovo.ModelManagerStore]::LibraryDirectory).TrimEnd('\') + '\'
foreach ($generatedRevision in @($revisionOne, $revisionTwo)) {
    if (-not [IO.Path]::GetFullPath($generatedRevision).StartsWith($revisionRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected test revision path.' }
    Remove-Item -LiteralPath $generatedRevision
}
[IO.Directory]::Delete([IO.Path]::GetDirectoryName($revisionOne), $false)
$doc = [Abovo.ModelManagerStore]::ToDocument($definition)
$doc.Root.SetAttributeValue('SchemaVersion', 99)
Expect-Rejection { [Abovo.ModelManagerStore]::Parse($doc) } 'Unsupported schema rejected'
$doc = [Abovo.ModelManagerStore]::ToDocument($definition)
$doc.Root.Add([System.Xml.Linq.XElement]::new([System.Xml.Linq.XName]::Get('Execute', [Abovo.ModelManagerStore]::XmlNamespace), 'arbitrary code'))
Expect-Rejection { [Abovo.ModelManagerStore]::Parse($doc) } 'Unknown executable-like element rejected'
$doc = [Abovo.ModelManagerStore]::ToDocument($definition)
$doc.Root.Element([System.Xml.Linq.XName]::Get('DefinitionId', [Abovo.ModelManagerStore]::XmlNamespace)).Remove()
Expect-Rejection { [Abovo.ModelManagerStore]::Parse($doc) } 'Missing identity is rejected rather than generated during load'
$duplicatePath = Join-Path $testDirectory 'duplicate-parts.xlsb'
$book = New-Object DevExpress.Spreadsheet.Workbook
try {
    [void]$book.LoadDocument($managed)
    [void]$book.CustomXmlParts.Add([Abovo.ModelManagerStore]::ToDocument($definition).ToString())
    $book.SaveDocument($duplicatePath, [DevExpress.Spreadsheet.DocumentFormat]::Xlsb)
} finally { $book.Dispose() }
Expect-Rejection { [Abovo.ModelManagerStore]::ReadEmbedded($duplicatePath) } 'Duplicate manager definitions rejected'
$invalidDefinition = New-Definition 'Missing client' $false
$invalidDefinition.ClientName = ''
Expect-Rejection { [Abovo.ModelManagerStore]::Validate($invalidDefinition) } 'Non-generic definition requires client identification'
$definition.Status = 'Approved'
Expect-Rejection { [Abovo.ModelManagerStore]::Validate($definition) } 'Draft trial cannot claim approval'
$definition.Status = 'Draft'
$verifier = [Abovo.ModelManagerAccess]::CreateVerifier('test-only-random-' + [guid]::NewGuid().ToString())
Assert-True (-not [Abovo.ModelManagerAccess]::Verify('incorrect', $verifier)) 'Incorrect password rejected'
$testPassword = [guid]::NewGuid().ToString()
$verifier = [Abovo.ModelManagerAccess]::CreateVerifier($testPassword)
Assert-True ([Abovo.ModelManagerAccess]::Verify($testPassword, $verifier)) 'Salted password verification works'
$testPassword = $null
Assert-True (-not [Abovo.ModelManagerAccess]::Verify('anything', '')) 'Missing live verifier fails closed'

if ($StoriSource -and $BlankSource -and $OutputDirectory) {
    [void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
    foreach ($spec in @(@($StoriSource, $false, 'Stori - Model Manager Trial.xlsb'), @($BlankSource, $true, 'Blank BP v26_0001 - Model Manager Trial.xlsb'))) {
        $source = [string]$spec[0]
        $generic = [bool]$spec[1]
        $copy = Join-Path $OutputDirectory $spec[2]
        $beforeHash = [Abovo.ModelManagerStore]::Fingerprint($source)
        $profile = New-Definition ([IO.Path]::GetFileNameWithoutExtension($source)) $generic
        $profile.ModelRole = if ($generic) { [Abovo.ManagedModelRole]::Template } else { [Abovo.ManagedModelRole]::PopulatedModel }
        $profile.Notes = 'Draft demonstration only. No population, bespoke changes, calculation or approval has been performed.'
        $evidence = New-Object Abovo.ModelManagerEvidence
        $evidence.Role = if ($generic) { 'Latest template' } else { 'Source' }
        $evidence.FileName = [IO.Path]::GetFileName($source)
        $evidence.SHA256 = $beforeHash
        $profile.Evidence.Add($evidence)
        if (-not $generic) {
            foreach ($description in @('Stock capacity and Transactional DB mirrors', 'Management cost capacity and related ranges', 'Stock grouping and cost-driver references', 'Client-specific covenants, workings and chart signs')) {
                $rule = New-Object Abovo.ModelManagerRule
                $rule.Description = $description
                $rule.Disposition = [Abovo.ModelRuleDisposition]::Unreviewed
                $rule.Notes = 'Known review area; exact executable rule and acceptance checks must be agreed before migration.'
                $profile.Rules.Add($rule)
            }
        }
        [Abovo.ModelManagerStore]::SaveWorkbookCopy($source, $copy, $profile)
        Assert-Preserved $source $copy
        Assert-True ([Abovo.ModelManagerStore]::Fingerprint($source) -eq $beforeHash) ('Original unchanged: ' + [IO.Path]::GetFileName($source))
        Assert-True ([Abovo.ModelManagerStore]::ReadEmbedded($copy).IsGeneric -eq $generic) ('Classification readback: ' + [IO.Path]::GetFileName($copy))
        $results.Add('OUTPUT: ' + $copy)
    }
}
$results | ForEach-Object { Write-Output $_ }
Write-Output ('Test artifacts: ' + $testDirectory)
