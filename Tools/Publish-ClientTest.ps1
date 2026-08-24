param(
    [string]$Version = '1.0.0.0',
    [string]$PublishDirectory = 'Z:\Sandbox\Deploy\Summit Test 1.00',
    [string]$ZipPath = 'Z:\Sandbox\Deploy\Abovo-Summit-Test-1.00-ClickOnce.zip'
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$projectPath = Join-Path $repositoryRoot 'Abovo Business Suite.vbproj'
$msbuild = 'C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe'

if (-not (Test-Path -LiteralPath $msbuild)) { throw "MSBuild was not found at $msbuild" }
if (Test-Path -LiteralPath $PublishDirectory) { throw "Publish destination already exists. Choose a new versioned folder: $PublishDirectory" }
if (Test-Path -LiteralPath $ZipPath) { throw "Delivery archive already exists. Choose a new archive path: $ZipPath" }

New-Item -ItemType Directory -Path $PublishDirectory | Out-Null

& $msbuild $projectPath /t:Publish /m /v:minimal /p:Configuration=Release /p:Platform=AnyCPU "/p:PublishDir=$PublishDirectory\" "/p:PublishUrl=$PublishDirectory\" "/p:ApplicationVersion=$Version" /p:ApplicationRevision=0 /p:SignManifests=false /p:BootstrapperEnabled=true
if ($LASTEXITCODE -ne 0) { throw "ClickOnce publish failed with exit code $LASTEXITCODE" }

$applicationDirectory = Get-ChildItem -LiteralPath (Join-Path $PublishDirectory 'Application Files') -Directory |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $applicationDirectory) { throw 'ClickOnce did not produce an application version directory.' }

$requiredRelativeFiles = @(
    'Structure.xml.deploy',
    'Help\index.html.deploy',
    'Help\README.md.deploy',
    'Help\assets\summit-help.css.deploy',
    'Help\assets\summit-help.js.deploy',
    'Help\data\interfaces.js.deploy',
    'Help\data\overrides.js.deploy'
)
$missing = @($requiredRelativeFiles | Where-Object { -not (Test-Path -LiteralPath (Join-Path $applicationDirectory.FullName $_)) })
if ($missing.Count -gt 0) { throw "Published package is missing required startup content: $($missing -join ', ')" }

$manifest = Get-ChildItem -LiteralPath $applicationDirectory.FullName -Filter '*.exe.manifest' | Select-Object -First 1
if ($null -eq $manifest) { throw 'ClickOnce application manifest was not produced.' }
$manifestText = Get-Content -Raw -LiteralPath $manifest.FullName
foreach ($manifestPath in @('Structure.xml', 'Help\index.html', 'Help\assets\summit-help.css', 'Help\assets\summit-help.js', 'Help\data\interfaces.js', 'Help\data\overrides.js')) {
    if ($manifestText -notmatch [regex]::Escape($manifestPath)) { throw "Application manifest does not include $manifestPath" }
}

$readme = @"
ABOVO SUMMIT TEST RELEASE 1.00
================================

Purpose
This unsigned ClickOnce package is for initial client functional testing only.

Installation
1. Extract the entire delivery archive to a local folder.
2. Run setup.exe.
3. Windows may display an Unknown Publisher warning because this test package is not code-signed.
4. Accept installation only if this archive was supplied directly by Abovo.

Requirements
- Windows with Microsoft .NET Framework 4.8.
- Access to the XLSB model files used for testing.

Package content contract
- Structure.xml is installed directly in the application startup directory.
- The complete Help folder is installed beneath the application startup directory.
- This is an offline test installation and does not automatically update.

Uninstall
Remove Abovo Summit Test from Windows Settings > Apps > Installed apps.
"@
[IO.File]::WriteAllText((Join-Path $PublishDirectory 'CLIENT_TEST_README.txt'), $readme, [Text.UTF8Encoding]::new($false))

Compress-Archive -Path (Join-Path $PublishDirectory '*') -DestinationPath $ZipPath -CompressionLevel Optimal

$publishedFiles = Get-ChildItem -LiteralPath $PublishDirectory -Recurse -File
$summary = [ordered]@{
    Version = $Version
    PublishDirectory = $PublishDirectory
    DeliveryArchive = $ZipPath
    ApplicationDirectory = $applicationDirectory.FullName
    PublishedFileCount = $publishedFiles.Count
    PublishedBytes = ($publishedFiles | Measure-Object Length -Sum).Sum
    ArchiveBytes = (Get-Item -LiteralPath $ZipPath).Length
    StructureAtStartup = $true
    HelpAtStartup = $true
    Signed = $false
}
$summary | Format-List
