param(
    [Parameter(Mandatory = $true)]
    [string]$DotmPath,

    [string]$ZipToolPath = ''
)

$ErrorActionPreference = 'Stop'

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$resolvedDotm = (Resolve-Path $DotmPath).Path
if ([string]::IsNullOrWhiteSpace($ZipToolPath)) {
    $ZipToolPath = Join-Path $projectRoot 'dist\FigurePackageZipTool.exe'
}
$resolvedZipTool = (Resolve-Path $ZipToolPath).Path
$startupDir = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$targetPath = Join-Path $startupDir 'OleRawDataInserter.dotm'
$zipToolTargetPath = Join-Path $startupDir 'FigurePackageZipTool.exe'

New-Item -ItemType Directory -Force -Path $startupDir | Out-Null
Copy-Item -LiteralPath $resolvedDotm -Destination $targetPath -Force
Copy-Item -LiteralPath $resolvedZipTool -Destination $zipToolTargetPath -Force

"Installed: $targetPath"
"Installed: $zipToolTargetPath"
