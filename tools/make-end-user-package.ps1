param(
    [string]$ProjectRoot = '',
    [string]$DotmPath = (Join-Path $ProjectRoot 'dist\OleRawDataInserter.dotm')
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    $ProjectRoot = (Resolve-Path (Join-Path $scriptRoot '..')).Path
}

if ([string]::IsNullOrWhiteSpace($DotmPath)) {
    $DotmPath = Join-Path $ProjectRoot 'dist\OleRawDataInserter.dotm'
}

$resolvedDotm = (Resolve-Path $DotmPath).Path
$packageTemplate = Join-Path $ProjectRoot 'installer\EndUserPackage'
$payloadDir = Join-Path $packageTemplate 'Payload'
$releaseDir = Join-Path $ProjectRoot 'release'
$releasePackageDir = Join-Path $releaseDir 'FigurePackageWordAddin'
$zipPath = Join-Path $releaseDir 'FigurePackageWordAddin.zip'

if (-not (Test-Path $resolvedDotm)) {
    throw "Could not find built add-in: $DotmPath"
}

if (Test-Path $releasePackageDir) {
    Remove-Item -LiteralPath $releasePackageDir -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $releasePackageDir | Out-Null

Copy-Item -Path (Join-Path $packageTemplate '*') -Destination $releasePackageDir -Recurse -Force
New-Item -ItemType Directory -Force -Path (Join-Path $releasePackageDir 'Payload') | Out-Null
Copy-Item -LiteralPath $resolvedDotm -Destination (Join-Path $releasePackageDir 'Payload\OleRawDataInserter.dotm') -Force

if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
Compress-Archive -Path (Join-Path $releasePackageDir '*') -DestinationPath $zipPath -Force

"Created package folder: $releasePackageDir"
"Created package zip: $zipPath"
