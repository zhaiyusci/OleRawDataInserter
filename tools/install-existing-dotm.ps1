param(
    [Parameter(Mandatory = $true)]
    [string]$DotmPath
)

$ErrorActionPreference = 'Stop'

$resolvedDotm = (Resolve-Path $DotmPath).Path
$startupDir = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$targetPath = Join-Path $startupDir 'OleRawDataInserter.dotm'

New-Item -ItemType Directory -Force -Path $startupDir | Out-Null
Copy-Item -LiteralPath $resolvedDotm -Destination $targetPath -Force

"Installed: $targetPath"
