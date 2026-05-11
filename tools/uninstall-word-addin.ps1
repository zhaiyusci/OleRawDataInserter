$ErrorActionPreference = 'Stop'

$startupDir = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$installedFiles = @(
    (Join-Path $startupDir 'OLEPackager.dotm'),
    (Join-Path $startupDir 'OLEPackagerZipTool.exe'),
    (Join-Path $startupDir 'OleRawDataInserter.dotm'),
    (Join-Path $startupDir 'FigurePackageZipTool.exe')
)

foreach ($installedFile in $installedFiles) {
    if (Test-Path $installedFile) {
        Remove-Item -LiteralPath $installedFile -Force
        "Uninstalled: $installedFile"
    }
    else {
        "File is not installed: $installedFile"
    }
}
