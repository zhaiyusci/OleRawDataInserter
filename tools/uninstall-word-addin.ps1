$ErrorActionPreference = 'Stop'

$addinPath = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\OleRawDataInserter.dotm'

if (Test-Path $addinPath) {
    Remove-Item -LiteralPath $addinPath -Force
    "Uninstalled: $addinPath"
}
else {
    "Add-in is not installed: $addinPath"
}
