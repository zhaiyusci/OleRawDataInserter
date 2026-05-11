$ErrorActionPreference = 'Stop'

$packageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$startupDir = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$targetPath = Join-Path $startupDir 'OLEPackager.dotm'
$zipToolTargetPath = Join-Path $startupDir 'OLEPackagerZipTool.exe'
$legacyTargetPath = Join-Path $startupDir 'OleRawDataInserter.dotm'
$legacyZipToolTargetPath = Join-Path $startupDir 'FigurePackageZipTool.exe'
$logPath = Join-Path $packageRoot 'install.log'

function Write-InstallLog {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Write-Host $line
    Add-Content -LiteralPath $logPath -Encoding UTF8 -Value $line
}

Add-Content -LiteralPath $logPath -Encoding UTF8 -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Starting uninstall"

try {
    $wordProcesses = Get-Process WINWORD -ErrorAction SilentlyContinue
    if ($wordProcesses) {
        Write-InstallLog 'Microsoft Word is currently running. Please close Word and run Uninstall.cmd again.'
        throw 'Microsoft Word is running.'
    }

    foreach ($path in @($targetPath, $zipToolTargetPath, $legacyTargetPath, $legacyZipToolTargetPath)) {
        if (Test-Path $path) {
            Write-InstallLog "Removing: $path"
            Remove-Item -LiteralPath $path -Force
        }
        else {
            Write-InstallLog "File is not installed: $path"
        }
    }
    Write-InstallLog 'Uninstall completed successfully.'
}
catch {
    Write-InstallLog "ERROR: $($_.Exception.Message)"
    exit 1
}
