$ErrorActionPreference = 'Stop'

$packageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$startupDir = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$targetPath = Join-Path $startupDir 'OleRawDataInserter.dotm'
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

    if (Test-Path $targetPath) {
        Write-InstallLog "Removing add-in: $targetPath"
        Remove-Item -LiteralPath $targetPath -Force
        Write-InstallLog 'Uninstall completed successfully.'
    }
    else {
        Write-InstallLog "Add-in is not installed: $targetPath"
    }
}
catch {
    Write-InstallLog "ERROR: $($_.Exception.Message)"
    exit 1
}
