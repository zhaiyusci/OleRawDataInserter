$ErrorActionPreference = 'Stop'

$packageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$payloadPath = Join-Path $packageRoot 'Payload\OleRawDataInserter.dotm'
$zipToolPayloadPath = Join-Path $packageRoot 'Payload\FigurePackageZipTool.exe'
$startupDir = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$targetPath = Join-Path $startupDir 'OleRawDataInserter.dotm'
$zipToolTargetPath = Join-Path $startupDir 'FigurePackageZipTool.exe'
$logPath = Join-Path $packageRoot 'install.log'

function Write-InstallLog {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Write-Host $line
    Add-Content -LiteralPath $logPath -Encoding UTF8 -Value $line
}

Set-Content -LiteralPath $logPath -Encoding UTF8 -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Starting installation"

try {
    if (-not (Test-Path $payloadPath)) {
        throw "Missing Payload\OleRawDataInserter.dotm. This installer package is incomplete."
    }
    if (-not (Test-Path $zipToolPayloadPath)) {
        throw "Missing Payload\FigurePackageZipTool.exe. This installer package is incomplete."
    }

    $wordProcesses = Get-Process WINWORD -ErrorAction SilentlyContinue
    if ($wordProcesses) {
        Write-InstallLog 'Microsoft Word is currently running. Please close Word and run Install.cmd again.'
        throw 'Microsoft Word is running.'
    }

    Write-InstallLog "Creating Word STARTUP folder: $startupDir"
    New-Item -ItemType Directory -Force -Path $startupDir | Out-Null

    Write-InstallLog "Installing add-in to: $targetPath"
    Copy-Item -LiteralPath $payloadPath -Destination $targetPath -Force
    Write-InstallLog "Installing zip helper to: $zipToolTargetPath"
    Copy-Item -LiteralPath $zipToolPayloadPath -Destination $zipToolTargetPath -Force

    Write-InstallLog 'Installation completed successfully.'
}
catch {
    Write-InstallLog "ERROR: $($_.Exception.Message)"
    exit 1
}
