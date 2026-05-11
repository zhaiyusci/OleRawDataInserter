param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [switch]$Install,
    [switch]$SkipVbomSetup,
    [switch]$CloseWord,
    [string]$OfficeVersion = ''
)

# Builds the Word add-in using the same Word automation sequence that has
# proven reliable on this project:
#   1. Save a plain .dotm to warm up Word's template save path.
#   2. Save a second .dotm after importing the VBA modules and UserForms.
#   3. Quit Word before copying to dist, injecting Ribbon XML, or installing.
#
# Keeping all file deployment outside the live Word process avoids locks on
# Word STARTUP templates and avoids SaveAs2 hangs caused by loaded add-ins.

$ErrorActionPreference = 'Stop'

$srcDir = Join-Path $ProjectRoot 'src'
$distDir = Join-Path $ProjectRoot 'dist'
$addinPath = Join-Path $distDir 'OLEPackager.dotm'
$zipToolPath = Join-Path $distDir 'OLEPackagerZipTool.exe'
$tempBuildDir = Join-Path $env:LOCALAPPDATA 'Temp\OLEPackagerBuild'
$plainAddinPath = Join-Path $tempBuildDir 'PlainSaveTest.dotm'
$tempAddinPath = Join-Path $tempBuildDir 'VbaSaveTest.dotm'
$logPath = Join-Path $distDir 'build-word-addin.log'
$startupDir = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$startupAddinPath = Join-Path $startupDir 'OLEPackager.dotm'
$startupZipToolPath = Join-Path $startupDir 'OLEPackagerZipTool.exe'
$legacyStartupAddinPath = Join-Path $startupDir 'OleRawDataInserter.dotm'
$legacyStartupZipToolPath = Join-Path $startupDir 'FigurePackageZipTool.exe'
$securityKey = $null
$valueName = 'AccessVBOM'
$hadOriginalValue = $false
$originalValue = $null

function Write-Step {
    param([string]$Message)

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Write-Host $line
    Add-Content -LiteralPath $logPath -Encoding UTF8 -Value $line
}

function Get-WordOfficeVersion {
    if (-not [string]::IsNullOrWhiteSpace($OfficeVersion)) {
        return $OfficeVersion
    }

    $versionKeys = @()
    foreach ($root in @('HKCU:\Software\Microsoft\Office', 'HKLM:\Software\Microsoft\Office')) {
        $versionKeys += @(Get-ChildItem -Path $root -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
            Where-Object { Test-Path (Join-Path $_.PSPath 'Word') } |
            ForEach-Object { $_.PSChildName })
    }

    $version = $versionKeys |
        Sort-Object { [version]$_ } -Descending |
        Select-Object -First 1

    if ([string]::IsNullOrWhiteSpace($version)) {
        $version = '16.0'
    }

    return $version
}

function Ensure-RegistryKey {
    param([string]$PowerShellPath)

    if (Test-Path $PowerShellPath) {
        return
    }

    $registryPath = $PowerShellPath -replace '^HKCU:\\', 'HKCU\'
    & reg.exe add $registryPath /f | Out-Null
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $PowerShellPath)) {
        throw "Could not create registry key: $PowerShellPath"
    }
}

function Set-TemporaryVbomAccess {
    if ($SkipVbomSetup) {
        Write-Step 'Skipping AccessVBOM registry setup'
        return
    }

    $script:securityKey = "HKCU:\Software\Microsoft\Office\$(Get-WordOfficeVersion)\Word\Security"
    Write-Step "Using Word security registry key: $script:securityKey"

    try {
        Ensure-RegistryKey $script:securityKey
        $existing = Get-ItemProperty -Path $script:securityKey -Name $valueName -ErrorAction SilentlyContinue
        if ($null -ne $existing) {
            $script:hadOriginalValue = $true
            $script:originalValue = $existing.$valueName
        }

        New-ItemProperty -Path $script:securityKey -Name $valueName -PropertyType DWord -Value 1 -Force | Out-Null
    }
    catch {
        throw "Cannot update Word VBA project access setting at $script:securityKey. Actual error: $($_.Exception.Message)."
    }
}

function Restore-VbomAccess {
    if ($SkipVbomSetup -or [string]::IsNullOrWhiteSpace($script:securityKey)) {
        return
    }

    if ($script:hadOriginalValue) {
        Set-ItemProperty -Path $script:securityKey -Name $valueName -Type DWord -Value $script:originalValue -ErrorAction SilentlyContinue
    }
    else {
        Remove-ItemProperty -Path $script:securityKey -Name $valueName -ErrorAction SilentlyContinue
    }
}

function Stop-WordProcessesIfRequested {
    if (-not $CloseWord) {
        return
    }

    $wordProcesses = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
    if ($wordProcesses.Count -eq 0) {
        return
    }

    Write-Step "Closing $($wordProcesses.Count) running WINWORD process(es) because -CloseWord was specified"
    foreach ($wordProcess in $wordProcesses) {
        Stop-Process -Id $wordProcess.Id -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2
}

function Copy-FileReplacing {
    param(
        [string]$SourcePath,
        [string]$DestinationPath
    )

    $resolvedSource = (Resolve-Path -LiteralPath $SourcePath).Path
    $destinationFolder = Split-Path -Parent $DestinationPath
    if (-not [string]::IsNullOrWhiteSpace($destinationFolder)) {
        [System.IO.Directory]::CreateDirectory($destinationFolder) | Out-Null
    }
    [System.IO.File]::Copy($resolvedSource, $DestinationPath, $true)
}

function Save-DotmWithWord {
    param(
        [string]$OutputPath,
        [switch]$ImportVba
    )

    $word = $null
    $doc = $null

    try {
        Write-Step "Starting Word.Application for $OutputPath"
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $word.DisplayAlerts = 0
        $word.AutomationSecurity = 1

        Write-Step 'Creating template document'
        $doc = $word.Documents.Add()

        if ($ImportVba) {
            $doc.VBProject.Name = 'OLEPackager'
            Write-Step 'Importing Localization.bas'
            $doc.VBProject.VBComponents.Import((Join-Path $srcDir 'Localization.bas')) | Out-Null
            Write-Step 'Importing RawDataOleInserter.bas'
            $doc.VBProject.VBComponents.Import((Join-Path $srcDir 'RawDataOleInserter.bas')) | Out-Null
            Write-Step 'Importing RibbonCallbacks.bas'
            $doc.VBProject.VBComponents.Import((Join-Path $srcDir 'RibbonCallbacks.bas')) | Out-Null
            Write-Step 'Importing ImageSupportFilesDialog.frm'
            $doc.VBProject.VBComponents.Import((Join-Path $srcDir 'ImageSupportFilesDialog.frm')) | Out-Null
            Write-Step 'Importing UsageHelpDialog.frm'
            $doc.VBProject.VBComponents.Import((Join-Path $srcDir 'UsageHelpDialog.frm')) | Out-Null
        }

        Write-Step "Calling SaveAs2 for $OutputPath"
        $doc.SaveAs2($OutputPath, 15)
        Write-Step "Saved: $OutputPath"
    }
    finally {
        if ($doc -ne $null) {
            try { $doc.Close($false) } catch { Write-Step "Warning: could not close document cleanly: $($_.Exception.Message)" }
        }
        if ($word -ne $null) {
            try { $word.Quit() } catch { Write-Step "Warning: could not quit Word cleanly: $($_.Exception.Message)" }
        }
    }
}

New-Item -ItemType Directory -Force -Path $distDir | Out-Null
New-Item -ItemType Directory -Force -Path $tempBuildDir | Out-Null
Set-Content -LiteralPath $logPath -Encoding UTF8 -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Starting build"

try {
    Stop-WordProcessesIfRequested

    foreach ($path in @($plainAddinPath, $tempAddinPath)) {
        if (Test-Path -LiteralPath $path) {
            Write-Step "Removing old temporary add-in: $path"
            Remove-Item -LiteralPath $path -Force
        }
    }

    Write-Step 'Preparing Word VBA project access'
    Set-TemporaryVbomAccess
    Save-DotmWithWord -OutputPath $plainAddinPath
    Save-DotmWithWord -OutputPath $tempAddinPath -ImportVba

    Write-Step "Copying temporary dotm to dist: $addinPath"
    Copy-FileReplacing $tempAddinPath $addinPath

    Write-Step 'Injecting Ribbon customUI'
    & (Join-Path $ProjectRoot 'tools\inject-ribbon.ps1') -DotmPath $addinPath -ProjectRoot $ProjectRoot

    if ($Install) {
        if (-not (Test-Path -LiteralPath $zipToolPath)) {
            throw "Cannot install because zip helper is missing: $zipToolPath"
        }

        Write-Step "Installing to Word STARTUP after Word is closed: $startupDir"
        foreach ($legacyPath in @($legacyStartupAddinPath, $legacyStartupZipToolPath)) {
            if (Test-Path -LiteralPath $legacyPath) {
                Write-Step "Removing legacy add-in file: $legacyPath"
                Remove-Item -LiteralPath $legacyPath -Force
            }
        }
        Copy-FileReplacing $addinPath $startupAddinPath
        Copy-FileReplacing $zipToolPath $startupZipToolPath
        "Installed: $startupAddinPath"
        "Installed: $startupZipToolPath"
    }

    "Built: $addinPath"
    Write-Step 'Build completed'
}
catch {
    Write-Step "ERROR: $($_.Exception.Message)"
    throw
}
finally {
    Write-Step 'Restoring Word VBA project access setting'
    Restore-VbomAccess
}
