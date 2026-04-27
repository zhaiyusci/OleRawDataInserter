param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [switch]$Install,
    [switch]$SkipVbomSetup
)

$ErrorActionPreference = 'Stop'

$srcDir = Join-Path $ProjectRoot 'src'
$customUiPath = Join-Path $ProjectRoot 'customUI\customUI14.xml'
$distDir = Join-Path $ProjectRoot 'dist'
$addinPath = Join-Path $distDir 'OleRawDataInserter.dotm'
$tempBuildDir = Join-Path $env:TEMP 'OleRawDataInserterBuild'
$tempAddinPath = Join-Path $tempBuildDir 'OleRawDataInserter.dotm'
$logPath = Join-Path $distDir 'build-word-addin.log'
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security'
$valueName = 'AccessVBOM'
$hadOriginalValue = $false
$originalValue = $null

function Write-Step {
    param([string]$Message)

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Write-Host $line
    Add-Content -LiteralPath $logPath -Encoding UTF8 -Value $line
}

function Set-TemporaryVbomAccess {
    if ($SkipVbomSetup) {
        Write-Step 'Skipping AccessVBOM registry setup'
        return
    }

    if (-not (Test-Path $securityKey)) {
        try {
            New-Item -Path $securityKey -Force | Out-Null
        }
        catch {
            throw "Cannot update Word VBA project access setting. Open Word -> File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings, enable 'Trust access to the VBA project object model', then rerun this script with -SkipVbomSetup."
        }
    }

    $existing = Get-ItemProperty -Path $securityKey -Name $valueName -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
        $script:hadOriginalValue = $true
        $script:originalValue = $existing.$valueName
    }

    try {
        Set-ItemProperty -Path $securityKey -Name $valueName -Type DWord -Value 1
    }
    catch {
        throw "Cannot update Word VBA project access setting. Open Word -> File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings, enable 'Trust access to the VBA project object model', then rerun this script with -SkipVbomSetup."
    }
}

function Restore-VbomAccess {
    if ($SkipVbomSetup) {
        return
    }

    if ($script:hadOriginalValue) {
        Set-ItemProperty -Path $securityKey -Name $valueName -Type DWord -Value $script:originalValue -ErrorAction SilentlyContinue
    }
    else {
        Remove-ItemProperty -Path $securityKey -Name $valueName -ErrorAction SilentlyContinue
    }
}

function Add-CustomUiToWordPackage {
    param(
        [string]$PackagePath,
        [string]$CustomUiPath
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::Open($PackagePath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        foreach ($entryName in @('customUI/customUI14.xml')) {
            $existing = $zip.GetEntry($entryName)
            if ($null -ne $existing) {
                $existing.Delete()
            }
        }

        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $zip,
            $CustomUiPath,
            'customUI/customUI14.xml'
        ) | Out-Null

        $relsEntry = $zip.GetEntry('_rels/.rels')
        $relsReader = New-Object System.IO.StreamReader($relsEntry.Open())
        $relsXml = $relsReader.ReadToEnd()
        $relsReader.Close()
        $relsEntry.Delete()

        if ($relsXml -notmatch 'customUI/customUI14.xml') {
            $relationship = '<Relationship Id="rIdOleRawDataInserterCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
            $relsXml = $relsXml -replace '</Relationships>', "$relationship</Relationships>"
        }

        $newRelsEntry = $zip.CreateEntry('_rels/.rels')
        $relsWriter = New-Object System.IO.StreamWriter($newRelsEntry.Open())
        $relsWriter.Write($relsXml)
        $relsWriter.Close()

        $contentTypesEntry = $zip.GetEntry('[Content_Types].xml')
        $ctReader = New-Object System.IO.StreamReader($contentTypesEntry.Open())
        $contentTypesXml = $ctReader.ReadToEnd()
        $ctReader.Close()
        $contentTypesEntry.Delete()

        if ($contentTypesXml -notmatch '/customUI/customUI14.xml') {
            $override = '<Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/>'
            $contentTypesXml = $contentTypesXml -replace '</Types>', "$override</Types>"
        }

        $newContentTypesEntry = $zip.CreateEntry('[Content_Types].xml')
        $ctWriter = New-Object System.IO.StreamWriter($newContentTypesEntry.Open())
        $ctWriter.Write($contentTypesXml)
        $ctWriter.Close()
    }
    finally {
        $zip.Dispose()
    }
}

New-Item -ItemType Directory -Force -Path $distDir | Out-Null
New-Item -ItemType Directory -Force -Path $tempBuildDir | Out-Null
Set-Content -LiteralPath $logPath -Encoding UTF8 -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Starting build"

if (Test-Path $addinPath) {
    Write-Step "Removing old add-in: $addinPath"
    Remove-Item -LiteralPath $addinPath -Force
}
if (Test-Path $tempAddinPath) {
    Write-Step "Removing old temporary add-in: $tempAddinPath"
    Remove-Item -LiteralPath $tempAddinPath -Force
}

$word = $null
$doc = $null

try {
    Write-Step 'Preparing Word VBA project access'
    Set-TemporaryVbomAccess

    Write-Step 'Starting Word.Application'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 1

    Write-Step 'Creating template document'
    $doc = $word.Documents.Add()
    $doc.VBProject.Name = 'OleRawDataInserter'
    Write-Step 'Importing RawDataOleInserter.bas'
    $doc.VBProject.VBComponents.Import((Join-Path $srcDir 'RawDataOleInserter.bas')) | Out-Null
    Write-Step 'Importing RibbonCallbacks.bas'
    $doc.VBProject.VBComponents.Import((Join-Path $srcDir 'RibbonCallbacks.bas')) | Out-Null
    Write-Step "Saving dotm to temporary ASCII path: $tempAddinPath"
    $doc.SaveAs2($tempAddinPath, 15)
    $doc.Close($false)
    $doc = $null

    Write-Step 'Injecting Ribbon customUI'
    Add-CustomUiToWordPackage -PackagePath $tempAddinPath -CustomUiPath $customUiPath

    Write-Step "Copying built add-in to dist: $addinPath"
    Copy-Item -LiteralPath $tempAddinPath -Destination $addinPath -Force

    if ($Install) {
        $startupDir = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
        Write-Step "Installing to Word STARTUP: $startupDir"
        New-Item -ItemType Directory -Force -Path $startupDir | Out-Null
        Copy-Item -LiteralPath $addinPath -Destination (Join-Path $startupDir 'OleRawDataInserter.dotm') -Force
        "Installed: $(Join-Path $startupDir 'OleRawDataInserter.dotm')"
    }

    "Built: $addinPath"
    Write-Step 'Build completed'
}
catch {
    Write-Step "ERROR: $($_.Exception.Message)"
    throw
}
finally {
    if ($doc -ne $null) {
        $doc.Close($false)
    }
    if ($word -ne $null) {
        $word.Quit()
    }
    Write-Step 'Restoring Word VBA project access setting'
    Restore-VbomAccess
}
