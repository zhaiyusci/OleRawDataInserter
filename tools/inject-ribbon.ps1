param(
    [Parameter(Mandatory = $true)]
    [string]$DotmPath,

    [string]$ProjectRoot = ''
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    $ProjectRoot = (Resolve-Path (Join-Path $scriptRoot '..')).Path
}

$customUiPath = Join-Path $ProjectRoot 'customUI\customUI14.xml'
$resolvedDotm = (Resolve-Path $DotmPath).Path
$customIcons = @(
    @{ Id = 'InsertOLEPackageIcon'; File = 'insert-ole-package-icon.png' },
    @{ Id = 'InsertImageFilesIcon'; File = 'insert-image-files-icon.png' },
    @{ Id = 'ManageImageOleFilesIcon'; File = 'manage-image-ole-files-icon.png' },
    @{ Id = 'UsageHelpIcon'; File = 'usage-help-icon.png' }
)

function Add-CustomUiToWordPackage {
    param(
        [string]$PackagePath,
        [string]$CustomUiPath,
        [array]$CustomIcons,
        [string]$ProjectRoot
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::Open($PackagePath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        foreach ($entryName in @(
            'customUI/customUI14.xml',
            'customUI/_rels/customUI14.xml.rels'
        )) {
            $existingEntry = $zip.GetEntry($entryName)
            if ($null -ne $existingEntry) {
                $existingEntry.Delete()
            }
        }
        @($zip.Entries | Where-Object { $_.FullName -like 'customUI/images/*' }) | ForEach-Object { $_.Delete() }

        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $zip,
            $CustomUiPath,
            'customUI/customUI14.xml'
        ) | Out-Null

        foreach ($icon in $CustomIcons) {
            $iconPath = Join-Path (Join-Path $ProjectRoot 'assets') $icon.File
            if (-not (Test-Path -LiteralPath $iconPath)) {
                throw "Missing Ribbon icon: $iconPath"
            }
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $zip,
                $iconPath,
                "customUI/images/$($icon.File)"
            ) | Out-Null
        }

        $customUiRelsEntry = $zip.CreateEntry('customUI/_rels/customUI14.xml.rels')
        $customUiRelsWriter = New-Object System.IO.StreamWriter($customUiRelsEntry.Open())
        $relationships = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        foreach ($icon in $CustomIcons) {
            $relationships += "<Relationship Id=""$($icon.Id)"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""images/$($icon.File)""/>"
        }
        $relationships += '</Relationships>'
        $customUiRelsWriter.Write($relationships)
        $customUiRelsWriter.Close()

        $relsEntry = $zip.GetEntry('_rels/.rels')
        if ($null -eq $relsEntry) {
            throw "Could not find _rels/.rels in $PackagePath"
        }

        $relsReader = New-Object System.IO.StreamReader($relsEntry.Open())
        $relsXml = $relsReader.ReadToEnd()
        $relsReader.Close()
        $relsEntry.Delete()

        if ($relsXml -notmatch 'customUI/customUI14.xml') {
            $relationship = '<Relationship Id="rIdOLEPackagerCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
            $relsXml = $relsXml -replace '</Relationships>', "$relationship</Relationships>"
        }

        $newRelsEntry = $zip.CreateEntry('_rels/.rels')
        $relsWriter = New-Object System.IO.StreamWriter($newRelsEntry.Open())
        $relsWriter.Write($relsXml)
        $relsWriter.Close()

        $contentTypesEntry = $zip.GetEntry('[Content_Types].xml')
        if ($null -eq $contentTypesEntry) {
            throw "Could not find [Content_Types].xml in $PackagePath"
        }

        $ctReader = New-Object System.IO.StreamReader($contentTypesEntry.Open())
        $contentTypesXml = $ctReader.ReadToEnd()
        $ctReader.Close()
        $contentTypesEntry.Delete()

        if ($contentTypesXml -notmatch 'Extension="png"') {
            $pngDefault = '<Default Extension="png" ContentType="image/png"/>'
            $contentTypesXml = $contentTypesXml -replace '</Types>', "$pngDefault</Types>"
        }

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

Add-CustomUiToWordPackage -PackagePath $resolvedDotm -CustomUiPath $customUiPath -CustomIcons $customIcons -ProjectRoot $ProjectRoot
"Injected Ribbon customUI into: $resolvedDotm"
