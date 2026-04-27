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
$customIconPath = Join-Path $ProjectRoot 'assets\figure-package-icon.png'
$resolvedDotm = (Resolve-Path $DotmPath).Path

function Add-CustomUiToWordPackage {
    param(
        [string]$PackagePath,
        [string]$CustomUiPath,
        [string]$CustomIconPath
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::Open($PackagePath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        foreach ($entryName in @(
            'customUI/customUI14.xml',
            'customUI/_rels/customUI14.xml.rels',
            'customUI/images/figure-package-icon.png'
        )) {
            $existingEntry = $zip.GetEntry($entryName)
            if ($null -ne $existingEntry) {
                $existingEntry.Delete()
            }
        }

        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $zip,
            $CustomUiPath,
            'customUI/customUI14.xml'
        ) | Out-Null

        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $zip,
            $CustomIconPath,
            'customUI/images/figure-package-icon.png'
        ) | Out-Null

        $customUiRelsEntry = $zip.CreateEntry('customUI/_rels/customUI14.xml.rels')
        $customUiRelsWriter = New-Object System.IO.StreamWriter($customUiRelsEntry.Open())
        $customUiRelsWriter.Write('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="FigurePackageIcon" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/figure-package-icon.png"/></Relationships>')
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
            $relationship = '<Relationship Id="rIdOleRawDataInserterCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
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

Add-CustomUiToWordPackage -PackagePath $resolvedDotm -CustomUiPath $customUiPath -CustomIconPath $customIconPath
"Injected Ribbon customUI into: $resolvedDotm"
