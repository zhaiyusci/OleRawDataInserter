$ErrorActionPreference = 'Stop'

$outDir = Join-Path $env:LOCALAPPDATA 'Temp\OleRawDataInserterBuild'
$plainDotm = Join-Path $outDir 'PlainSaveTest.dotm'
$vbaDotm = Join-Path $outDir 'VbaSaveTest.dotm'
$projectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$srcDir = Join-Path $projectRoot 'src'

New-Item -ItemType Directory -Force -Path $outDir | Out-Null
foreach ($path in @($plainDotm, $vbaDotm)) {
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

function Test-SaveDotm {
    param(
        [string]$OutputPath,
        [switch]$ImportVba
    )

    $word = $null
    $doc = $null
    try {
        "Starting Word for $OutputPath"
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $word.DisplayAlerts = 0
        $word.AutomationSecurity = 1

        $doc = $word.Documents.Add()
        if ($ImportVba) {
            "Importing VBA modules"
            $doc.VBProject.Name = 'OleRawDataInserter'
            $doc.VBProject.VBComponents.Import((Join-Path $srcDir 'RawDataOleInserter.bas')) | Out-Null
            $doc.VBProject.VBComponents.Import((Join-Path $srcDir 'RibbonCallbacks.bas')) | Out-Null
            $doc.VBProject.VBComponents.Import((Join-Path $srcDir 'ImageSupportFilesDialog.frm')) | Out-Null
        }

        "Calling SaveAs2"
        $doc.SaveAs2($OutputPath, 15)
        "Saved: $OutputPath"
    }
    finally {
        if ($doc -ne $null) {
            try {
                $doc.Close($false)
            }
            catch {
                "Warning: could not close document cleanly: $($_.Exception.Message)"
            }
        }
        if ($word -ne $null) {
            try {
                $word.Quit()
            }
            catch {
                "Warning: could not quit Word cleanly: $($_.Exception.Message)"
            }
        }
    }
}

Test-SaveDotm -OutputPath $plainDotm
Test-SaveDotm -OutputPath $vbaDotm -ImportVba
