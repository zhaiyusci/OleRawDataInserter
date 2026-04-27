$ErrorActionPreference = 'Continue'

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Add()

    'Word version: ' + $word.Version
    'Documents count: ' + $word.Documents.Count
    'Document name: ' + $doc.Name

    try {
        $vbProject = $doc.VBProject
        if ($null -eq $vbProject) {
            'VBProject: null'
        }
        else {
            'VBProject name: ' + $vbProject.Name
            'VBComponents count: ' + $vbProject.VBComponents.Count
        }
    }
    catch {
        'VBProject error: ' + $_.Exception.Message
    }

    $versions = @('16.0', '15.0', '14.0')
    foreach ($version in $versions) {
        $key = "HKCU:\Software\Microsoft\Office\$version\Word\Security"
        if (Test-Path $key) {
            $value = Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue
            "Registry $version AccessVBOM: " + $(if ($null -eq $value) { '<missing>' } else { $value.AccessVBOM })
        }
    }
}
finally {
    if ($doc -ne $null) {
        $doc.Close($false)
    }
    if ($word -ne $null) {
        $word.Quit()
    }
}
