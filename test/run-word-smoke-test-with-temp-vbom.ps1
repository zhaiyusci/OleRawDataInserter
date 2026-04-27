$ErrorActionPreference = 'Stop'

$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security'
$valueName = 'AccessVBOM'
$hadOriginalValue = $false
$originalValue = $null

try {
    if (-not (Test-Path $securityKey)) {
        New-Item -Path $securityKey -Force | Out-Null
    }

    $existing = Get-ItemProperty -Path $securityKey -Name $valueName -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
        $hadOriginalValue = $true
        $originalValue = $existing.$valueName
    }

    Set-ItemProperty -Path $securityKey -Name $valueName -Type DWord -Value 1

    & (Join-Path $PSScriptRoot 'run-word-smoke-test.ps1')
}
finally {
    if ($hadOriginalValue) {
        Set-ItemProperty -Path $securityKey -Name $valueName -Type DWord -Value $originalValue
    }
    else {
        Remove-ItemProperty -Path $securityKey -Name $valueName -ErrorAction SilentlyContinue
    }
}
