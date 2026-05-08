param(
    [int]$TimeoutSeconds = 120,
    [string]$OfficeVersion = ''
)

$ErrorActionPreference = 'Stop'

$securityKey = $null
$valueName = 'AccessVBOM'
$hadOriginalValue = $false
$originalValue = $null
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$timeoutRunner = Join-Path $projectRoot 'tools\run-word-script-with-timeout.ps1'
$smokeTestScript = Join-Path $PSScriptRoot 'run-word-smoke-test.ps1'

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

try {
    $securityKey = "HKCU:\Software\Microsoft\Office\$(Get-WordOfficeVersion)\Word\Security"
    Ensure-RegistryKey $securityKey

    $existing = Get-ItemProperty -Path $securityKey -Name $valueName -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
        $hadOriginalValue = $true
        $originalValue = $existing.$valueName
    }

    New-ItemProperty -Path $securityKey -Name $valueName -PropertyType DWord -Value 1 -Force | Out-Null

    & $timeoutRunner -ScriptPath $smokeTestScript -TimeoutSeconds $TimeoutSeconds
}
finally {
    if (-not [string]::IsNullOrWhiteSpace($securityKey)) {
        if ($hadOriginalValue) {
            Set-ItemProperty -Path $securityKey -Name $valueName -Type DWord -Value $originalValue
        }
        else {
            Remove-ItemProperty -Path $securityKey -Name $valueName -ErrorAction SilentlyContinue
        }
    }
}
