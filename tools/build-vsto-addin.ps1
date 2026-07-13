param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug'
)

$ErrorActionPreference = 'Stop'

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) {
    throw "vswhere.exe was not found. Install Visual Studio with Microsoft 365 development tools."
}

$vsPath = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Workload.Office -property installationPath
if ([string]::IsNullOrWhiteSpace($vsPath)) {
    throw "Visual Studio Office workload was not found. Install the Microsoft 365 development workload."
}

$msbuild = Join-Path $vsPath 'MSBuild\Current\Bin\MSBuild.exe'
if (-not (Test-Path -LiteralPath $msbuild)) {
    throw "MSBuild.exe was not found under: $vsPath"
}

function Get-OrCreate-VstoDevelopmentCertificate {
    $subject = 'CN=OLEPackager Temporary VSTO Development Certificate'
    $certificate = Get-ChildItem Cert:\CurrentUser\My |
        Where-Object { $_.Subject -eq $subject -and $_.NotAfter -gt (Get-Date).AddDays(30) } |
        Sort-Object NotBefore -Descending |
        Select-Object -First 1

    if ($null -ne $certificate) {
        return $certificate
    }

    return New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject $subject `
        -CertStoreLocation Cert:\CurrentUser\My `
        -KeyExportPolicy Exportable `
        -NotAfter (Get-Date).AddYears(3)
}

$solution = Join-Path $ProjectRoot 'OLEPackager.Vsto.sln'
$certificate = Get-OrCreate-VstoDevelopmentCertificate
& $msbuild $solution /m /restore /p:Configuration=$Configuration /p:Platform='Any CPU' /p:VisualStudioVersion=18.0 /p:ManifestCertificateThumbprint=$($certificate.Thumbprint)
if ($LASTEXITCODE -ne 0) {
    throw "VSTO build failed with exit code $LASTEXITCODE."
}
