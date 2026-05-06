param(
    [Parameter(Mandatory = $true)]
    [string]$ScriptPath,

    [string[]]$ScriptArguments = @(),

    [int]$TimeoutSeconds = 120,

    [string]$StdoutPath = '',

    [string]$StderrPath = ''
)

$ErrorActionPreference = 'Stop'

function Quote-ProcessArgument {
    param([string]$Value)

    '"' + $Value.Replace('"', '\"') + '"'
}

function Get-WordProcessIds {
    @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
}

function Stop-NewWordProcesses {
    param([int[]]$ExistingWordIds)

    $wordProcesses = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
    foreach ($wordProcess in $wordProcesses) {
        if ($ExistingWordIds -notcontains $wordProcess.Id) {
            Stop-Process -Id $wordProcess.Id -Force -ErrorAction SilentlyContinue
        }
    }
}

$resolvedScript = (Resolve-Path $ScriptPath).Path
$scriptDir = Split-Path -Parent $resolvedScript
$scriptBase = [IO.Path]::GetFileNameWithoutExtension($resolvedScript)

if ([string]::IsNullOrWhiteSpace($StdoutPath)) {
    $StdoutPath = Join-Path $scriptDir "$scriptBase.stdout.log"
}
if ([string]::IsNullOrWhiteSpace($StderrPath)) {
    $StderrPath = Join-Path $scriptDir "$scriptBase.stderr.log"
}

foreach ($path in @($StdoutPath, $StderrPath)) {
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

$existingWordIds = Get-WordProcessIds
$arguments = @(
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    (Quote-ProcessArgument $resolvedScript)
)
foreach ($argument in $ScriptArguments) {
    $arguments += Quote-ProcessArgument $argument
}

$child = Start-Process `
    -FilePath 'powershell.exe' `
    -ArgumentList ($arguments -join ' ') `
    -RedirectStandardOutput $StdoutPath `
    -RedirectStandardError $StderrPath `
    -WindowStyle Hidden `
    -PassThru

try {
    if (-not $child.WaitForExit($TimeoutSeconds * 1000)) {
        Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue
        Stop-NewWordProcesses -ExistingWordIds $existingWordIds
        throw "Timed out after $TimeoutSeconds second(s). Killed the child PowerShell process and any Word process started by this run."
    }
    $child.Refresh()

    if (Test-Path -LiteralPath $StdoutPath) {
        Get-Content -LiteralPath $StdoutPath
    }
    if (Test-Path -LiteralPath $StderrPath) {
        $stderr = Get-Content -LiteralPath $StderrPath -Raw
        if (-not [string]::IsNullOrWhiteSpace($stderr)) {
            Write-Error $stderr
        }
    }

    if ($null -ne $child.ExitCode -and $child.ExitCode -ne 0) {
        Stop-NewWordProcesses -ExistingWordIds $existingWordIds
        throw "Script exited with code $($child.ExitCode): $resolvedScript"
    }
}
catch {
    Stop-NewWordProcesses -ExistingWordIds $existingWordIds
    throw
}
