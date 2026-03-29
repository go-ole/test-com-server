# Register TestCOMServer COM classes using regsvr32
# Run as Administrator
# Usage: .\register.ps1 [-Unregister]

param(
    [switch]$Unregister
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$comhost = Join-Path $scriptDir "publish\TestCOMServer.comhost.dll"

if (-not (Test-Path $comhost)) {
    Write-Error "comhost.dll not found at $comhost. Run 'dotnet publish' first."
    exit 1
}

if ($Unregister) {
    Write-Host "Unregistering TestCOMServer..."
    regsvr32 /u /s $comhost
} else {
    Write-Host "Registering TestCOMServer..."
    regsvr32 /s $comhost
}

if ($LASTEXITCODE -eq 0) {
    Write-Host "Done."
} else {
    Write-Error "Registration failed with exit code $LASTEXITCODE"
}
