[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$ConfigPath,

    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"
$scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) "src\hanger_automation.py"
$arguments = @($scriptPath, "run", "--input", $InputPath, "--config", $ConfigPath)
if ($Overwrite) {
    $arguments += "--overwrite"
}

& python @arguments
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

