param(
    [string]$Path = "",
    [string]$Output = "",
    [string]$Html = "",
    [string]$Settings = "",
    [switch]$NoOpen
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$SettingsPath = if ($Settings) {
    $Settings
} else {
    Join-Path $Root "settings.json"
}

function Find-Python {
    $launcher = Get-Command "py.exe" -ErrorAction SilentlyContinue
    if ($launcher -and $launcher.Source) {
        $candidate = & $launcher.Source -3 -c "import sys; print(sys.executable)" 2>$null
        if ($LASTEXITCODE -eq 0 -and $candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate.Trim()
        }
    }

    foreach ($name in @("python.exe", "python3.exe")) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if (-not $command -or -not $command.Source) {
            continue
        }
        if ($command.Source -like "*\WindowsApps\*") {
            continue
        }
        return $command.Source
    }

    throw "Python 3 was not found. Install Python 3.12 or use the released executable."
}

$Python = Find-Python
$env:PYTHONUTF8 = "1"
$env:PYTHONDONTWRITEBYTECODE = "1"

$ReportArgs = @(
    "--settings", $SettingsPath
)
if ($Output) {
    $ReportArgs += @("--output-csv", $Output)
}
if ($Html) {
    $ReportArgs += @("--output-html", $Html)
}
if ($Path) {
    $ReportArgs += @("--path", $Path)
}
if ($NoOpen) {
    $ReportArgs += "--no-open"
}

Write-Host "Kindle reader is starting..."
& $Python (Join-Path $Root "kindle_report.py") @ReportArgs
exit $LASTEXITCODE
