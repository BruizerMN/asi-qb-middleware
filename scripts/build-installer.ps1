# ASI QB Middleware -- Installer Builder
#
# Run this script on the Windows workstation to compile the Inno Setup installer.
#
# Prerequisites:
#   - Inno Setup 6 installed (https://jrsoftware.org/isinfo.php)
#   - installer\bundled.env populated with real values (copy from bundled.env.example)
#
# Output: installer\Output\ASI-QB-Middleware-Setup.exe
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File scripts\build-installer.ps1

$ErrorActionPreference = "Stop"

$RepoRoot    = Split-Path -Parent $PSScriptRoot
$InstallerDir = Join-Path $RepoRoot "installer"
$BundledEnv  = Join-Path $InstallerDir "bundled.env"
$IssFile     = Join-Path $InstallerDir "setup.iss"
$OutputDir   = Join-Path $InstallerDir "Output"

$ISSCCandidates = @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe"
)

function Write-Header($text) {
    Write-Host ""
    Write-Host $text -ForegroundColor Cyan
    Write-Host ("-" * $text.Length) -ForegroundColor Cyan
}
function Write-Step($text) { Write-Host "  -> $text" -ForegroundColor Yellow }
function Write-OK($text)   { Write-Host "  OK $text" -ForegroundColor Green  }
function Abort($text) {
    Write-Host ""
    Write-Host "  FAILED: $text" -ForegroundColor Red
    Write-Host ""
    exit 1
}

Write-Host ""
Write-Host "  ASI QB Middleware -- Installer Builder" -ForegroundColor Cyan
Write-Host "  =======================================" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# Validate bundled.env
# ---------------------------------------------------------------------------

Write-Header "Checking bundled.env"

if (-not (Test-Path $BundledEnv)) {
    Abort "installer\bundled.env not found.`n`n  Copy installer\bundled.env.example to installer\bundled.env and fill in real values."
}

$envContent = Get-Content $BundledEnv -Raw

foreach ($key in @("API_KEY", "QB_FILE_ACOUSTICAL", "QB_FILE_ARCHITECTURAL")) {
    if ($envContent -notmatch "(?m)^$key\s*=\s*.+") {
        Abort "bundled.env is missing a value for: $key"
    }
}
Write-OK "bundled.env present and populated"

# ---------------------------------------------------------------------------
# Read version from version.py
# ---------------------------------------------------------------------------

Write-Header "Reading version"

$versionFile = Join-Path $RepoRoot "version.py"
if (-not (Test-Path $versionFile)) { Abort "version.py not found at repo root." }

$versionLine = Get-Content $versionFile | Where-Object { $_ -match '^VERSION\s*=' } | Select-Object -First 1
$version = ($versionLine -replace '.*=\s*["\x27](.+)["\x27].*', '$1').Trim()
if ([string]::IsNullOrWhiteSpace($version)) { Abort "Could not parse VERSION from version.py." }

Write-OK "Version: $version"

# ---------------------------------------------------------------------------
# Find Inno Setup compiler
# ---------------------------------------------------------------------------

Write-Header "Locating Inno Setup"

$iscc = $null
foreach ($candidate in $ISSCCandidates) {
    if (Test-Path $candidate) { $iscc = $candidate; break }
}
if (-not $iscc) {
    Abort "ISCC.exe not found. Install Inno Setup 6 from https://jrsoftware.org/isinfo.php"
}
Write-OK "Found: $iscc"

# ---------------------------------------------------------------------------
# Compile
# ---------------------------------------------------------------------------

Write-Header "Compiling installer"

Write-Step "Running ISCC..."
& $iscc /DMyAppVersion="$version" $IssFile
if ($LASTEXITCODE -ne 0) { Abort "ISCC compilation failed (exit code $LASTEXITCODE)." }

$outputExe = Join-Path $OutputDir "ASI-QB-Middleware-Setup.exe"
if (-not (Test-Path $outputExe)) { Abort "Compilation reported success but output file not found: $outputExe" }

$sizeMB = [math]::Round((Get-Item $outputExe).Length / 1MB, 1)

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "  =============================================" -ForegroundColor Green
Write-Host "  Installer built successfully!" -ForegroundColor Green
Write-Host "  =============================================" -ForegroundColor Green
Write-Host ""
Write-Host "  File    : installer\Output\ASI-QB-Middleware-Setup.exe" -ForegroundColor White
Write-Host "  Version : $version" -ForegroundColor White
Write-Host "  Size    : $sizeMB MB" -ForegroundColor White
Write-Host ""
Write-Host "  Distribute this file to workstations." -ForegroundColor Gray
Write-Host "  IT double-clicks to install -- no other steps needed." -ForegroundColor Gray
Write-Host ""
