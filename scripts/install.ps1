# ASI QuickBooks Middleware -- Workstation Installer
#
# USB / network share:
#   Right-click this file and choose "Run with PowerShell"
#   (or: powershell -ExecutionPolicy Bypass -File .\install.ps1)
#
# Safe to re-run -- updates repo, dependencies, and task registration in place.

Set-StrictMode -Off
$ErrorActionPreference = "Stop"

$RepoUrl  = "https://github.com/BruizerMN/asi-qb-middleware.git"
$RepoPath = "C:\Services\asi-qb-middleware"
$TaskName = "ASI QB Middleware"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Write-Header($text) {
    Write-Host ""
    Write-Host $text -ForegroundColor Cyan
    Write-Host ("-" * $text.Length) -ForegroundColor Cyan
}

function Write-Step($text)  { Write-Host "  -> $text" -ForegroundColor Yellow }
function Write-OK($text)    { Write-Host "  OK $text" -ForegroundColor Green  }
function Write-Warn($text)  { Write-Host "  !! $text" -ForegroundColor Yellow }

function Abort($text) {
    Write-Host ""
    Write-Host "  FAILED: $text" -ForegroundColor Red
    Write-Host ""
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

function Refresh-Path {
    $machine = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $user    = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = $machine + ";" + $user
}

function Command-Exists($cmd) {
    return [bool](Get-Command $cmd -ErrorAction SilentlyContinue)
}

# ---------------------------------------------------------------------------
# Banner
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "  ASI QuickBooks Middleware -- Workstation Installer" -ForegroundColor Cyan
Write-Host "  ===================================================" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# Check winget
# ---------------------------------------------------------------------------

Write-Header "Checking prerequisites"

if (-not (Command-Exists "winget")) {
    Abort "winget not found. Update Windows (Win 10 1709+ or Win 11 required) or install App Installer from the Microsoft Store."
}
Write-OK "winget available"

# ---------------------------------------------------------------------------
# Install Python
# ---------------------------------------------------------------------------

if (Command-Exists "python") {
    $pyver = & python --version 2>&1
    Write-OK "Python already installed ($pyver)"
} else {
    Write-Step "Installing Python 3.12..."
    winget install --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) { Abort "Python install failed (winget exit $LASTEXITCODE)." }
    Refresh-Path
    if (-not (Command-Exists "python")) {
        Abort "Python installed but not yet on PATH. Close this window, reopen PowerShell, and re-run the installer."
    }
    $pyver = & python --version 2>&1
    Write-OK "Python installed ($pyver)"
}

# ---------------------------------------------------------------------------
# Install Git
# ---------------------------------------------------------------------------

if (Command-Exists "git") {
    $gitver = & git --version 2>&1
    Write-OK "Git already installed ($gitver)"
} else {
    Write-Step "Installing Git..."
    winget install --id Git.Git --silent --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) { Abort "Git install failed (winget exit $LASTEXITCODE)." }
    Refresh-Path
    if (-not (Command-Exists "git")) {
        Abort "Git installed but not yet on PATH. Close this window, reopen PowerShell, and re-run the installer."
    }
    $gitver = & git --version 2>&1
    Write-OK "Git installed ($gitver)"
}

# ---------------------------------------------------------------------------
# Clone or update repo
# ---------------------------------------------------------------------------

Write-Header "Middleware code"

if (Test-Path "$RepoPath\.git") {
    Write-Step "Pulling latest from GitHub..."
    Set-Location $RepoPath
    git pull origin main
    if ($LASTEXITCODE -ne 0) { Abort "git pull failed." }
    Write-OK "Repo updated"
} else {
    Write-Step "Cloning repo to $RepoPath ..."
    New-Item -ItemType Directory -Force -Path $RepoPath | Out-Null
    git clone $RepoUrl $RepoPath
    if ($LASTEXITCODE -ne 0) { Abort "git clone failed." }
    Set-Location $RepoPath
    Write-OK "Repo cloned"
}

# ---------------------------------------------------------------------------
# Python dependencies
# ---------------------------------------------------------------------------

Write-Header "Python dependencies"

Write-Step "Running pip install..."
& python -m pip install --quiet -r "$RepoPath\requirements.txt"
if ($LASTEXITCODE -ne 0) { Abort "pip install failed." }
Write-OK "Dependencies installed"

# ---------------------------------------------------------------------------
# Create .env if missing
# ---------------------------------------------------------------------------

$EnvFile = "$RepoPath\.env"

Write-Header "Configuration"

if (Test-Path $EnvFile) {
    Write-OK ".env already exists -- skipping"
} else {
    Write-Host ""
    Write-Host "  Enter the API key for this middleware." -ForegroundColor White
    Write-Host "  This must match the QB_APIKey value in FileMaker's Reference table." -ForegroundColor Gray
    Write-Host ""
    $apiKey = Read-Host "  API_KEY"
    if ([string]::IsNullOrWhiteSpace($apiKey)) { Abort "API_KEY cannot be blank." }

    Write-Host ""
    Write-Host "  Port for the middleware (press Enter for default 5100):" -ForegroundColor White
    $port = Read-Host "  PORT"
    if ([string]::IsNullOrWhiteSpace($port)) { $port = "5100" }

    $envContent = "API_KEY=" + $apiKey + "`r`nPORT=" + $port + "`r`n"
    [System.IO.File]::WriteAllText($EnvFile, $envContent, [System.Text.Encoding]::UTF8)

    Write-OK ".env created"
}

# Read port for health check
$portVal = "5100"
$lines = Get-Content $EnvFile -ErrorAction SilentlyContinue
foreach ($line in $lines) {
    if ($line -match "^PORT=(.+)") { $portVal = $Matches[1].Trim() }
}

# ---------------------------------------------------------------------------
# Task Scheduler
# ---------------------------------------------------------------------------
# Must run as the current user (not SYSTEM) so it shares the Windows session
# with QuickBooks Desktop -- required for COM/QBFC to work.

Write-Header "Task Scheduler"

$action = New-ScheduledTaskAction `
    -Execute "python" `
    -Argument "$RepoPath\app.py" `
    -WorkingDirectory $RepoPath

$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 0) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -StartWhenAvailable

$principal = New-ScheduledTaskPrincipal `
    -UserId $env:USERNAME `
    -LogonType Interactive `
    -RunLevel Limited

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if ($existing) {
    Write-Step "Updating existing task..."
    Set-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal | Out-Null
    Write-OK "Task updated"
} else {
    Write-Step "Registering new task..."
    Register-ScheduledTask `
        -TaskName  $TaskName `
        -Action    $action `
        -Trigger   $trigger `
        -Settings  $settings `
        -Principal $principal `
        -Force | Out-Null
    Write-OK "Task registered -- middleware will start automatically at each logon"
}

# ---------------------------------------------------------------------------
# Start now
# ---------------------------------------------------------------------------

Write-Header "Starting middleware"

Write-Step "Starting task..."
Start-ScheduledTask -TaskName $TaskName
Start-Sleep -Seconds 4

try {
    $resp = Invoke-RestMethod -Uri "http://localhost:$portVal/health" -TimeoutSec 5
    Write-OK "Middleware is running!"
    Write-Host ""
    Write-Host "      Version : $($resp.version)" -ForegroundColor White
    Write-Host "      Build   : $($resp.build)"   -ForegroundColor White
} catch {
    Write-Warn "Health check timed out -- middleware may still be starting."
    Write-Host "    Check manually: Invoke-RestMethod http://localhost:$portVal/health" -ForegroundColor Gray
}

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "  All done! The middleware will start automatically at each logon." -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
