# ASI QuickBooks Middleware — Workstation Setup / Update Script
# Run this on each FM+QB user workstation. Safe to re-run for updates.
# Must be run as the logged-in user (NOT as Administrator) so the Task Scheduler
# logon task runs in the correct user context for COM/QBFC access to QuickBooks.
#
# Usage:
#   First time:  right-click → Run with PowerShell
#   Updates:     double-click, or from PowerShell prompt

$RepoPath    = "C:\Services\asi-qb-middleware"
$TaskName    = "ASI QB Middleware"
$PythonExe   = "python"
$EnvFile     = "$RepoPath\.env"

# ── Helpers ──────────────────────────────────────────────────────────────────

function Check-Command($cmd) {
    return [bool](Get-Command $cmd -ErrorAction SilentlyContinue)
}

function Abort($msg) {
    Write-Host ""
    Write-Host "ERROR: $msg" -ForegroundColor Red
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# ── Preflight ─────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "ASI QB Middleware — Workstation Setup" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Check-Command $PythonExe)) {
    Abort "Python not found. Install Python 3.10+ from python.org and re-run this script."
}

if (-not (Check-Command "git")) {
    Abort "Git not found. Install Git from git-scm.com and re-run this script."
}

# ── Clone or update repo ──────────────────────────────────────────────────────

if (Test-Path "$RepoPath\.git") {
    Write-Host "Repo exists — pulling latest..." -ForegroundColor Yellow
    Set-Location $RepoPath
    git pull origin main
} else {
    Write-Host "Cloning repo to $RepoPath ..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Force -Path $RepoPath | Out-Null
    git clone https://github.com/BruizerMN/asi-qb-middleware.git $RepoPath
    Set-Location $RepoPath
}

# ── Install / update dependencies ────────────────────────────────────────────

Write-Host ""
Write-Host "Installing Python dependencies..." -ForegroundColor Yellow
& $PythonExe -m pip install --quiet -r requirements.txt
if ($LASTEXITCODE -ne 0) { Abort "pip install failed." }

# ── Create .env if missing ───────────────────────────────────────────────────

if (-not (Test-Path $EnvFile)) {
    Write-Host ""
    Write-Host "Creating .env file..." -ForegroundColor Yellow

    $apiKey = Read-Host "Enter API_KEY (must match FileMaker's QB_APIKey field)"
    $port   = Read-Host "Enter PORT [default: 5100]"
    if ([string]::IsNullOrWhiteSpace($port)) { $port = "5100" }

    @"
API_KEY=$apiKey
PORT=$port
"@ | Set-Content $EnvFile -Encoding UTF8

    Write-Host ".env created at $EnvFile" -ForegroundColor Green
} else {
    Write-Host ".env already exists — skipping." -ForegroundColor Gray
}

# ── Register Task Scheduler logon task ───────────────────────────────────────
# Runs as the current user at logon — required for COM/QBFC access to QuickBooks.
# (A Windows service runs as LocalSystem and cannot see user-space COM objects.)

Write-Host ""
Write-Host "Registering Task Scheduler logon task..." -ForegroundColor Yellow

$taskExists = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

$action  = New-ScheduledTaskAction `
    -Execute $PythonExe `
    -Argument "$RepoPath\app.py" `
    -WorkingDirectory $RepoPath

$trigger = New-ScheduledTaskTrigger -AtLogOn

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 0) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -StartWhenAvailable

if ($taskExists) {
    Set-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings | Out-Null
    Write-Host "Task updated." -ForegroundColor Green
} else {
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -RunLevel Limited `
        -Force | Out-Null
    Write-Host "Task registered." -ForegroundColor Green
}

# ── Start the middleware now (don't wait for next logon) ──────────────────────

Write-Host ""
Write-Host "Starting middleware..." -ForegroundColor Yellow
Start-ScheduledTask -TaskName $TaskName

Start-Sleep -Seconds 3

# Quick health check
try {
    $port = (Get-Content $EnvFile | Select-String "^PORT=").ToString().Split("=")[1].Trim()
    if ([string]::IsNullOrWhiteSpace($port)) { $port = "5100" }
    $resp = Invoke-RestMethod -Uri "http://localhost:$port/health" -TimeoutSec 5
    Write-Host ""
    Write-Host "Middleware is running!" -ForegroundColor Green
    Write-Host "  Version: $($resp.version)  Build: $($resp.build)" -ForegroundColor Green
} catch {
    Write-Host ""
    Write-Host "WARNING: Health check failed. The middleware may still be starting." -ForegroundColor Yellow
    Write-Host "Check Task Scheduler or run: python $RepoPath\app.py" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Setup complete. The middleware will start automatically at each logon." -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
