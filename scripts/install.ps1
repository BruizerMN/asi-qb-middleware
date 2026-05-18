# ASI QuickBooks Middleware -- Workstation Installer
#
# USB / network share:
#   Right-click this file and choose "Run with PowerShell"
#   (or: powershell -ExecutionPolicy Bypass -File .\install.ps1)
#
# Safe to re-run -- updates repo, dependencies, and task registration in place.

Set-StrictMode -Off
$ErrorActionPreference = "Stop"

$ScriptBuild = "0014"

$RepoUrl  = "https://github.com/BruizerMN/asi-qb-middleware.git"
$RepoPath = "C:\Services\asi-qb-middleware"
$TaskName = "ASI QB Middleware"

$PythonInstallerUrl = "https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe"
$GitInstallerUrl    = "https://github.com/git-for-windows/git/releases/download/v2.45.2.windows.1/Git-2.45.2-64-bit.exe"

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

function Install-ViaWinget($id, $label) {
    Write-Step "Installing $label via winget..."
    winget install --id $id --silent --accept-package-agreements --accept-source-agreements
    return ($LASTEXITCODE -eq 0)
}

function Install-ViaDownload($url, $label, $silentArgs) {
    $installer = "$env:TEMP\asi-installer-$([System.Guid]::NewGuid().ToString('N'))-$([System.IO.Path]::GetFileName($url))"
    Write-Step "Downloading $label installer..."
    try {
        Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing
    } catch {
        Abort "Download failed for $label. Check internet connection and try again. ($($_.Exception.Message))"
    }
    Write-Step "Installing $label..."
    $proc = Start-Process -FilePath $installer -ArgumentList $silentArgs -Wait -PassThru
    Remove-Item $installer -ErrorAction SilentlyContinue
    if ($proc.ExitCode -ne 0) {
        Abort "$label installer exited with code $($proc.ExitCode)."
    }
}

# ---------------------------------------------------------------------------
# Banner
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "  ASI QuickBooks Middleware -- Workstation Installer" -ForegroundColor Cyan
Write-Host "  ===================================================" -ForegroundColor Cyan
Write-Host "  Installer build: $ScriptBuild" -ForegroundColor Gray

# ---------------------------------------------------------------------------
# Install Python
# ---------------------------------------------------------------------------

Write-Header "Python"

if ((Command-Exists "py") -or (Command-Exists "python")) {
    $pyver = & py --version 2>&1
    Write-OK "Already installed ($pyver)"
} else {
    $wingetOk = $false
    if (Command-Exists "winget") {
        $wingetOk = Install-ViaWinget "Python.Python.3.12" "Python 3.12"
        Refresh-Path
    }

    if (-not $wingetOk -or -not (Command-Exists "python")) {
        if ($wingetOk) { Write-Warn "winget reported success but python not found -- trying direct download..." }
        Install-ViaDownload $PythonInstallerUrl "Python 3.12.4" "/quiet InstallAllUsers=0 PrependPath=1 Include_test=0"
        Refresh-Path
    }

    if (-not (Command-Exists "python")) {
        Abort "Python installed but not found on PATH. Close this window, reopen PowerShell, and re-run the installer."
    }
    $pyver = & python --version 2>&1
    Write-OK "Installed ($pyver)"
}

# ---------------------------------------------------------------------------
# Install Git
# ---------------------------------------------------------------------------

Write-Header "Git"

if (Command-Exists "git") {
    $gitver = & git --version 2>&1
    Write-OK "Already installed ($gitver)"
} else {
    $wingetOk = $false
    if (Command-Exists "winget") {
        $wingetOk = Install-ViaWinget "Git.Git" "Git"
        Refresh-Path
    }

    if (-not $wingetOk -or -not (Command-Exists "git")) {
        if ($wingetOk) { Write-Warn "winget reported success but git not found -- trying direct download..." }
        Install-ViaDownload $GitInstallerUrl "Git 2.45.2" "/VERYSILENT /NORESTART /NOCANCEL /SP- /CLOSEAPPLICATIONS"
        Refresh-Path
    }

    if (-not (Command-Exists "git")) {
        Abort "Git installed but not found on PATH. Close this window, reopen PowerShell, and re-run the installer."
    }
    $gitver = & git --version 2>&1
    Write-OK "Installed ($gitver)"
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
    if ($LASTEXITCODE -ne 0) { Abort "git clone failed. Make sure you have network access to GitHub." }
    Set-Location $RepoPath
    Write-OK "Repo cloned"
}

# ---------------------------------------------------------------------------
# Python dependencies
# ---------------------------------------------------------------------------

Write-Header "Python dependencies"

Write-Step "Running pip install..."
& py -m pip install --quiet -r "$RepoPath\requirements.txt"
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

    $port = "5100"

    $envContent = "API_KEY=" + $apiKey + "`r`nPORT=" + $port + "`r`n"
    $utf8NoBom = New-Object System.Text.UTF8Encoding $False
    [System.IO.File]::WriteAllText($EnvFile, $envContent, $utf8NoBom)

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
    -Execute "py" `
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

$maxAttempts = 8
$attempt     = 0
$running     = $false

while ($attempt -lt $maxAttempts -and -not $running) {
    $attempt++
    Start-Sleep -Seconds 2
    Write-Step "Health check (attempt $attempt of $maxAttempts)..."
    try {
        $resp    = Invoke-RestMethod -Uri "http://localhost:$portVal/health" -TimeoutSec 3
        $running = $true
    } catch {
        # still starting
    }
}

if ($running) {
    Write-OK "Middleware is running!"
    Write-Host ""
    Write-Host "      Version : $($resp.version)" -ForegroundColor White
    Write-Host "      Build   : $($resp.build)"   -ForegroundColor White
} else {
    Write-Warn "Middleware did not respond after $($maxAttempts * 2) seconds."
    Write-Host "    Check Task Scheduler for errors, or run manually:" -ForegroundColor Gray
    Write-Host "    python $RepoPath\app.py" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    To test once running: Invoke-RestMethod http://localhost:$portVal/health" -ForegroundColor Gray
}

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "  All done! The middleware will start automatically at each logon." -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
