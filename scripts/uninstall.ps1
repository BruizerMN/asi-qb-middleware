# ASI QuickBooks Middleware -- Uninstaller
#
# Always removes:
#   - The "ASI QB Middleware" Task Scheduler task
#   - The middleware directory (C:\Services\asi-qb-middleware\)
#
# Optionally removes Python and Git -- use flags to skip the interactive prompts:
#   -RemovePython   Uninstall Python without prompting
#   -RemoveGit      Uninstall Git without prompting
#
# Full clean sweep (e.g. resetting a dev VM):
#   powershell -ExecutionPolicy Bypass -File uninstall.ps1 -RemovePython -RemoveGit
#
# Workstation removal (keep Python/Git for other uses):
#   powershell -ExecutionPolicy Bypass -File uninstall.ps1

param(
    [switch]$RemovePython,
    [switch]$RemoveGit
)

Set-StrictMode -Off
$ErrorActionPreference = "SilentlyContinue"

$RepoPath = "C:\Services\asi-qb-middleware"
$TaskName = "ASI QB Middleware"

function Write-Header($text) {
    Write-Host ""
    Write-Host $text -ForegroundColor Cyan
    Write-Host ("-" * $text.Length) -ForegroundColor Cyan
}
function Write-Step($text) { Write-Host "  -> $text" -ForegroundColor Yellow }
function Write-OK($text)   { Write-Host "  OK $text" -ForegroundColor Green  }
function Write-Warn($text) { Write-Host "  !! $text" -ForegroundColor Yellow }

function Prompt-YesNo($question) {
    Write-Host ""
    Write-Host "  $question" -ForegroundColor White
    $answer = Read-Host "  Remove? [y/N]"
    return ($answer -match "^[Yy]")
}

function Uninstall-ViaMsi($displayName) {
    # Search user and machine uninstall registry keys for a matching display name.
    $regPaths = @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    foreach ($regPath in $regPaths) {
        if (-not (Test-Path $regPath)) { continue }
        $entries = Get-ChildItem $regPath | Get-ItemProperty |
            Where-Object { $_.DisplayName -match $displayName }
        foreach ($entry in $entries) {
            $uninstStr = $entry.UninstallString
            if (-not $uninstStr) { continue }
            Write-Step "Running uninstaller for: $($entry.DisplayName)"
            # MsiExec-based uninstall
            if ($uninstStr -match "MsiExec") {
                $productCode = ($uninstStr -replace ".*(/X|/I)(\{[^}]+\}).*", '$2')
                Start-Process "msiexec.exe" -ArgumentList "/X $productCode /quiet /norestart" -Wait
            } else {
                # EXE-based uninstall -- append silent flags
                $exePath = $uninstStr -replace '"', '' -replace ' /.*', ''
                $args    = "/uninstall /quiet /norestart"
                if (Test-Path $exePath) {
                    Start-Process -FilePath $exePath -ArgumentList $args -Wait
                }
            }
            return $true
        }
    }
    return $false
}

# ---------------------------------------------------------------------------
# Banner
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "  ASI QB Middleware -- Uninstaller" -ForegroundColor Cyan
Write-Host "  ==================================" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# Stop and remove Task Scheduler task
# ---------------------------------------------------------------------------

Write-Header "Task Scheduler"

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
    Write-Step "Stopping task..."
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    Write-Step "Killing any remaining Python processes..."
    Get-Process python* -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1

    Write-Step "Removing task..."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-OK "Task removed"
} else {
    Write-OK "Task not found -- already removed or never installed"
}

# ---------------------------------------------------------------------------
# Remove middleware directory
# ---------------------------------------------------------------------------

Write-Header "Middleware files"

if (Test-Path $RepoPath) {
    Write-Step "Deleting $RepoPath ..."
    Remove-Item -Recurse -Force $RepoPath
    if (Test-Path $RepoPath) {
        Write-Warn "Could not fully delete $RepoPath -- some files may be locked. Try again after a reboot."
    } else {
        Write-OK "Middleware directory removed"
    }
} else {
    Write-OK "Directory not found -- already removed or never installed"
}

# ---------------------------------------------------------------------------
# Python
# ---------------------------------------------------------------------------

Write-Header "Python"

$doPython = $RemovePython -or (Prompt-YesNo "Remove Python? (Safe to keep if used for other tools on this machine.)")

if ($doPython) {
    $removed = $false

    # Try winget first
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Step "Uninstalling via winget..."
        foreach ($id in @("Python.Python.3.13", "Python.Python.3.12", "Python.Python.3.11")) {
            winget uninstall --id $id --silent --accept-source-agreements 2>$null
            if ($LASTEXITCODE -eq 0) { $removed = $true }
        }
    }

    # Fall back to registry-based uninstall
    if (-not $removed) {
        Write-Step "Trying registry uninstall..."
        $removed = Uninstall-ViaMsi "^Python 3\."
    }

    if ($removed) {
        Write-OK "Python removed"
    } else {
        Write-Warn "Python not found or could not be uninstalled automatically."
        Write-Host "    To remove manually: Settings > Apps > search 'Python' > Uninstall" -ForegroundColor Gray
    }
} else {
    Write-OK "Python kept"
}

# ---------------------------------------------------------------------------
# Git
# ---------------------------------------------------------------------------

Write-Header "Git"

$doGit = $RemoveGit -or (Prompt-YesNo "Remove Git? (Safe to keep if used for other development on this machine.)")

if ($doGit) {
    $removed = $false

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Step "Uninstalling via winget..."
        winget uninstall --id Git.Git --silent --accept-source-agreements 2>$null
        if ($LASTEXITCODE -eq 0) { $removed = $true }
    }

    if (-not $removed) {
        Write-Step "Trying registry uninstall..."
        $removed = Uninstall-ViaMsi "^Git"
    }

    if ($removed) {
        Write-OK "Git removed"
    } else {
        Write-Warn "Git not found or could not be uninstalled automatically."
        Write-Host "    To remove manually: Settings > Apps > search 'Git' > Uninstall" -ForegroundColor Gray
    }
} else {
    Write-OK "Git kept"
}

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "  Uninstall complete." -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
