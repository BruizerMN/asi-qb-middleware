# ASI QuickBooks Middleware — Windows deploy script
# Run this on the Windows QB server after deploy.sh has pushed to GitHub.
# Double-click to run, or execute from PowerShell.
#
# First-time setup: fill in the path and service name below.

$RepoPath    = "C:\Services\asi-qb-middleware"   # adjust if installed elsewhere
$ServiceName = "asi-qb-middleware"                # must match NSSM service name

Set-Location $RepoPath

Write-Host "Pulling latest from GitHub..."
git pull origin main

Write-Host "Restarting service..."
nssm restart $ServiceName

Write-Host ""
Write-Host "Deploy complete." -ForegroundColor Green
