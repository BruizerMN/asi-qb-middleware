#!/bin/bash
# Deploy to production Windows QB server.
# Run this after committing: the commit itself bumps the build number.
#
# Deployment is a two-step process:
#   1. This script pushes to GitHub (automated, runs on Mac)
#   2. You RDP into the Windows server and run deploy-windows.ps1 (one double-click)

set -e

echo "Pushing to GitHub..."
git push origin main

echo ""
echo "✅ Push complete."
echo ""
echo "To finish deploying:"
echo "  1. RDP into the Windows QB server"
echo "  2. Run deploy-windows.ps1 (double-click, or from PowerShell)"
echo ""
