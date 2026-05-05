#!/bin/bash
# Deploy to production Windows QB server.
# Run this after committing: the commit itself bumps the build number.
#
# TODO (once VPN + server are confirmed):
#   - Set PROD_HOST to the Windows server hostname/IP
#   - Set PROD_USER to the SSH username
#   - Set PROD_PATH to the repo path on the Windows server
#   - Set SERVICE_NAME to the NSSM service name
#
# When complete this script will:
#   1. Push current main to GitHub
#   2. SSH into the Windows server, git pull, and restart the NSSM service

set -e

PROD_HOST=""        # e.g. 192.168.1.50
PROD_USER=""        # e.g. administrator
PROD_PATH=""        # e.g. C:/Services/asi-qb-middleware  (use forward slashes for ssh)
SERVICE_NAME=""     # e.g. asi-qb-middleware

echo "Pushing to GitHub..."
git push origin main

if [ -z "$PROD_HOST" ]; then
    echo ""
    echo "⚠️  Windows server not yet configured in deploy.sh."
    echo "   Push complete. Log into the server and run:"
    echo "   cd $PROD_PATH && git pull && nssm restart $SERVICE_NAME"
    exit 0
fi

echo "Deploying to $PROD_HOST..."
ssh "$PROD_USER@$PROD_HOST" "cd '$PROD_PATH' && git pull && nssm restart '$SERVICE_NAME'"
echo "Deploy complete."
