#!/bin/bash
# ============================================================
# deploy.sh — Upload mne-bank-parser files to production server
# Rebuilds and restarts are done manually on the server.
# Usage: ./deploy.sh
# ============================================================
set -euo pipefail

SSH_KEY="${SSH_KEY:-$HOME/.ssh/id_tools_eddsa}"
SSH_HOST="${SSH_HOST:-sshuser@10.252.1.47}"
REMOTE_DIR="C:/scripts/mne-bank-parser-main"
PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"

GREEN='\033[0;32m'
NC='\033[0m'
log() { echo -e "${GREEN}[DEPLOY]${NC} $*"; }

scp_cmd() { scp -i "$SSH_KEY" -o ConnectTimeout=10 "$@"; }

# ---- Upload app ----
log "Uploading app/*.py ..."
for f in "$PROJECT_DIR"/app/*.py; do
    scp_cmd "$f" "$SSH_HOST:$REMOTE_DIR/app/"
done

# ---- Upload parsers ----
log "Uploading app/parsers/*.py ..."
for f in "$PROJECT_DIR"/app/parsers/*.py; do
    scp_cmd "$f" "$SSH_HOST:$REMOTE_DIR/app/parsers/"
done

# ---- Upload templates ----
log "Uploading app/templates/ ..."
scp_cmd -r "$PROJECT_DIR/app/templates/" "$SSH_HOST:$REMOTE_DIR/app/templates/"

# ---- Upload static ----
log "Uploading app/static/ ..."
scp_cmd -r "$PROJECT_DIR/app/static/" "$SSH_HOST:$REMOTE_DIR/app/static/"

# ---- Upload config ----
log "Uploading config.yaml ..."
scp_cmd "$PROJECT_DIR/config.yaml" "$SSH_HOST:$REMOTE_DIR/config.yaml"

# ---- Upload Dockerfile & requirements (docker-compose.yml is managed on server) ----
log "Uploading Dockerfile & requirements.txt ..."
scp_cmd "$PROJECT_DIR/Dockerfile" "$SSH_HOST:$REMOTE_DIR/Dockerfile"
scp_cmd "$PROJECT_DIR/requirements.txt" "$SSH_HOST:$REMOTE_DIR/requirements.txt"

echo ""
log "Files uploaded to $SSH_HOST:$REMOTE_DIR"
log ""
log "To rebuild and restart on the server:"
log "  cd C:\\scripts\\mne-bank-parser-main"
log "  set DOCKER_BUILDKIT=0&& docker compose up --build -d"

# ============================================================
# Loader1C — .NET service for loading statements into 1C
# ============================================================
LOADER_REMOTE="C:/scripts/loader1c"
LOADER_SRC="C:/NET/Epsilon/izvod/loader1c"

log ""
log "Uploading loader1c sources ..."
ssh_cmd() { ssh -i "$SSH_KEY" -o ConnectTimeout=10 "$SSH_HOST" "$@"; }

# Upload all C# source files and project to build directory
for f in "$PROJECT_DIR"/loader1c/*.cs "$PROJECT_DIR"/loader1c/*.csproj "$PROJECT_DIR"/loader1c/appsettings.json; do
    [ -f "$f" ] && scp_cmd "$f" "$SSH_HOST:$LOADER_SRC/$(basename "$f")"
done

log "Building loader1c ..."
ssh_cmd "sc stop Loader1C 2>nul; timeout /t 2 >nul"
ssh_cmd "cd '$LOADER_SRC' && dotnet publish Loader1C.csproj -c Release -o '$LOADER_REMOTE' --nologo -v quiet"

log "Starting Loader1C service ..."
ssh_cmd "sc start Loader1C 2>nul || echo Service not yet registered"

log ""
log "Loader1C deployed to $LOADER_REMOTE"
log "If service is not registered yet, run on server:"
log "  sc create Loader1C binPath= \"C:\\scripts\\loader1c\\Loader1C.exe\" start= auto obj= \".\\USR1CV8\" password= \"<password>\""
