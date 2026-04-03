#!/bin/bash
# ============================================================
# deploy.sh — Upload mne-bank-parser files to production server
# Usage: ./deploy.sh           — deploy everything
#        ./deploy.sh loader    — deploy only loader1c
#        ./deploy.sh izvod     — deploy only izvod (Python app)
# ============================================================
set -euo pipefail

SSH_KEY="${SSH_KEY:-$HOME/.ssh/id_tools_eddsa}"
SSH_HOST="${SSH_HOST:-sshuser@10.252.1.47}"
PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
TARGET="${1:-all}"

GREEN='\033[0;32m'
NC='\033[0m'
log() { echo -e "${GREEN}[DEPLOY]${NC} $*"; }

scp_cmd() { scp -i "$SSH_KEY" -o ConnectTimeout=10 -o ServerAliveInterval=15 -o ServerAliveCountMax=4 "$@"; }
ssh_cmd() { ssh -i "$SSH_KEY" -o ConnectTimeout=10 -o ServerAliveInterval=15 -o ServerAliveCountMax=4 -T "$SSH_HOST" "$@" </dev/null; }

# ============================================================
# Izvod — Python bank statement parser
# ============================================================
deploy_izvod() {
    REMOTE_DIR="C:/scripts/mne-bank-parser-main"

    log "Uploading app/*.py ..."
    for f in "$PROJECT_DIR"/app/*.py; do
        scp_cmd "$f" "$SSH_HOST:$REMOTE_DIR/app/"
    done

    log "Uploading app/parsers/*.py + *.json ..."
    for f in "$PROJECT_DIR"/app/parsers/*.py "$PROJECT_DIR"/app/parsers/*.json; do
        [ -f "$f" ] && scp_cmd "$f" "$SSH_HOST:$REMOTE_DIR/app/parsers/"
    done

    log "Uploading app/templates/ ..."
    scp_cmd -r "$PROJECT_DIR/app/templates/" "$SSH_HOST:$REMOTE_DIR/app/templates/"

    log "Uploading app/static/ ..."
    scp_cmd -r "$PROJECT_DIR/app/static/" "$SSH_HOST:$REMOTE_DIR/app/static/"

    log "Uploading config.yaml ..."
    scp_cmd "$PROJECT_DIR/config.yaml" "$SSH_HOST:$REMOTE_DIR/config.yaml"

    log "Uploading Dockerfile & requirements.txt ..."
    scp_cmd "$PROJECT_DIR/Dockerfile" "$SSH_HOST:$REMOTE_DIR/Dockerfile"
    scp_cmd "$PROJECT_DIR/requirements.txt" "$SSH_HOST:$REMOTE_DIR/requirements.txt"

    echo ""
    log "Files uploaded to $SSH_HOST:$REMOTE_DIR"
    log "To rebuild and restart on the server:"
    log "  cd C:\\scripts\\mne-bank-parser-main"
    log "  set DOCKER_BUILDKIT=0&& docker compose up --build -d"
}

# ============================================================
# Loader1C — .NET service for loading statements into 1C
# ============================================================
deploy_loader() {
    LOADER_DIR="C:/scripts/loader1c"

    log "Uploading loader1c sources ..."
    for f in "$PROJECT_DIR"/loader1c/*.cs "$PROJECT_DIR"/loader1c/appsettings.json; do
        [ -f "$f" ] || continue
        # Skip test files
        case "$(basename "$f")" in Test*|Check*|check*) continue ;; esac
        scp_cmd "$f" "$SSH_HOST:$LOADER_DIR/$(basename "$f")"
    done
    scp_cmd "$PROJECT_DIR/loader1c/Loader1C.csproj" "$SSH_HOST:$LOADER_DIR/Loader1C.csproj"
    scp_cmd "$PROJECT_DIR/loader1c/op_types.json" "$SSH_HOST:$LOADER_DIR/op_types.json"

    log "Stopping Loader1C service ..."
    ssh_cmd "sc stop Loader1C 2>nul & timeout /t 5 >nul & taskkill /F /IM Loader1C.exe 2>nul & timeout /t 2 >nul & ver >nul" || true

    log "Building loader1c ..."
    ssh_cmd "cmd /c \"cd /d $LOADER_DIR && dotnet publish Loader1C.csproj -c Release -o $LOADER_DIR --nologo -v quiet <nul\""

    log "Starting Loader1C service ..."
    ssh_cmd "sc start Loader1C 2>nul & ver >nul" || true

    log ""
    log "Loader1C deployed to $LOADER_DIR"
}

# ============================================================
case "$TARGET" in
    loader)  deploy_loader ;;
    izvod)   deploy_izvod ;;
    all)     deploy_izvod; echo ""; deploy_loader ;;
    *)       echo "Usage: $0 [all|izvod|loader]"; exit 1 ;;
esac
