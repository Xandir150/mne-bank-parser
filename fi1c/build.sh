#!/bin/bash
# ============================================================
# build.sh — сборка FinIskaziCG.epf на прод-сервере через 1C Designer
# Usage: ./build.sh          — залить исходники, собрать, показать лог
# ============================================================
set -euo pipefail

SSH_KEY="${SSH_KEY:-$HOME/.ssh/id_tools_eddsa}"
SSH_HOST="${SSH_HOST:-sshuser@10.252.1.47}"
DIR="$(cd "$(dirname "$0")" && pwd)"
REMOTE=C:/scripts/fi-epf
V8='C:\Program Files\1cv8\8.3.27.1859\bin\1cv8.exe'

ssh_cmd() { ssh -i "$SSH_KEY" -o ConnectTimeout=10 -T "$SSH_HOST" "$@" </dev/null; }
scp_cmd() { scp -i "$SSH_KEY" -o ConnectTimeout=10 "$@"; }

echo "[build] ensuring UTF-8 BOM on sources ..."
find "$DIR/src" -type f \( -name '*.xml' -o -name '*.bsl' -o -name '*.txt' \) | while read -r f; do
    if [ "$(head -c 3 "$f" | xxd -p)" != "efbbbf" ]; then
        printf '\xEF\xBB\xBF' | cat - "$f" > "$f.tmp" && mv "$f.tmp" "$f"
        echo "  +BOM $(basename "$f")"
    fi
done

echo "[build] uploading sources ..."
ssh_cmd "cmd /c \"rmdir /s /q $(echo $REMOTE | tr / '\\')\\src\\FinIskaziCG 2>nul & mkdir $(echo $REMOTE | tr / '\\')\\src 2>nul & ver >nul\"" || true
scp_cmd -r "$DIR/src/FinIskaziCG.xml" "$SSH_HOST:$REMOTE/src/"
scp_cmd -r "$DIR/src/FinIskaziCG" "$SSH_HOST:$REMOTE/src/"

echo "[build] running designer ..."
# Сборка против реальной базы (нужен контекст конфигурации для типов вроде CatalogRef.Организации).
# /LoadExternalDataProcessorOrReportFromFiles ничего не меняет в базе — только конвертирует XML в epf.
BUILD_BASE="${BUILD_BASE:-base_m}"
# .bat с точной командной строкой (квотинг /N"..." — канонический для 1С CLI)
cat > /tmp/fi_build.bat <<BATEOF
@echo off
del C:\\scripts\\fi-epf\\build\\FinIskaziCG.epf 2>nul
start "" /wait "$V8" DESIGNER /S navus-server\\$BUILD_BASE /N"Zaykov Andrey" /P"19700214" /DisableStartupDialogs /LoadExternalDataProcessorOrReportFromFiles C:\\scripts\\fi-epf\\src\\FinIskaziCG.xml C:\\scripts\\fi-epf\\build\\FinIskaziCG.epf /Out C:\\scripts\\fi-epf\\load.log
echo exit: %ERRORLEVEL%
BATEOF
scp_cmd /tmp/fi_build.bat "$SSH_HOST:C:/scripts/fi-epf/fi_build.bat"
ssh_cmd "cmd /c C:\\scripts\\fi-epf\\fi_build.bat"

ssh_cmd "powershell -NoProfile -Command \"[Console]::OutputEncoding=[Text.Encoding]::UTF8; Get-Content 'C:\\scripts\\fi-epf\\load.log' -Encoding UTF8; Write-Output '---'; if (Test-Path 'C:\\scripts\\fi-epf\\build\\FinIskaziCG.epf') { (Get-Item 'C:\\scripts\\fi-epf\\build\\FinIskaziCG.epf').Length } else { Write-Output 'EPF MISSING' }\""
