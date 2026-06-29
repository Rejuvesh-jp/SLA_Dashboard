#!/bin/bash
# ─────────────────────────────────────────────────────────────
# SLA Dashboard — Linux startup script
# Usage:  bash start_app.sh
# ─────────────────────────────────────────────────────────────

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$SCRIPT_DIR/venv"
APP_HOST="0.0.0.0"
APP_PORT="8001"

cd "$SCRIPT_DIR"

# ── 1. Create virtual environment if missing ──────────────────
if [ ! -d "$VENV_DIR" ]; then
    echo "[setup] Creating Python virtual environment..."
    python3 -m venv "$VENV_DIR"
fi

# ── 2. Activate venv ─────────────────────────────────────────
source "$VENV_DIR/bin/activate"

# ── 3. Install / upgrade dependencies ────────────────────────
echo "[setup] Installing dependencies..."
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt

# ── 4. Start the app ─────────────────────────────────────────
echo "[start] Starting SLA Dashboard on http://$APP_HOST:$APP_PORT"
exec python main.py
