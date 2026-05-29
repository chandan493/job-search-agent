#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
VENV_DIR="$APP_DIR/.venv"
PYTHON_BIN="${JOB_AGENT_BOOTSTRAP_PYTHON:-}"
REQUIREMENTS_PATH="$APP_DIR/requirements.txt"
STAMP_PATH="$VENV_DIR/.requirements.sha256"

if [[ -z "$PYTHON_BIN" ]]; then
  if command -v python3 >/dev/null 2>&1; then
    PYTHON_BIN="$(command -v python3)"
  elif [[ -x /usr/bin/python3 ]]; then
    PYTHON_BIN="/usr/bin/python3"
  else
    echo "python3 is required but was not found."
    echo "Install Python 3 from https://www.python.org/downloads/ and run this launcher again."
    exit 1
  fi
fi

if [[ ! -x "$VENV_DIR/bin/python" ]]; then
  echo "Creating local Python environment..."
  "$PYTHON_BIN" -m venv "$VENV_DIR"
fi

VENV_PYTHON="$VENV_DIR/bin/python"
CURRENT_HASH="$("$VENV_PYTHON" - "$REQUIREMENTS_PATH" <<'PY'
import hashlib
import sys
from pathlib import Path

path = Path(sys.argv[1])
print(hashlib.sha256(path.read_bytes()).hexdigest() if path.exists() else "")
PY
)"
INSTALLED_HASH="$(cat "$STAMP_PATH" 2>/dev/null || true)"

if [[ "$CURRENT_HASH" != "$INSTALLED_HASH" ]]; then
  echo "Installing Python dependencies..."
  "$VENV_PYTHON" -m pip install --upgrade pip
  "$VENV_PYTHON" -m pip install -r "$REQUIREMENTS_PATH"
  echo "$CURRENT_HASH" > "$STAMP_PATH"
else
  if ! "$VENV_PYTHON" - <<'PY' >/dev/null 2>&1
import docx
PY
  then
    echo "Repairing Python dependencies..."
    "$VENV_PYTHON" -m pip install -r "$REQUIREMENTS_PATH"
    echo "$CURRENT_HASH" > "$STAMP_PATH"
  fi
fi

echo "$VENV_PYTHON"
