#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
REPO_DIR="$(cd "$APP_DIR/.." && pwd)"
CONFIG_PATH="$REPO_DIR/config.json"
FALLBACK_CONFIG_PATH="$APP_DIR/config.json"
VENV_PYTHON="$("$SCRIPT_DIR/bootstrap_env.sh" | tail -n 1)"

if [[ ! -f "$CONFIG_PATH" && -f "$FALLBACK_CONFIG_PATH" ]]; then
  CONFIG_PATH="$FALLBACK_CONFIG_PATH"
fi

if [[ ! -f "$CONFIG_PATH" ]]; then
  echo "Missing $REPO_DIR/config.json"
  echo "Copy config.example.json to config.json and set resume_path first."
  exit 1
fi

cd "$REPO_DIR"
exec env JOB_AGENT_CONFIG="$CONFIG_PATH" JOB_AGENT_PYTHON="$VENV_PYTHON" "$VENV_PYTHON" "$APP_DIR/job_agent.py" --config "$CONFIG_PATH"
