#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
REPO_DIR="$(cd "$APP_DIR/.." && pwd)"
CONFIG_PATH="$REPO_DIR/config.json"
EXAMPLE_CONFIG="$REPO_DIR/config.example.json"
ENV_PATH="$REPO_DIR/.env"
EXAMPLE_ENV="$REPO_DIR/.env.example"
LOG_DIR="$APP_DIR/logs"
PROGRESS_PID=""

render_progress() {
  local percent="$1"
  local label="$2"
  local width=40
  local filled=$(( percent * width / 100 ))
  local empty=$(( width - filled ))
  local bar=""
  local spaces=""
  if (( filled > 0 )); then
    bar="$(printf "%${filled}s" "" | tr " " "=")"
  fi
  if (( empty > 0 )); then
    spaces="$(printf "%${empty}s" "")"
  fi
  printf "[%s%s] %3d%%  %s\n" "$bar" "$spaces" "$percent" "$label" >&2
}

finish_progress_line() {
  true
}

start_progress() {
  local start="$1"
  local end="$2"
  local label="$3"
  local delay="${4:-0.35}"
  (
    local pct="$start"
    while true; do
      render_progress "$pct" "$label"
      if (( pct < end )); then
        pct=$(( pct + 1 ))
      fi
      sleep "$delay"
    done
  ) &
  PROGRESS_PID="$!"
}

stop_progress() {
  local percent="$1"
  local label="$2"
  if [[ -n "${PROGRESS_PID:-}" ]]; then
    kill "$PROGRESS_PID" >/dev/null 2>&1 || true
    wait "$PROGRESS_PID" 2>/dev/null || true
    PROGRESS_PID=""
  fi
  render_progress "$percent" "$label"
  finish_progress_line
}

run_with_progress() {
  local start="$1"
  local end="$2"
  local done_percent="$3"
  local label="$4"
  shift 4
  start_progress "$start" "$end" "$label"
  "$@"
  local exit_code="$?"
  if [[ "$exit_code" -eq 0 ]]; then
    stop_progress "$done_percent" "$label done"
  else
    stop_progress "$start" "$label failed"
  fi
  return "$exit_code"
}

render_progress 1 "Preparing launcher"
finish_progress_line
VENV_PYTHON="$(
  run_with_progress 2 20 25 "Checking Python dependencies" "$SCRIPT_DIR/bootstrap_env.sh" | tail -n 1
)"

if [[ ! -f "$CONFIG_PATH" && -f "$EXAMPLE_CONFIG" ]]; then
  render_progress 28 "Creating config.json"
  finish_progress_line
  cp "$EXAMPLE_CONFIG" "$CONFIG_PATH"
  echo "Created $CONFIG_PATH from the example config."
  echo "Update resume_path and salary preferences when you are ready."
fi

if [[ ! -f "$ENV_PATH" && -f "$EXAMPLE_ENV" ]]; then
  render_progress 30 "Creating .env"
  finish_progress_line
  cp "$EXAMPLE_ENV" "$ENV_PATH"
  echo "Created $ENV_PATH from the example env file."
fi

PORT="$(
  "$VENV_PYTHON" - "$CONFIG_PATH" <<'PY'
import json
import sys
from pathlib import Path

path = Path(sys.argv[1])
if not path.exists():
    print("8765")
    raise SystemExit
try:
    config = json.loads(path.read_text(encoding="utf-8"))
except json.JSONDecodeError as exc:
    print(
        f"Invalid JSON in {path}: line {exc.lineno}, column {exc.colno}. "
        "Please fix config.json. Common issue: missing quote or comma in resume_path.",
        file=sys.stderr,
    )
    raise SystemExit(2)
print(config.get("dashboard", {}).get("port", 8765))
PY
)"
SERVER_PORT="$PORT"

mkdir -p "$LOG_DIR"
cd "$REPO_DIR"

if [[ -f "$CONFIG_PATH" ]]; then
  JOB_AGENT_CONFIG="$CONFIG_PATH" JOB_AGENT_PYTHON="$VENV_PYTHON" run_with_progress 35 82 85 "Running fresh job search" "$VENV_PYTHON" "$APP_DIR/job_agent.py" --config "$CONFIG_PATH" || {
    echo "Fresh job search failed. Starting dashboard anyway so you can inspect the setup."
  }
fi

if /usr/sbin/lsof -nP -iTCP:"$PORT" -sTCP:LISTEN >/dev/null 2>&1; then
  SERVER_PORT="$(
    "$VENV_PYTHON" - <<'PY'
import socket
with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
    sock.bind(("127.0.0.1", 0))
    print(sock.getsockname()[1])
PY
  )"
  echo "Port $PORT is already in use. Starting this copy on http://127.0.0.1:$SERVER_PORT"
fi

if ! /usr/sbin/lsof -nP -iTCP:"$SERVER_PORT" -sTCP:LISTEN >/dev/null 2>&1; then
  render_progress 90 "Starting dashboard server"
  finish_progress_line
  nohup env JOB_AGENT_CONFIG="$CONFIG_PATH" JOB_AGENT_PYTHON="$VENV_PYTHON" "$VENV_PYTHON" "$APP_DIR/job_agent_server.py" --port "$SERVER_PORT" > "$LOG_DIR/dashboard.out.log" 2> "$LOG_DIR/dashboard.err.log" &
  sleep 1
else
  echo "Dashboard server is already running on http://127.0.0.1:$SERVER_PORT"
fi

render_progress 96 "Opening dashboard"
finish_progress_line
open "http://127.0.0.1:$SERVER_PORT"
render_progress 100 "Dashboard ready"
finish_progress_line
echo "Dashboard ready: http://127.0.0.1:$SERVER_PORT"
