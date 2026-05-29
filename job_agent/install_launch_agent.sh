#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
CONFIG_PATH="$REPO_DIR/config.json"
if [[ ! -f "$CONFIG_PATH" && -f "$SCRIPT_DIR/config.json" ]]; then
  CONFIG_PATH="$SCRIPT_DIR/config.json"
fi
PLIST_PATH="$HOME/Library/LaunchAgents/com.codex.job-match-agent.plist"
SERVER_PLIST_PATH="$HOME/Library/LaunchAgents/com.codex.job-match-dashboard.plist"
LOG_DIR="$SCRIPT_DIR/logs"

if [[ ! -f "$CONFIG_PATH" ]]; then
  echo "Missing $CONFIG_PATH"
  echo "Copy config.example.json to config.json and set resume_path first."
  exit 1
fi

mkdir -p "$HOME/Library/LaunchAgents" "$LOG_DIR"

cat > "$PLIST_PATH" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>com.codex.job-match-agent</string>
  <key>ProgramArguments</key>
  <array>
    <string>/usr/bin/python3</string>
    <string>$SCRIPT_DIR/job_agent_daemon.py</string>
    <string>--config</string>
    <string>$CONFIG_PATH</string>
  </array>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>StandardOutPath</key>
  <string>$LOG_DIR/job-agent.out.log</string>
  <key>StandardErrorPath</key>
  <string>$LOG_DIR/job-agent.err.log</string>
</dict>
</plist>
PLIST

cat > "$SERVER_PLIST_PATH" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>com.codex.job-match-dashboard</string>
  <key>ProgramArguments</key>
  <array>
    <string>/usr/bin/python3</string>
    <string>$SCRIPT_DIR/job_agent_server.py</string>
  </array>
  <key>EnvironmentVariables</key>
  <dict>
    <key>JOB_AGENT_CONFIG</key>
    <string>$CONFIG_PATH</string>
  </dict>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>StandardOutPath</key>
  <string>$LOG_DIR/job-dashboard.out.log</string>
  <key>StandardErrorPath</key>
  <string>$LOG_DIR/job-dashboard.err.log</string>
</dict>
</plist>
PLIST

launchctl unload "$PLIST_PATH" 2>/dev/null || true
launchctl load "$PLIST_PATH"
launchctl start com.codex.job-match-agent
launchctl unload "$SERVER_PLIST_PATH" 2>/dev/null || true
launchctl load "$SERVER_PLIST_PATH"
launchctl start com.codex.job-match-dashboard

echo "Installed launch agent: $PLIST_PATH"
echo "Installed dashboard server: $SERVER_PLIST_PATH"
echo "The job search runs at login, after wake, and every 6 hours. The dashboard runs at http://127.0.0.1:8765."
