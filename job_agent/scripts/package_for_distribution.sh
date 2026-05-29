#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
REPO_DIR="$(cd "$APP_DIR/.." && pwd)"
DIST_DIR="$REPO_DIR/dist"
PACKAGE_DIR="$DIST_DIR/job-agent-util-share"
PACKAGE_PATH="$DIST_DIR/job-agent-util.zip"

mkdir -p "$DIST_DIR"
rm -rf "$PACKAGE_DIR" "$PACKAGE_PATH"
mkdir -p "$PACKAGE_DIR"

cd "$REPO_DIR"
cp -R job_agent "$PACKAGE_DIR/"
cp .gitignore config.example.json .env.example "$PACKAGE_DIR/"
rm -f "$PACKAGE_DIR/job_agent/config.json" "$PACKAGE_DIR/job_agent/.env"
rm -rf "$PACKAGE_DIR/job_agent/data" "$PACKAGE_DIR/job_agent/logs" "$PACKAGE_DIR/job_agent/.venv" "$PACKAGE_DIR/job_agent/__pycache__"
find "$PACKAGE_DIR" -name ".DS_Store" -delete
find "$PACKAGE_DIR" -name "__pycache__" -type d -prune -exec rm -rf {} +
find "$PACKAGE_DIR" -name "*.pyc" -delete

cd "$DIST_DIR"
zip -r "$PACKAGE_PATH" job-agent-util-share -x "*.DS_Store" -x "*.pyc" -x "*/__pycache__/*"

echo "$PACKAGE_DIR"
echo "$PACKAGE_PATH"
