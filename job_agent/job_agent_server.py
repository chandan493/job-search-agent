#!/usr/bin/env python3
"""Local dashboard server for latest jobs and on-demand ATS resume downloads."""

from __future__ import annotations

import argparse
import cgi
import datetime as dt
import io
import json
import mimetypes
import os
import errno
import re
import sqlite3
import subprocess
import sys
import threading
import urllib.parse
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
REPO_DIR = BASE_DIR.parent
CONFIG_PATH = Path(
    os.environ.get(
        "JOB_AGENT_CONFIG",
        str(REPO_DIR / "config.json" if (REPO_DIR / "config.json").exists() else BASE_DIR / "config.json"),
    )
).expanduser()
DATA_DIR = CONFIG_PATH.parent / "data"
LOCAL_VENV_PYTHON = BASE_DIR / ".venv" / "bin" / "python"
BUNDLED_PYTHON = Path(os.environ["JOB_AGENT_BUNDLED_PYTHON"]).expanduser() if os.environ.get("JOB_AGENT_BUNDLED_PYTHON") else None
RUN_LOCK = threading.Lock()
SOURCE_LABELS = {
    "jobicy": "Jobicy",
    "remotejobsorg": "RemoteJobs.org",
    "remoteok": "RemoteOK",
    "remotive": "Remotive",
    "arbeitnow": "Arbeitnow",
    "careernest": "Career Nest",
    "workanywhere": "WorkAnywhere.pro",
    "adzuna": "Adzuna India",
    "jobsora": "Jobsora India",
    "shine": "Shine",
    "indeed_india": "Indeed India",
    "naukri": "Naukri",
    "instahyre": "Instahyre",
    "foundit": "foundit",
    "timesjobs": "TimesJobs",
    "cutshort": "Cutshort",
    "hirist": "Hirist",
    "apna": "Apna",
}


def load_config() -> dict:
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"Invalid JSON in {CONFIG_PATH}: line {exc.lineno}, column {exc.colno}. "
            "Check for a missing quote or comma, especially around resume_path."
        ) from exc


def write_config(config: dict) -> None:
    CONFIG_PATH.write_text(json.dumps(config, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def resolve_config_path(path_value: str) -> Path:
    path = Path(os.path.expanduser(path_value))
    if path.is_absolute():
        return path
    return CONFIG_PATH.parent / path


def settings_db_path(config: dict) -> Path:
    return resolve_config_path(config.get("database_path", "data/seen_jobs.sqlite3"))


def source_options(config: dict) -> list[dict]:
    sources = config.get("sources", {})
    keys = list(SOURCE_LABELS)
    for key in sources:
        if key not in keys:
            keys.append(key)
    return [
        {"key": key, "label": SOURCE_LABELS.get(key, key.replace("_", " ").title()), "enabled": bool(sources.get(key, False))}
        for key in keys
    ]


def settings_from_config(config: dict) -> dict:
    expected = config.get("expected_salary", {})
    search = config.get("search", {})
    return {
        "expected_salary": {
            "amount": expected.get("amount", 0),
            "currency": expected.get("currency", "INR"),
            "period": expected.get("period", "year"),
        },
        "roles": search.get("roles", []),
        "locations": search.get("locations", []),
        "required_keywords": search.get("required_keywords", []),
        "excluded_keywords": search.get("excluded_keywords", []),
        "min_score": search.get("min_score", 2),
        "freshness": search.get("freshness", "today"),
        "sources": source_options(config),
    }


def split_list(value: object) -> list[str]:
    if isinstance(value, list):
        raw = value
    else:
        raw = re.split(r"[\n,]", str(value or ""))
    return [re.sub(r"\s+", " ", str(item)).strip() for item in raw if re.sub(r"\s+", " ", str(item)).strip()]


def apply_settings_to_config(config: dict, settings: dict) -> dict:
    updated = json.loads(json.dumps(config))
    salary = settings.get("expected_salary", {})
    updated.setdefault("expected_salary", {})
    if "amount" in salary:
        updated["expected_salary"]["amount"] = int(float(salary.get("amount") or 0))
    if salary.get("currency"):
        updated["expected_salary"]["currency"] = str(salary.get("currency"))
    if salary.get("period"):
        updated["expected_salary"]["period"] = str(salary.get("period"))

    search = updated.setdefault("search", {})
    for key in ("roles", "locations", "required_keywords", "excluded_keywords"):
        if key in settings:
            search[key] = split_list(settings.get(key))
    if "min_score" in settings:
        search["min_score"] = int(float(settings.get("min_score") or 0))
    if settings.get("freshness") in {"today", "all"}:
        search["freshness"] = settings["freshness"]

    if "sources" in settings:
        sources = updated.setdefault("sources", {})
        incoming = settings["sources"]
        if isinstance(incoming, dict):
            for key, enabled in incoming.items():
                if key in SOURCE_LABELS or key in sources:
                    sources[key] = bool(enabled)
        elif isinstance(incoming, list):
            enabled_keys = {str(item) for item in incoming}
            for key in list(SOURCE_LABELS) + [key for key in sources if key not in SOURCE_LABELS]:
                sources[key] = key in enabled_keys
    return updated


def save_settings_to_sqlite(config: dict, settings: dict) -> None:
    db_path = settings_db_path(config)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS app_settings (
              key TEXT PRIMARY KEY,
              value TEXT NOT NULL,
              updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            "INSERT OR REPLACE INTO app_settings (key, value, updated_at) VALUES (?, ?, ?)",
            (
                "job_search_preferences",
                json.dumps(settings, indent=2, ensure_ascii=False),
                dt.datetime.now(dt.timezone.utc).isoformat(),
            ),
        )
        conn.commit()


def python_executable() -> str:
    configured = os.environ.get("JOB_AGENT_PYTHON")
    if configured:
        return configured
    if LOCAL_VENV_PYTHON.exists():
        return str(LOCAL_VENV_PYTHON)
    if BUNDLED_PYTHON and BUNDLED_PYTHON.exists():
        return str(BUNDLED_PYTHON)
    return sys.executable


def safe_download_name(value: str) -> str:
    safe = "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in value).strip("_")
    return (safe[:120] or "ats_resume") + ".docx"


def safe_upload_name(value: str) -> str:
    base = Path(value or "resume").name
    stem = "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in Path(base).stem).strip("_")
    suffix = Path(base).suffix.lower()
    if suffix not in {".pdf", ".docx", ".txt"}:
        raise ValueError("Upload a PDF, DOCX, or TXT resume.")
    return f"{stem[:70] or 'resume'}_{uuid.uuid4().hex[:8]}{suffix}"


def config_relative_path(path: Path) -> str:
    try:
        return str(path.relative_to(CONFIG_PATH.parent))
    except ValueError:
        return str(path)


def configured_resume_path(config: dict) -> Path | None:
    value = str(config.get("resume_path") or "").strip()
    if not value or value.startswith("/absolute/path/to/"):
        return None
    return resolve_config_path(value)


def has_configured_resume(config: dict) -> bool:
    path = configured_resume_path(config)
    return bool(path and path.exists() and path.is_file())


class ReusableThreadingHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = True


class DashboardHandler(BaseHTTPRequestHandler):
    server_version = "JobAgentDashboard/1.0"

    def log_message(self, fmt: str, *args) -> None:
        print(f"{self.address_string()} - {fmt % args}", flush=True)

    def send_text(self, status: int, text: str) -> None:
        payload = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def send_json(self, status: int, payload: dict) -> None:
        body = json.dumps(payload, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_html(self, status: int, html_text: str) -> None:
        payload = html_text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def serve_file(self, path: Path, content_type: str | None = None, download_name: str | None = None) -> None:
        if not path.exists():
            self.send_text(404, f"Not found: {path}")
            return
        payload = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type or mimetypes.guess_type(path.name)[0] or "application/octet-stream")
        self.send_header("Content-Length", str(len(payload)))
        if download_name:
            self.send_header("Content-Disposition", f'attachment; filename="{download_name}"')
        self.end_headers()
        self.wfile.write(payload)

    def serve_dashboard(self) -> None:
        try:
            config = load_config()
        except Exception as exc:
            self.send_text(500, str(exc))
            return
        path = DATA_DIR / "latest_jobs.html"
        if not has_configured_resume(config) or not path.exists():
            self.serve_resume_wizard(config)
            return
        text = path.read_text(encoding="utf-8", errors="replace")
        text = rebase_dashboard_resume_links(text)
        text = inject_dashboard_run_control(text)
        self.send_html(200, text)

    def serve_resume_wizard(self, config: dict) -> None:
        current_resume = configured_resume_path(config)
        current_label = current_resume.name if current_resume and current_resume.exists() else "No resume uploaded yet"
        html_text = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>RoleForge - Upload Resume</title>
  <style>
    :root {{
      --ink: #ffffff; --muted: #b8b8b8; --paper: #050505; --panel: #101010;
      --line: #2c2c2c; --accent: #ffd21f; --accent-strong: #f2b705;
      --shadow: 0 18px 50px rgba(0,0,0,.45);
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }}
    * {{ box-sizing: border-box; }}
    body {{ min-height: 100vh; margin: 0; color: var(--ink); background: var(--paper); display: grid; place-items: center; padding: 20px; }}
    .wizard {{ width: min(620px, 100%); display: grid; gap: 18px; }}
    .brand {{ display: inline-flex; align-items: center; gap: 12px; font-weight: 950; }}
    .brand-name {{ color: var(--ink); font-weight: 950; white-space: nowrap; }}
    .brand-mark {{ position: relative; width: 42px; height: 42px; flex: 0 0 42px; display: grid; place-items: center; overflow: hidden; border-radius: 10px; background: linear-gradient(135deg, var(--accent), var(--accent-strong)); color: #050505; font-size: 13px; font-weight: 950; box-shadow: 0 12px 28px rgba(255,210,31,.16); }}
    .brand-mark::before {{ content: ""; position: absolute; inset: 7px; border: 2px solid rgba(5,5,5,.24); border-radius: 7px; transform: rotate(-8deg); }}
    .brand-mark::after {{ content: ""; position: absolute; right: 7px; top: 7px; width: 7px; height: 7px; border-radius: 50%; background: rgba(5,5,5,.66); box-shadow: -14px 16px 0 rgba(5,5,5,.36); }}
    .panel {{ border: 1px solid rgba(255,210,31,.24); border-radius: 8px; background: var(--panel); box-shadow: var(--shadow); overflow: hidden; }}
    .panel-head {{ padding: 18px; border-bottom: 1px solid var(--line); }}
    h1 {{ margin: 0; color: var(--accent); font-size: clamp(2rem, 7vw, 4.2rem); line-height: .95; letter-spacing: 0; }}
    .copy {{ margin: 12px 0 0; color: rgba(255,255,255,.78); font-size: 1rem; }}
    .body {{ display: grid; gap: 14px; padding: 18px; }}
    .dropzone {{ min-height: 190px; display: grid; place-items: center; gap: 8px; padding: 22px; border: 1px dashed rgba(255,210,31,.5); border-radius: 8px; background: rgba(255,255,255,.04); cursor: pointer; text-align: center; }}
    .dropzone.is-dragging {{ border-color: var(--accent); background: rgba(255,210,31,.09); }}
    .dropzone strong {{ display: block; color: var(--accent); font-size: 17px; }}
    .dropzone span {{ display: block; color: var(--muted); font-size: 12px; }}
    .meta {{ color: var(--muted); font-size: 12px; }}
    .actions {{ display: flex; justify-content: flex-end; gap: 10px; padding: 14px 18px; border-top: 1px solid var(--line); }}
    .button {{ min-height: 40px; border: 1px solid transparent; border-radius: 8px; padding: 10px 13px; font: inherit; font-weight: 950; cursor: pointer; }}
    .button.primary {{ color: #050505; background: var(--accent); }}
    .button.primary:hover {{ background: var(--accent-strong); }}
    .button.secondary {{ color: var(--ink); border-color: rgba(255,210,31,.42); background: rgba(255,255,255,.06); }}
    .button:disabled {{ cursor: wait; opacity: .7; }}
    .loader {{ --progress-angle: 3.6deg; display: none; align-items: center; gap: 10px; color: var(--accent); font-weight: 950; }}
    .loader.is-active {{ display: inline-flex; }}
    .ring {{ position: relative; width: 42px; height: 42px; border-radius: 50%; display: grid; place-items: center; background: conic-gradient(var(--accent) var(--progress-angle), rgba(255,255,255,.14) 0); }}
    .ring::after {{ content: ""; position: absolute; width: 30px; height: 30px; border-radius: 50%; background: #050505; }}
    .percent {{ position: relative; z-index: 1; color: var(--ink); font-size: 10px; }}
  </style>
</head>
<body>
  <main class="wizard">
    <div class="brand"><span class="brand-mark">RF</span><span class="brand-name">RoleForge</span></div>
    <section class="panel">
      <div class="panel-head">
        <h1>Upload your resume</h1>
        <p class="copy">Start by adding a resume. RoleForge will parse it, find matching jobs, and build your dashboard.</p>
      </div>
      <form id="resumeUploadForm">
        <div class="body">
          <label class="dropzone" id="resumeDropzone" for="resumeUploadInput">
            <input id="resumeUploadInput" name="resume" type="file" accept=".pdf,.docx,.txt,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,text/plain" hidden>
            <span><strong>Drop your resume here</strong><span>PDF, DOCX, or TXT</span></span>
          </label>
          <div class="meta" id="resumeFileName">{html_escape(current_label)}</div>
          <div class="loader" id="wizardLoader" aria-live="polite">
            <span class="ring" id="wizardRing"><span class="percent" id="wizardPercent">1%</span></span>
            <span id="wizardMessage">Waiting for resume</span>
          </div>
          <div class="meta" id="wizardStatus"></div>
        </div>
        <div class="actions">
          <button class="button primary" id="resumeUploadSubmitButton" type="submit">Upload and run</button>
        </div>
      </form>
    </section>
  </main>
  <script>
    const form = document.getElementById('resumeUploadForm');
    const input = document.getElementById('resumeUploadInput');
    const dropzone = document.getElementById('resumeDropzone');
    const fileName = document.getElementById('resumeFileName');
    const statusEl = document.getElementById('wizardStatus');
    const loader = document.getElementById('wizardLoader');
    const ring = document.getElementById('wizardRing');
    const percent = document.getElementById('wizardPercent');
    const message = document.getElementById('wizardMessage');
    const submit = document.getElementById('resumeUploadSubmitButton');
    let timer = null;
    let progress = 1;
    function selectedFile() {{ return input.files && input.files[0]; }}
    function updateFileName() {{ const file = selectedFile(); fileName.textContent = file ? file.name : '{html_escape(current_label)}'; }}
    function setProgress(value, label) {{
      progress = Math.max(1, Math.min(100, Math.round(value)));
      loader.classList.add('is-active');
      ring.style.setProperty('--progress-angle', `${{progress * 3.6}}deg`);
      percent.textContent = `${{progress}}%`;
      if (label) message.textContent = label;
    }}
    function startProgress(label) {{
      stopProgress(false);
      setProgress(1, label || 'Uploading resume');
      timer = window.setInterval(() => {{
        const step = progress < 70 ? 3 : progress < 94 ? 2 : 1;
        if (progress < 98) setProgress(progress + step);
      }}, 650);
    }}
    function stopProgress(hide = true) {{
      if (timer) window.clearInterval(timer);
      timer = null;
      if (hide) loader.classList.remove('is-active');
    }}
    async function uploadAndRun(event) {{
      event.preventDefault();
      const file = selectedFile();
      if (!file) {{ statusEl.textContent = 'Choose a resume file first.'; return; }}
      const name = file.name.toLowerCase();
      if (!['.pdf', '.docx', '.txt'].some((suffix) => name.endsWith(suffix))) {{
        statusEl.textContent = 'Upload a PDF, DOCX, or TXT resume.';
        return;
      }}
      submit.disabled = true;
      statusEl.textContent = '';
      startProgress('Uploading resume');
      try {{
        const formData = new FormData();
        formData.append('resume', file);
        const uploadResponse = await fetch('/resume-upload', {{ method: 'POST', body: formData }});
        const uploadPayload = await uploadResponse.json();
        if (!uploadResponse.ok || !uploadPayload.ok) throw new Error(uploadPayload.message || 'Resume upload failed');
        setProgress(28, 'Parsing resume');
        const runResponse = await fetch('/run-agent', {{ method: 'POST' }});
        const runPayload = await runResponse.json();
        if (!runResponse.ok || !runPayload.ok) throw new Error(runPayload.message || runPayload.stderr || 'Fresh report failed');
        stopProgress(false);
        setProgress(100, 'Opening dashboard');
        window.setTimeout(() => window.location.reload(), 500);
      }} catch (error) {{
        stopProgress();
        statusEl.textContent = error.message || 'Resume upload failed';
        submit.disabled = false;
      }}
    }}
    input.addEventListener('change', updateFileName);
    form.addEventListener('submit', uploadAndRun);
    ['dragenter', 'dragover'].forEach((name) => dropzone.addEventListener(name, (event) => {{
      event.preventDefault();
      dropzone.classList.add('is-dragging');
    }}));
    ['dragleave', 'drop'].forEach((name) => dropzone.addEventListener(name, (event) => {{
      event.preventDefault();
      dropzone.classList.remove('is-dragging');
    }}));
    dropzone.addEventListener('drop', (event) => {{
      const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
      if (!file) return;
      const transfer = new DataTransfer();
      transfer.items.add(file);
      input.files = transfer.files;
      updateFileName();
    }});
  </script>
</body>
</html>"""
        self.send_html(200, html_text)

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        if path in {"/", "/latest_jobs.html"}:
            self.serve_dashboard()
            return
        if path == "/latest_jobs.json":
            self.serve_file(DATA_DIR / "latest_jobs.json", "application/json; charset=utf-8")
            return
        if path == "/settings":
            self.get_settings()
            return
        if path.startswith("/resume/"):
            job_id = urllib.parse.unquote(path.rsplit("/", 1)[-1])
            self.generate_resume(job_id)
            return
        self.send_text(404, "Unknown dashboard route")

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path in {"/run-agent", "/run"}:
            self.run_agent()
            return
        if parsed.path == "/settings":
            self.save_settings()
            return
        if parsed.path == "/resume-upload":
            self.upload_resume()
            return
        self.send_text(404, "Unknown dashboard route")

    def read_json_body(self) -> dict:
        length = int(self.headers.get("Content-Length", "0") or "0")
        if length <= 0:
            return {}
        body = self.rfile.read(length).decode("utf-8", errors="replace")
        return json.loads(body)

    def get_settings(self) -> None:
        config = load_config()
        self.send_json(200, {"ok": True, "settings": settings_from_config(config)})

    def save_settings(self) -> None:
        try:
            incoming = self.read_json_body()
            config = load_config()
            updated = apply_settings_to_config(config, incoming)
            write_config(updated)
            settings = settings_from_config(updated)
            save_settings_to_sqlite(updated, settings)
        except Exception as exc:
            self.send_json(400, {"ok": False, "message": f"Could not save settings: {exc}"})
            return
        self.send_json(200, {"ok": True, "settings": settings})

    def upload_resume(self) -> None:
        try:
            content_type = self.headers.get("Content-Type", "")
            length = int(self.headers.get("Content-Length", "0") or "0")
            if length <= 0:
                raise ValueError("No resume file was uploaded.")
            if length > 25 * 1024 * 1024:
                raise ValueError("Resume upload is too large. Please keep it under 25 MB.")
            body = self.rfile.read(length)
            form = cgi.FieldStorage(
                fp=io.BytesIO(body),
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": content_type,
                    "CONTENT_LENGTH": str(length),
                },
                keep_blank_values=True,
            )
            field = form["resume"] if "resume" in form else None
            if field is None or not getattr(field, "filename", ""):
                raise ValueError("Choose a resume file before uploading.")
            filename = safe_upload_name(field.filename)
            payload = field.file.read()
            if not payload:
                raise ValueError("Uploaded resume file is empty.")
            upload_dir = DATA_DIR / "uploaded_resumes"
            upload_dir.mkdir(parents=True, exist_ok=True)
            resume_path = upload_dir / filename
            resume_path.write_bytes(payload)
            config = load_config()
            config["resume_path"] = config_relative_path(resume_path)
            write_config(config)
        except Exception as exc:
            self.send_json(400, {"ok": False, "message": f"Could not upload resume: {exc}"})
            return
        self.send_json(
            200,
            {
                "ok": True,
                "resume_path": config_relative_path(resume_path),
                "filename": resume_path.name,
                "message": "Resume updated. Running a fresh report now.",
            },
        )

    def run_agent(self) -> None:
        if not RUN_LOCK.acquire(blocking=False):
            self.send_json(409, {"ok": False, "message": "A job search run is already in progress."})
            return
        cmd = [
            python_executable(),
            str(BASE_DIR / "job_agent.py"),
            "--config",
            str(CONFIG_PATH),
        ]
        env = os.environ.copy()
        env["JOB_AGENT_CONFIG"] = str(CONFIG_PATH)
        try:
            result = subprocess.run(
                cmd,
                cwd=str(BASE_DIR.parent),
                check=False,
                capture_output=True,
                text=True,
                timeout=900,
                env=env,
            )
        except subprocess.TimeoutExpired:
            self.send_json(504, {"ok": False, "message": "Job search timed out after 15 minutes."})
            return
        except Exception as exc:
            self.send_json(500, {"ok": False, "message": f"Job search failed: {exc}"})
            return
        finally:
            RUN_LOCK.release()

        stdout = result.stdout.strip()
        parsed_stdout = None
        if stdout:
            try:
                parsed_stdout = json.loads(stdout)
            except json.JSONDecodeError:
                parsed_stdout = stdout
        payload = {
            "ok": result.returncode == 0,
            "return_code": result.returncode,
            "result": parsed_stdout,
            "stderr": result.stderr.strip(),
        }
        self.send_json(200 if result.returncode == 0 else 500, payload)

    def generate_resume(self, job_id: str) -> None:
        latest_path = DATA_DIR / "latest_jobs.json"
        if not latest_path.exists():
            self.send_text(404, "No latest job run found yet. Run job_agent.py first.")
            return
        payload = json.loads(latest_path.read_text(encoding="utf-8"))
        job = next((item for item in payload.get("jobs", []) if item.get("job_uid") == job_id), None)
        if not job:
            self.send_text(404, f"Job not found in latest run: {job_id}")
            return
        cmd = [
            python_executable(),
            str(BASE_DIR / "build_ats_resume.py"),
            "--job-id",
            job_id,
        ]
        env = os.environ.copy()
        env["JOB_AGENT_CONFIG"] = str(CONFIG_PATH)
        try:
            result = subprocess.run(cmd, cwd=str(BASE_DIR.parent), check=True, capture_output=True, text=True, timeout=180, env=env)
        except subprocess.CalledProcessError as exc:
            self.send_text(500, f"Resume generation failed:\n{exc.stderr or exc.stdout}")
            return
        except Exception as exc:
            self.send_text(500, f"Resume generation failed: {exc}")
            return
        output_text = result.stdout.strip().splitlines()[-1] if result.stdout.strip() else ""
        output = Path(output_text).expanduser() if output_text else DATA_DIR / "generated_resumes" / safe_download_name(f"tailored_resume_{job_id}")
        if not output.is_absolute():
            output = BASE_DIR.parent / output
        self.serve_file(
            output,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            output.name,
        )


def rebase_dashboard_resume_links(html_text: str) -> str:
    return re.sub(r"https?://127\.0\.0\.1:\d+(/resume/)", r"\1", html_text)


def inject_dashboard_run_control(html_text: str) -> str:
    html_text = inject_roleforge_branding(html_text)
    html_text = inject_roleforge_footer(html_text)
    html_text = inject_candidate_profile(html_text)
    html_text = inject_resume_insights(html_text)
    html_text = inject_dashboard_download_labels(html_text)
    html_text = inject_job_detail_buttons(html_text)
    html_text = inject_resume_download_loader(html_text)
    if "runAgentButton" not in html_text:
        html_text = html_text.replace(
            '<div class="nav-meta">Local dashboard · on-demand ATS resumes</div>',
            '<div class="nav-actions"><button class="button secondary" id="settingsButton" type="button">Config</button>'
            '<button class="button secondary" id="runAgentButton" type="button">Run fresh report</button>'
            '<span class="run-status" id="runAgentStatus"></span>'
            '<span class="run-loader" id="runLoader" aria-live="polite" aria-label="Fresh report progress">'
            '<span class="run-loader-ring"><span class="run-loader-percent" id="runProgressPercent">1%</span></span>'
            '<span class="run-loader-label" id="runProgressLabel">Starting</span></span></div>',
        )
    if "settingsButton" not in html_text:
        html_text = html_text.replace(
            '<div class="nav-actions">',
            '<div class="nav-actions"><button class="button secondary" id="settingsButton" type="button">Config</button>',
            1,
        )
    if "resumeChangeButton" not in html_text:
        html_text = html_text.replace(
            '<button class="button secondary" id="settingsButton" type="button">Config</button>',
            '<button class="button secondary" id="settingsButton" type="button">Config</button>'
            '<button class="button secondary" id="resumeChangeButton" type="button">Change resume</button>',
            1,
        )
    if "runLoader" not in html_text:
        html_text = html_text.replace(
            '<span class="run-status" id="runAgentStatus"></span>',
            '<span class="run-status" id="runAgentStatus"></span>'
            '<span class="run-loader" id="runLoader" aria-live="polite" aria-label="Fresh report progress">'
            '<span class="run-loader-ring"><span class="run-loader-percent" id="runProgressPercent">1%</span></span>'
            '<span class="run-loader-label" id="runProgressLabel">Starting</span></span>',
        )
    if ".nav-actions" not in html_text:
        html_text = html_text.replace(
            ".nav-meta { color: var(--muted); font-size: 0.92rem; }",
            ".nav-meta { color: var(--muted); font-size: 0.92rem; }\n"
            "    .nav-actions { display: inline-flex; align-items: center; gap: 10px; }\n"
            "    .run-status { min-width: 96px; color: var(--muted); font-size: .82rem; }",
        )
    if ".run-loader" not in html_text:
        marker = "    .run-status { min-width: 96px; color: var(--muted); font-size: .82rem; }"
        html_text = html_text.replace(
            marker,
            marker + "\n"
            "    .run-loader { --progress-angle: 3.6deg; display: none; align-items: center; gap: 8px; color: var(--accent); font-size: .82rem; font-weight: 950; }\n"
            "    .run-loader.is-active { display: inline-flex; }\n"
            "    .run-loader-ring { position: relative; width: 34px; height: 34px; border-radius: 50%; display: grid; place-items: center; background: conic-gradient(var(--accent) var(--progress-angle), rgba(255,255,255,.14) 0); box-shadow: 0 0 0 1px rgba(255,210,31,.2); }\n"
            "    .run-loader-ring::after { content: \"\"; width: 24px; height: 24px; border-radius: 50%; background: #050505; position: absolute; }\n"
            "    .run-loader-percent { position: relative; z-index: 1; color: var(--ink); font-size: 9px; line-height: 1; }",
        )
    if ".settings-modal" not in html_text:
        html_text = html_text.replace(
            "</style>",
            """
    .settings-modal { position: fixed; inset: 0; z-index: 50; display: none; align-items: center; justify-content: center; padding: 20px; background: rgba(0,0,0,.72); }
    .settings-modal.is-open { display: flex; }
    .settings-panel { width: min(760px, 100%); max-height: min(760px, calc(100vh - 40px)); overflow: auto; border: 1px solid rgba(255,210,31,.24); border-radius: 8px; background: #101010; box-shadow: var(--shadow); }
    .settings-head { min-height: 58px; display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 14px 16px; border-bottom: 1px solid var(--line); }
    .settings-title { margin: 0; color: var(--ink); font-size: 1rem; }
    .settings-body { display: grid; gap: 14px; padding: 16px; }
    .settings-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; }
    .settings-field { display: grid; gap: 6px; }
    .settings-field label, .sources-title { color: var(--accent); font-size: 11px; font-weight: 950; text-transform: uppercase; }
    .settings-field textarea { width: 100%; min-height: 86px; border: 1px solid rgba(255,210,31,.24); border-radius: 8px; background: var(--panel); color: var(--ink); padding: 11px 13px; font: inherit; resize: vertical; }
    .sources-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
    .source-option { min-height: 38px; display: flex; align-items: center; gap: 8px; padding: 8px 10px; border: 1px solid var(--line); border-radius: 8px; background: rgba(255,255,255,.04); color: var(--ink); font-size: 13px; }
    .source-option input { width: 16px; min-height: 16px; accent-color: var(--accent); }
    .settings-actions { display: flex; justify-content: flex-end; gap: 10px; padding: 14px 16px; border-top: 1px solid var(--line); }
    .settings-message { min-height: 18px; color: var(--muted); font-size: 12px; }
    @media (max-width: 760px) { .settings-grid, .sources-grid { grid-template-columns: 1fr; } }
  </style>""",
        )
    if "settingsModal" not in html_text:
        modal = """
  <div class="settings-modal" id="settingsModal" aria-hidden="true">
    <section class="settings-panel" role="dialog" aria-modal="true" aria-labelledby="settingsTitle">
      <div class="settings-head">
        <h2 class="settings-title" id="settingsTitle">Configuration</h2>
        <button class="button secondary" id="settingsCloseButton" type="button">Close</button>
      </div>
      <form id="settingsForm">
        <div class="settings-body">
          <div class="settings-grid">
            <div class="settings-field">
              <label for="settingsSalary">Minimum salary</label>
              <input id="settingsSalary" name="salary" type="number" min="0" step="100000">
            </div>
            <div class="settings-field">
              <label for="settingsCurrency">Currency</label>
              <select id="settingsCurrency" name="currency">
                <option value="INR">INR</option>
                <option value="USD">USD</option>
              </select>
            </div>
            <div class="settings-field">
              <label for="settingsMinScore">Minimum score</label>
              <input id="settingsMinScore" name="min_score" type="number" min="0" step="1">
            </div>
          </div>
          <div class="settings-grid">
            <div class="settings-field">
              <label for="settingsFreshness">Freshness</label>
              <select id="settingsFreshness" name="freshness">
                <option value="today">Today only</option>
                <option value="all">All fetched jobs</option>
              </select>
            </div>
            <div class="settings-field">
              <label for="settingsPeriod">Salary period</label>
              <select id="settingsPeriod" name="period">
                <option value="year">Year</option>
                <option value="month">Month</option>
              </select>
            </div>
            <div class="settings-field">
              <label>Status</label>
              <div class="settings-message" id="settingsMessage"></div>
            </div>
          </div>
          <div class="settings-field">
            <label for="settingsRoles">Roles</label>
            <textarea id="settingsRoles" name="roles" placeholder="One role per line"></textarea>
          </div>
          <div class="settings-field">
            <label for="settingsLocations">Location preferences</label>
            <textarea id="settingsLocations" name="locations" placeholder="One location per line"></textarea>
          </div>
          <div>
            <div class="sources-title">Job search sites</div>
            <div class="sources-grid" id="settingsSources"></div>
          </div>
        </div>
        <div class="settings-actions">
          <button class="button secondary" id="settingsCancelButton" type="button">Cancel</button>
          <button class="button primary" id="settingsSaveButton" type="submit">Save configuration</button>
        </div>
      </form>
    </section>
  </div>
"""
        html_text = html_text.replace("</body>", f"{modal}</body>")
    if ".resume-upload-modal" not in html_text:
        html_text = html_text.replace(
            "</style>",
            """
    .resume-upload-modal { position: fixed; inset: 0; z-index: 55; display: none; align-items: center; justify-content: center; padding: 20px; background: rgba(0,0,0,.72); }
    .resume-upload-modal.is-open { display: flex; }
    .resume-upload-panel { width: min(560px, 100%); border: 1px solid rgba(255,210,31,.24); border-radius: 8px; background: #101010; box-shadow: var(--shadow); }
    .resume-upload-head { min-height: 58px; display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 14px 16px; border-bottom: 1px solid var(--line); }
    .resume-upload-title { margin: 0; color: var(--ink); font-size: 1rem; }
    .resume-upload-body { display: grid; gap: 14px; padding: 16px; }
    .resume-dropzone { min-height: 170px; display: grid; place-items: center; gap: 8px; padding: 20px; border: 1px dashed rgba(255,210,31,.48); border-radius: 8px; background: rgba(255,255,255,.04); color: var(--ink); text-align: center; cursor: pointer; }
    .resume-dropzone.is-dragging { border-color: var(--accent); background: rgba(255,210,31,.09); }
    .resume-dropzone strong { display: block; color: var(--accent); font-size: 15px; }
    .resume-dropzone span { display: block; color: var(--muted); font-size: 12px; }
    .resume-file-name { min-height: 18px; color: var(--muted); font-size: 12px; }
    .resume-upload-message { min-height: 18px; color: var(--muted); font-size: 12px; }
    .resume-upload-actions { display: flex; justify-content: flex-end; gap: 10px; padding: 14px 16px; border-top: 1px solid var(--line); }
  </style>""",
        )
    if "resumeUploadModal" not in html_text:
        modal = """
  <div class="resume-upload-modal" id="resumeUploadModal" aria-hidden="true">
    <section class="resume-upload-panel" role="dialog" aria-modal="true" aria-labelledby="resumeUploadTitle">
      <div class="resume-upload-head">
        <h2 class="resume-upload-title" id="resumeUploadTitle">Change resume</h2>
        <button class="button secondary" id="resumeUploadCloseButton" type="button">Close</button>
      </div>
      <form id="resumeUploadForm">
        <div class="resume-upload-body">
          <label class="resume-dropzone" id="resumeDropzone" for="resumeUploadInput">
            <input id="resumeUploadInput" name="resume" type="file" accept=".pdf,.docx,.txt,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,text/plain" hidden>
            <span>
              <strong>Drop your resume here</strong>
              <span>PDF, DOCX, or TXT. This updates config.json and runs a fresh report.</span>
            </span>
          </label>
          <div class="resume-file-name" id="resumeFileName">No file selected</div>
          <div class="resume-upload-message" id="resumeUploadMessage"></div>
        </div>
        <div class="resume-upload-actions">
          <button class="button secondary" id="resumeUploadCancelButton" type="button">Cancel</button>
          <button class="button primary" id="resumeUploadSubmitButton" type="submit">Upload and run</button>
        </div>
      </form>
    </section>
  </div>
"""
        html_text = html_text.replace("</body>", f"{modal}</body>")
    if "serverInjectedRunFreshReport" not in html_text:
        script = """
  <script>
    let serverInjectedRunProgressTimer = null;
    let serverInjectedRunProgressValue = 1;
    function serverInjectedSetRunProgress(value, label) {
      const loader = document.getElementById('runLoader');
      const percent = document.getElementById('runProgressPercent');
      const progressLabel = document.getElementById('runProgressLabel');
      serverInjectedRunProgressValue = Math.max(1, Math.min(100, Math.round(value)));
      if (loader) {
        loader.classList.add('is-active');
        loader.style.setProperty('--progress-angle', `${serverInjectedRunProgressValue * 3.6}deg`);
      }
      if (percent) percent.textContent = `${serverInjectedRunProgressValue}%`;
      if (progressLabel && label) progressLabel.textContent = label;
    }
    function serverInjectedStopRunProgress(hide = true) {
      if (serverInjectedRunProgressTimer) {
        window.clearInterval(serverInjectedRunProgressTimer);
        serverInjectedRunProgressTimer = null;
      }
      const loader = document.getElementById('runLoader');
      if (hide && loader) loader.classList.remove('is-active');
    }
    function serverInjectedStartRunProgress(label) {
      serverInjectedStopRunProgress(false);
      serverInjectedSetRunProgress(1, label || 'Starting');
      serverInjectedRunProgressTimer = window.setInterval(() => {
        const step = serverInjectedRunProgressValue < 70 ? 3 : serverInjectedRunProgressValue < 92 ? 2 : 1;
        if (serverInjectedRunProgressValue < 98) serverInjectedSetRunProgress(serverInjectedRunProgressValue + step, label || 'Running');
      }, 650);
    }
    async function serverInjectedRunFreshReport(event) {
      if (event) {
        event.preventDefault();
        event.stopImmediatePropagation();
      }
      const runAgentButton = document.getElementById('runAgentButton');
      const runAgentStatus = document.getElementById('runAgentStatus');
      if (!runAgentButton) return;
      if (runAgentButton.disabled) return;
      runAgentButton.disabled = true;
      runAgentButton.textContent = 'Running...';
      if (runAgentStatus) runAgentStatus.textContent = 'Searching jobs';
      serverInjectedStartRunProgress('Searching jobs');
      try {
        const response = await fetch('/run-agent', { method: 'POST' });
        const payload = await response.json();
        if (!response.ok || !payload.ok) {
          throw new Error(payload.message || payload.stderr || 'Run failed');
        }
        if (runAgentStatus) runAgentStatus.textContent = 'Refreshing';
        serverInjectedStopRunProgress(false);
        serverInjectedSetRunProgress(100, 'Refreshing');
        window.setTimeout(() => window.location.reload(), 450);
      } catch (error) {
        serverInjectedStopRunProgress();
        if (runAgentStatus) runAgentStatus.textContent = error.message || 'Run failed';
        runAgentButton.disabled = false;
        runAgentButton.textContent = 'Run fresh report';
      }
    }
    const serverInjectedRunAgentButton = document.getElementById('runAgentButton');
    if (serverInjectedRunAgentButton) serverInjectedRunAgentButton.addEventListener('click', serverInjectedRunFreshReport, true);
  </script>
"""
        html_text = html_text.replace("</body>", f"{script}</body>")
    if "serverInjectedSettingsModal" not in html_text:
        script = """
  <script>
    const serverInjectedSettingsModal = document.getElementById('settingsModal');
    const serverInjectedSettingsButton = document.getElementById('settingsButton');
    const serverInjectedSettingsForm = document.getElementById('settingsForm');
    const serverInjectedSettingsMessage = document.getElementById('settingsMessage');
    const serverInjectedSources = document.getElementById('settingsSources');

    function serverInjectedSetSettingsMessage(message) {
      if (serverInjectedSettingsMessage) serverInjectedSettingsMessage.textContent = message || '';
    }
    function serverInjectedOpenSettings() {
      if (!serverInjectedSettingsModal) return;
      serverInjectedSettingsModal.classList.add('is-open');
      serverInjectedSettingsModal.setAttribute('aria-hidden', 'false');
      serverInjectedLoadSettings();
    }
    function serverInjectedCloseSettings() {
      if (!serverInjectedSettingsModal) return;
      serverInjectedSettingsModal.classList.remove('is-open');
      serverInjectedSettingsModal.setAttribute('aria-hidden', 'true');
    }
    function serverInjectedLines(value) {
      return Array.isArray(value) ? value.join('\\n') : '';
    }
    function serverInjectedCheckedSources() {
      return Array.from(document.querySelectorAll('[data-source-key]')).reduce((acc, input) => {
        acc[input.dataset.sourceKey] = input.checked;
        return acc;
      }, {});
    }
    async function serverInjectedLoadSettings() {
      serverInjectedSetSettingsMessage('Loading');
      const response = await fetch('/settings');
      const payload = await response.json();
      if (!response.ok || !payload.ok) throw new Error(payload.message || 'Could not load settings');
      const settings = payload.settings;
      document.getElementById('settingsSalary').value = settings.expected_salary.amount || 0;
      document.getElementById('settingsCurrency').value = settings.expected_salary.currency || 'INR';
      document.getElementById('settingsPeriod').value = settings.expected_salary.period || 'year';
      document.getElementById('settingsMinScore').value = settings.min_score || 0;
      document.getElementById('settingsFreshness').value = settings.freshness || 'today';
      document.getElementById('settingsRoles').value = serverInjectedLines(settings.roles);
      document.getElementById('settingsLocations').value = serverInjectedLines(settings.locations);
      if (serverInjectedSources) {
        serverInjectedSources.innerHTML = '';
        settings.sources.forEach((source) => {
          const label = document.createElement('label');
          label.className = 'source-option';
          label.innerHTML = `<input type="checkbox" data-source-key="${source.key}"> <span>${source.label}</span>`;
          const input = label.querySelector('input');
          input.checked = Boolean(source.enabled);
          serverInjectedSources.appendChild(label);
        });
      }
      serverInjectedSetSettingsMessage('Ready');
    }
    async function serverInjectedSaveSettings(event) {
      event.preventDefault();
      serverInjectedSetSettingsMessage('Saving');
      const body = {
        expected_salary: {
          amount: Number(document.getElementById('settingsSalary').value || 0),
          currency: document.getElementById('settingsCurrency').value,
          period: document.getElementById('settingsPeriod').value
        },
        min_score: Number(document.getElementById('settingsMinScore').value || 0),
        freshness: document.getElementById('settingsFreshness').value,
        roles: document.getElementById('settingsRoles').value.split('\\n'),
        locations: document.getElementById('settingsLocations').value.split('\\n'),
        sources: serverInjectedCheckedSources()
      };
      const response = await fetch('/settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
      });
      const payload = await response.json();
      if (!response.ok || !payload.ok) {
        serverInjectedSetSettingsMessage(payload.message || 'Save failed');
        return;
      }
      serverInjectedSetSettingsMessage('Saved for next run');
      window.setTimeout(serverInjectedCloseSettings, 500);
    }
    if (serverInjectedSettingsButton) serverInjectedSettingsButton.addEventListener('click', serverInjectedOpenSettings);
    if (serverInjectedSettingsForm) serverInjectedSettingsForm.addEventListener('submit', serverInjectedSaveSettings);
    const serverInjectedCloseButton = document.getElementById('settingsCloseButton');
    const serverInjectedCancelButton = document.getElementById('settingsCancelButton');
    if (serverInjectedCloseButton) serverInjectedCloseButton.addEventListener('click', serverInjectedCloseSettings);
    if (serverInjectedCancelButton) serverInjectedCancelButton.addEventListener('click', serverInjectedCloseSettings);
    if (serverInjectedSettingsModal) {
      serverInjectedSettingsModal.addEventListener('click', (event) => {
        if (event.target === serverInjectedSettingsModal) serverInjectedCloseSettings();
      });
    }
  </script>
"""
        html_text = html_text.replace("</body>", f"{script}</body>")
    if "serverInjectedResumeUpload" not in html_text:
        script = """
  <script>
    const serverInjectedResumeUploadModal = document.getElementById('resumeUploadModal');
    const serverInjectedResumeChangeButton = document.getElementById('resumeChangeButton');
    const serverInjectedResumeUploadForm = document.getElementById('resumeUploadForm');
    const serverInjectedResumeUploadInput = document.getElementById('resumeUploadInput');
    const serverInjectedResumeDropzone = document.getElementById('resumeDropzone');
    const serverInjectedResumeFileName = document.getElementById('resumeFileName');
    const serverInjectedResumeUploadMessage = document.getElementById('resumeUploadMessage');
    const serverInjectedResumeUploadSubmit = document.getElementById('resumeUploadSubmitButton');

    function serverInjectedSetResumeUploadMessage(message) {
      if (serverInjectedResumeUploadMessage) serverInjectedResumeUploadMessage.textContent = message || '';
    }
    function serverInjectedResumeSelectedFile() {
      return serverInjectedResumeUploadInput && serverInjectedResumeUploadInput.files && serverInjectedResumeUploadInput.files[0];
    }
    function serverInjectedUpdateResumeFileName() {
      const file = serverInjectedResumeSelectedFile();
      if (serverInjectedResumeFileName) serverInjectedResumeFileName.textContent = file ? file.name : 'No file selected';
    }
    function serverInjectedOpenResumeUpload() {
      if (!serverInjectedResumeUploadModal) return;
      serverInjectedResumeUploadModal.classList.add('is-open');
      serverInjectedResumeUploadModal.setAttribute('aria-hidden', 'false');
      serverInjectedSetResumeUploadMessage('');
      serverInjectedUpdateResumeFileName();
    }
    function serverInjectedCloseResumeUpload() {
      if (!serverInjectedResumeUploadModal) return;
      serverInjectedResumeUploadModal.classList.remove('is-open');
      serverInjectedResumeUploadModal.setAttribute('aria-hidden', 'true');
    }
    function serverInjectedStartReportProgress(label) {
      const runAgentButton = document.getElementById('runAgentButton');
      const runAgentStatus = document.getElementById('runAgentStatus');
      if (runAgentButton) {
        runAgentButton.disabled = true;
        runAgentButton.textContent = 'Running...';
      }
      if (runAgentStatus) runAgentStatus.textContent = label || 'Searching jobs';
      if (typeof serverInjectedStartRunProgress === 'function') serverInjectedStartRunProgress(label || 'Searching jobs');
      else if (typeof startRunProgress === 'function') startRunProgress(label || 'Searching jobs');
    }
    function serverInjectedFinishReportProgress(label) {
      const runAgentStatus = document.getElementById('runAgentStatus');
      if (runAgentStatus) runAgentStatus.textContent = label || 'Refreshing';
      if (typeof serverInjectedStopRunProgress === 'function') serverInjectedStopRunProgress(false);
      if (typeof serverInjectedSetRunProgress === 'function') serverInjectedSetRunProgress(100, label || 'Refreshing');
      else if (typeof completeRunProgress === 'function') completeRunProgress(label || 'Refreshing');
    }
    async function serverInjectedUploadResumeAndRun(event) {
      event.preventDefault();
      const file = serverInjectedResumeSelectedFile();
      if (!file) {
        serverInjectedSetResumeUploadMessage('Choose a PDF, DOCX, or TXT resume first.');
        return;
      }
      const allowed = ['.pdf', '.docx', '.txt'];
      const name = file.name.toLowerCase();
      if (!allowed.some((suffix) => name.endsWith(suffix))) {
        serverInjectedSetResumeUploadMessage('Upload a PDF, DOCX, or TXT resume.');
        return;
      }
      if (serverInjectedResumeUploadSubmit) serverInjectedResumeUploadSubmit.disabled = true;
      serverInjectedSetResumeUploadMessage('Uploading resume');
      try {
        const formData = new FormData();
        formData.append('resume', file);
        const uploadResponse = await fetch('/resume-upload', { method: 'POST', body: formData });
        const uploadPayload = await uploadResponse.json();
        if (!uploadResponse.ok || !uploadPayload.ok) {
          throw new Error(uploadPayload.message || 'Resume upload failed');
        }
        serverInjectedSetResumeUploadMessage('Resume updated. Running a fresh report.');
        serverInjectedCloseResumeUpload();
        serverInjectedStartReportProgress('Parsing new resume');
        const runResponse = await fetch('/run-agent', { method: 'POST' });
        const runPayload = await runResponse.json();
        if (!runResponse.ok || !runPayload.ok) {
          throw new Error(runPayload.message || runPayload.stderr || 'Fresh report failed');
        }
        serverInjectedFinishReportProgress('Refreshing');
        window.setTimeout(() => window.location.reload(), 450);
      } catch (error) {
        serverInjectedSetResumeUploadMessage(error.message || 'Resume update failed');
        const runAgentButton = document.getElementById('runAgentButton');
        const runAgentStatus = document.getElementById('runAgentStatus');
        if (runAgentButton) {
          runAgentButton.disabled = false;
          runAgentButton.textContent = 'Run fresh report';
        }
        if (runAgentStatus) runAgentStatus.textContent = error.message || 'Resume update failed';
        if (typeof serverInjectedStopRunProgress === 'function') serverInjectedStopRunProgress();
        else if (typeof stopRunProgress === 'function') stopRunProgress();
      } finally {
        if (serverInjectedResumeUploadSubmit) serverInjectedResumeUploadSubmit.disabled = false;
      }
    }
    if (serverInjectedResumeChangeButton) serverInjectedResumeChangeButton.addEventListener('click', serverInjectedOpenResumeUpload);
    if (serverInjectedResumeUploadInput) serverInjectedResumeUploadInput.addEventListener('change', serverInjectedUpdateResumeFileName);
    if (serverInjectedResumeUploadForm) serverInjectedResumeUploadForm.addEventListener('submit', serverInjectedUploadResumeAndRun);
    const serverInjectedResumeUploadClose = document.getElementById('resumeUploadCloseButton');
    const serverInjectedResumeUploadCancel = document.getElementById('resumeUploadCancelButton');
    if (serverInjectedResumeUploadClose) serverInjectedResumeUploadClose.addEventListener('click', serverInjectedCloseResumeUpload);
    if (serverInjectedResumeUploadCancel) serverInjectedResumeUploadCancel.addEventListener('click', serverInjectedCloseResumeUpload);
    if (serverInjectedResumeUploadModal) {
      serverInjectedResumeUploadModal.addEventListener('click', (event) => {
        if (event.target === serverInjectedResumeUploadModal) serverInjectedCloseResumeUpload();
      });
    }
    if (serverInjectedResumeDropzone && serverInjectedResumeUploadInput) {
      ['dragenter', 'dragover'].forEach((name) => {
        serverInjectedResumeDropzone.addEventListener(name, (event) => {
          event.preventDefault();
          serverInjectedResumeDropzone.classList.add('is-dragging');
        });
      });
      ['dragleave', 'drop'].forEach((name) => {
        serverInjectedResumeDropzone.addEventListener(name, (event) => {
          event.preventDefault();
          serverInjectedResumeDropzone.classList.remove('is-dragging');
        });
      });
      serverInjectedResumeDropzone.addEventListener('drop', (event) => {
        const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
        if (!file) return;
        const transfer = new DataTransfer();
        transfer.items.add(file);
        serverInjectedResumeUploadInput.files = transfer.files;
        serverInjectedUpdateResumeFileName();
      });
    }
  </script>
"""
        html_text = html_text.replace("</body>", f"{script}</body>")
    return html_text


def inject_resume_download_loader(html_text: str) -> str:
    if ".resume-download-overlay" not in html_text:
        html_text = html_text.replace(
            "    .download-button { min-height: 34px; padding: 7px 10px; font-size: 12px; }",
            "    .download-button { min-height: 34px; padding: 7px 10px; font-size: 12px; }\n"
            "    body.is-blocked { overflow: hidden; }\n"
            "    .resume-download-overlay { position: fixed; inset: 0; z-index: 80; display: none; place-items: center; padding: 20px; background: rgba(0,0,0,.78); backdrop-filter: blur(10px); }\n"
            "    .resume-download-overlay.is-active { display: grid; }\n"
            "    .resume-download-panel { width: min(360px, 100%); display: grid; justify-items: center; gap: 14px; padding: 24px; border: 1px solid rgba(255,210,31,.28); border-radius: 8px; background: #101010; box-shadow: var(--shadow); text-align: center; }\n"
            "    .resume-download-ring { --download-progress-angle: 3.6deg; position: relative; width: 92px; height: 92px; border-radius: 50%; display: grid; place-items: center; background: conic-gradient(var(--accent) var(--download-progress-angle), rgba(255,255,255,.14) 0); box-shadow: 0 0 0 1px rgba(255,210,31,.18), 0 18px 48px rgba(0,0,0,.45); }\n"
            "    .resume-download-ring::after { content: \"\"; position: absolute; width: 68px; height: 68px; border-radius: 50%; background: #050505; }\n"
            "    .resume-download-percent { position: relative; z-index: 1; color: var(--ink); font-size: 18px; font-weight: 950; }\n"
            "    .resume-download-title { color: var(--accent); font-size: 12px; font-weight: 950; text-transform: uppercase; }\n"
            "    .resume-download-message { color: var(--ink); font-size: 16px; font-weight: 900; }\n"
            "    .resume-download-sub { max-width: 280px; color: var(--muted); font-size: 12px; }",
        )
    if "resumeDownloadOverlay" not in html_text:
        modal = """
  <div class="resume-download-overlay" id="resumeDownloadOverlay" aria-live="polite" aria-modal="true" role="dialog" aria-label="Preparing tailored resume">
    <div class="resume-download-panel">
      <div class="resume-download-ring" id="resumeDownloadRing"><span class="resume-download-percent" id="resumeDownloadPercent">1%</span></div>
      <div class="resume-download-title">Tailored resume</div>
      <div class="resume-download-message" id="resumeDownloadMessage">Connecting to OpenAI</div>
      <div class="resume-download-sub">Please keep this window open while your job-specific resume is prepared.</div>
    </div>
  </div>
"""
        html_text = html_text.replace("</body>", f"{modal}</body>")
    if "serverInjectedResumeDownload" not in html_text and "downloadTailoredResume" not in html_text:
        script = """
  <script>
    const serverInjectedResumeDownloadOverlay = document.getElementById('resumeDownloadOverlay');
    const serverInjectedResumeDownloadRing = document.getElementById('resumeDownloadRing');
    const serverInjectedResumeDownloadPercent = document.getElementById('resumeDownloadPercent');
    const serverInjectedResumeDownloadMessage = document.getElementById('resumeDownloadMessage');
    let serverInjectedResumeDownloadTimer = null;
    let serverInjectedResumeDownloadMessageTimer = null;
    let serverInjectedResumeDownloadProgress = 1;
    const serverInjectedResumeDownloadMessages = [
      'Connecting to OpenAI',
      'Reading the job description',
      'Creating your tailored resume',
      'Publishing the download'
    ];
    function serverInjectedSetResumeDownloadProgress(value, message) {
      serverInjectedResumeDownloadProgress = Math.max(1, Math.min(100, Math.round(value)));
      if (serverInjectedResumeDownloadRing) serverInjectedResumeDownloadRing.style.setProperty('--download-progress-angle', `${serverInjectedResumeDownloadProgress * 3.6}deg`);
      if (serverInjectedResumeDownloadPercent) serverInjectedResumeDownloadPercent.textContent = `${serverInjectedResumeDownloadProgress}%`;
      if (serverInjectedResumeDownloadMessage && message) serverInjectedResumeDownloadMessage.textContent = message;
    }
    function serverInjectedStopResumeDownloadLoader(hide = true) {
      if (serverInjectedResumeDownloadTimer) {
        window.clearInterval(serverInjectedResumeDownloadTimer);
        serverInjectedResumeDownloadTimer = null;
      }
      if (serverInjectedResumeDownloadMessageTimer) {
        window.clearInterval(serverInjectedResumeDownloadMessageTimer);
        serverInjectedResumeDownloadMessageTimer = null;
      }
      if (hide) {
        if (serverInjectedResumeDownloadOverlay) serverInjectedResumeDownloadOverlay.classList.remove('is-active');
        document.body.classList.remove('is-blocked');
      }
    }
    function serverInjectedStartResumeDownloadLoader() {
      serverInjectedStopResumeDownloadLoader(false);
      document.body.classList.add('is-blocked');
      if (serverInjectedResumeDownloadOverlay) serverInjectedResumeDownloadOverlay.classList.add('is-active');
      serverInjectedSetResumeDownloadProgress(1, serverInjectedResumeDownloadMessages[0]);
      let messageIndex = 0;
      serverInjectedResumeDownloadMessageTimer = window.setInterval(() => {
        messageIndex = Math.min(messageIndex + 1, serverInjectedResumeDownloadMessages.length - 1);
        serverInjectedSetResumeDownloadProgress(Math.max(serverInjectedResumeDownloadProgress, 20 + messageIndex * 22), serverInjectedResumeDownloadMessages[messageIndex]);
      }, 1700);
      serverInjectedResumeDownloadTimer = window.setInterval(() => {
        const step = serverInjectedResumeDownloadProgress < 55 ? 4 : serverInjectedResumeDownloadProgress < 88 ? 2 : 1;
        if (serverInjectedResumeDownloadProgress < 96) serverInjectedSetResumeDownloadProgress(serverInjectedResumeDownloadProgress + step);
      }, 420);
    }
    function serverInjectedFilenameFromDisposition(value) {
      const match = /filename\\*?=(?:UTF-8''|")?([^";]+)/i.exec(value || '');
      if (!match) return '';
      return decodeURIComponent(match[1].replace(/"/g, '').trim());
    }
    async function serverInjectedResumeDownload(event) {
      event.preventDefault();
      const link = event.currentTarget;
      if (!link || link.dataset.loading === 'true') return;
      link.dataset.loading = 'true';
      serverInjectedStartResumeDownloadLoader();
      try {
        const response = await fetch(link.href);
        if (!response.ok) {
          const message = await response.text();
          throw new Error(message || 'Resume generation failed');
        }
        serverInjectedSetResumeDownloadProgress(92, 'Publishing the download');
        const blob = await response.blob();
        const fileName = serverInjectedFilenameFromDisposition(response.headers.get('Content-Disposition')) || 'tailored_resume.docx';
        const downloadUrl = URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = downloadUrl;
        anchor.download = fileName;
        document.body.appendChild(anchor);
        serverInjectedSetResumeDownloadProgress(100, 'Download ready');
        anchor.click();
        anchor.remove();
        window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 30000);
        window.setTimeout(() => serverInjectedStopResumeDownloadLoader(), 650);
      } catch (error) {
        serverInjectedSetResumeDownloadProgress(100, error.message || 'Resume generation failed');
        window.setTimeout(() => serverInjectedStopResumeDownloadLoader(), 2400);
      } finally {
        link.dataset.loading = 'false';
      }
    }
    document.querySelectorAll('.download-button').forEach((link) => link.addEventListener('click', serverInjectedResumeDownload));
  </script>
"""
        html_text = html_text.replace("</body>", f"{script}</body>")
    return html_text


def inject_dashboard_download_labels(html_text: str) -> str:
    html_text = html_text.replace("<th>Resume</th>", "<th>Download tailored resume</th>")
    html_text = re.sub(
        r'<a class="button primary" href="([^"]*/resume/[^"]*)">Download ATS[- ]friendly resume</a>',
        r'<a class="button primary download-button" href="\1">Download</a>',
        html_text,
    )
    html_text = re.sub(
        r'<a class="button primary" href="([^"]*/resume/[^"]*)">Download</a>',
        r'<a class="button primary download-button" href="\1">Download</a>',
        html_text,
    )
    if ".download-button" not in html_text:
        html_text = html_text.replace(
            "    .button:disabled { cursor: wait; opacity: .7; transform: none; }",
            "    .button:disabled { cursor: wait; opacity: .7; transform: none; }\n"
            "    .download-button { min-height: 34px; padding: 7px 10px; font-size: 12px; }",
        )
    return html_text


def experience_range_from_job_text(text: str) -> str:
    text = clean_dashboard_text(text)
    if not text:
        return "Not specified"
    lowered = text.lower()
    if re.search(r"\b(fresher|freshers|entry[- ]level|graduate trainee|internship)\b", lowered):
        return "0-2 years"
    range_match = re.search(
        r"\b(\d{1,2})\s*(?:\+)?\s*(?:-|–|—|to)\s*(\d{1,2})\s*(?:\+)?\s*(?:years?|yrs?)\b",
        lowered,
        re.I,
    )
    if range_match:
        return f"{range_match.group(1)}-{range_match.group(2)} years"
    plus_match = re.search(
        r"\b(?:minimum|min\.?|at least|over|more than)?\s*(\d{1,2})\s*\+?\s*(?:years?|yrs?)\b",
        lowered,
        re.I,
    )
    if plus_match:
        return f"{plus_match.group(1)}+ years"
    return "Not specified"


def load_jobs_by_uid_for_dashboard() -> dict:
    latest_path = DATA_DIR / "latest_jobs.json"
    if not latest_path.exists():
        return {}
    try:
        payload = json.loads(latest_path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    jobs = {}
    for job in payload.get("jobs", []) or []:
        uid = str(job.get("job_uid") or "")
        if uid:
            jobs[uid] = job
    return jobs


def job_detail_button_html(job: dict) -> str:
    title = clean_dashboard_text(job.get("title", ""))
    company = clean_dashboard_text(job.get("company", ""))
    description = clean_dashboard_text(job.get("description", "")) or "No job description was available from the source."
    experience = clean_dashboard_text(job.get("experience_range", "")) or experience_range_from_job_text(
        " ".join([title, description])
    )
    label = f"Show details for {title or 'this job'}"
    if company:
        label += f" at {company}"
    return (
        f'<button class="job-info-button" type="button" aria-label="{html_escape(label)}" '
        f'data-title="{html_escape(title or "Job details")}" data-company="{html_escape(company)}" '
        f'data-experience="{html_escape(experience)}" data-description="{html_escape(description)}">i</button>'
    )


def job_detail_css() -> str:
    return """
    .job-title-line { display: flex; align-items: flex-start; gap: 8px; }
    .job-info-button { width: 20px; height: 20px; flex: 0 0 auto; display: inline-grid; place-items: center; border: 1px solid rgba(255,210,31,.5); border-radius: 50%; background: rgba(255,210,31,.08); color: var(--accent); font: inherit; font-size: 11px; font-weight: 950; cursor: help; }
    .job-info-button:hover, .job-info-button:focus-visible { background: var(--accent); color: #050505; outline: none; }
    .job-detail-popover { position: fixed; z-index: 90; width: min(460px, calc(100vw - 28px)); max-height: min(420px, calc(100vh - 28px)); display: none; overflow: auto; padding: 13px; border: 1px solid rgba(255,210,31,.32); border-radius: 8px; background: #101010; color: var(--ink); box-shadow: var(--shadow); }
    .job-detail-popover.is-open { display: block; }
    .job-detail-title { color: var(--accent); font-size: 14px; font-weight: 950; line-height: 1.25; }
    .job-detail-meta { margin-top: 6px; color: var(--muted); font-size: 12px; font-weight: 850; }
    .job-detail-copy { margin-top: 10px; white-space: pre-wrap; color: rgba(255,255,255,.86); font-size: 12px; line-height: 1.45; }
"""


def job_detail_script() -> str:
    return """
  <script>
    const serverInjectedJobInfoButtons = Array.from(document.querySelectorAll('.job-info-button'));
    const serverInjectedJobDetailPopover = document.getElementById('jobDetailPopover');
    function serverInjectedPositionJobDetailPopover(button) {
      if (!serverInjectedJobDetailPopover || !button) return;
      const rect = button.getBoundingClientRect();
      const margin = 12;
      const popoverRect = serverInjectedJobDetailPopover.getBoundingClientRect();
      const width = popoverRect.width || Math.min(460, window.innerWidth - 28);
      const height = popoverRect.height || Math.min(420, window.innerHeight - 28);
      let left = rect.right + margin;
      if (left + width > window.innerWidth - margin) left = Math.max(margin, window.innerWidth - width - margin);
      let top = rect.top - 8;
      if (top + height > window.innerHeight - margin) top = Math.max(margin, window.innerHeight - height - margin);
      serverInjectedJobDetailPopover.style.left = `${left}px`;
      serverInjectedJobDetailPopover.style.top = `${top}px`;
    }
    function serverInjectedShowJobDetail(button) {
      if (!serverInjectedJobDetailPopover || !button) return;
      const titleEl = document.createElement('div');
      titleEl.className = 'job-detail-title';
      titleEl.textContent = button.dataset.title || 'Job details';
      const metaEl = document.createElement('div');
      metaEl.className = 'job-detail-meta';
      const company = button.dataset.company || '';
      metaEl.textContent = `${company ? company + ' · ' : ''}Experience required: ${button.dataset.experience || 'Not specified'}`;
      const copyEl = document.createElement('div');
      copyEl.className = 'job-detail-copy';
      copyEl.textContent = button.dataset.description || 'No job description was available from the source.';
      serverInjectedJobDetailPopover.textContent = '';
      serverInjectedJobDetailPopover.append(titleEl, metaEl, copyEl);
      serverInjectedJobDetailPopover.classList.add('is-open');
      serverInjectedJobDetailPopover.setAttribute('aria-hidden', 'false');
      button.setAttribute('aria-describedby', 'jobDetailPopover');
      serverInjectedPositionJobDetailPopover(button);
    }
    function serverInjectedHideJobDetail() {
      if (!serverInjectedJobDetailPopover) return;
      serverInjectedJobDetailPopover.classList.remove('is-open');
      serverInjectedJobDetailPopover.setAttribute('aria-hidden', 'true');
      serverInjectedJobInfoButtons.forEach((button) => button.removeAttribute('aria-describedby'));
    }
    serverInjectedJobInfoButtons.forEach((button) => {
      button.addEventListener('mouseenter', () => serverInjectedShowJobDetail(button));
      button.addEventListener('focus', () => serverInjectedShowJobDetail(button));
      button.addEventListener('mouseleave', serverInjectedHideJobDetail);
      button.addEventListener('blur', serverInjectedHideJobDetail);
      button.addEventListener('click', (event) => {
        event.stopPropagation();
        if (serverInjectedJobDetailPopover && serverInjectedJobDetailPopover.classList.contains('is-open') && button.getAttribute('aria-describedby')) {
          serverInjectedHideJobDetail();
        } else {
          serverInjectedShowJobDetail(button);
        }
      });
    });
    window.addEventListener('scroll', serverInjectedHideJobDetail, true);
    window.addEventListener('resize', serverInjectedHideJobDetail);
    document.addEventListener('click', serverInjectedHideJobDetail);
    document.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') serverInjectedHideJobDetail();
    });
  </script>
"""


def inject_job_detail_buttons(html_text: str) -> str:
    jobs_by_uid = load_jobs_by_uid_for_dashboard()
    if not jobs_by_uid:
        return html_text
    had_buttons = "job-info-button" in html_text
    if ".job-info-button" not in html_text:
        html_text = html_text.replace("</style>", f"{job_detail_css()}  </style>", 1)
    if 'id="jobDetailPopover"' not in html_text:
        html_text = html_text.replace(
            '<div class="resume-download-overlay"',
            '<div class="job-detail-popover" id="jobDetailPopover" role="tooltip" aria-hidden="true"></div>\n  <div class="resume-download-overlay"',
            1,
        )
    if "jobInfoButtons" not in html_text and "serverInjectedJobInfoButtons" not in html_text:
        html_text = html_text.replace("</body>", f"{job_detail_script()}</body>", 1)
    if had_buttons:
        return html_text

    def add_detail_button(match: re.Match) -> str:
        row = match.group(0)
        uid_match = re.search(r'/resume/([^"#?]+)', row)
        if not uid_match:
            return row
        uid = urllib.parse.unquote(uid_match.group(1))
        job = jobs_by_uid.get(uid)
        if not job:
            return row
        button = job_detail_button_html(job)
        strong_match = re.search(r"<strong>.*?</strong>", row, re.S)
        if not strong_match:
            return row
        title_markup = strong_match.group(0)
        wrapped = f'<div class="job-title-line">{title_markup}{button}</div>'
        return row[: strong_match.start()] + wrapped + row[strong_match.end() :]

    return re.sub(r"<tr\b[\s\S]*?</tr>", add_detail_button, html_text)


def roleforge_footer_html() -> str:
    year = dt.datetime.now().year
    return f"""
  <footer class="site-footer" data-roleforge-footer>
    <div class="site-footer-inner">
      <span>&copy; {year} RoleForge. All rights reserved.</span>
      <span>Developed by Chandan Ghosh: <a href="https://www.linkedin.com/in/chandan-ghosh-43350650/" target="_blank" rel="noopener">https://www.linkedin.com/in/chandan-ghosh-43350650/</a></span>
    </div>
  </footer>"""


def roleforge_footer_css() -> str:
    return """
    .site-footer { border-top: 1px solid rgba(255,210,31,.14); background: #080808; }
    .site-footer-inner { width: min(1240px, calc(100% - 32px)); margin: 0 auto; display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 18px 0; color: var(--muted); font-size: 12px; }
    .site-footer a { color: var(--accent); font-weight: 850; text-decoration: none; }
    .site-footer a:hover { color: var(--accent-strong); }
    @media (max-width: 820px) { .site-footer-inner { flex-direction: column; align-items: flex-start; } }
"""


def inject_roleforge_footer(html_text: str) -> str:
    if ".site-footer {" not in html_text:
        html_text = html_text.replace("</style>", f"{roleforge_footer_css()}  </style>", 1)
    if "data-roleforge-footer" not in html_text and "</main>" in html_text:
        html_text = html_text.replace("</main>", f"</main>{roleforge_footer_html()}", 1)
    return html_text


def roleforge_brand_css() -> str:
    return (
        "    .brand-name { color: var(--ink); font-weight: 950; white-space: nowrap; }\n"
        "    .brand-mark { position: relative; width: 42px; height: 42px; flex: 0 0 42px; border-radius: 10px; display: grid; place-items: center; overflow: hidden; background: linear-gradient(135deg, var(--accent), var(--accent-strong)); color: #050505; font-size: 13px; font-weight: 950; box-shadow: 0 12px 28px rgba(255,210,31,.16); }\n"
        "    .brand-mark::before { content: \"\"; position: absolute; inset: 7px; border: 2px solid rgba(5,5,5,.24); border-radius: 7px; transform: rotate(-8deg); }\n"
        "    .brand-mark::after { content: \"\"; position: absolute; right: 7px; top: 7px; width: 7px; height: 7px; border-radius: 50%; background: rgba(5,5,5,.66); box-shadow: -14px 16px 0 rgba(5,5,5,.36); }\n"
        "    @media (max-width: 820px) { .nav { min-height: 62px; flex-wrap: wrap; gap: 10px; padding: 10px 0; } .nav-actions { width: 100%; flex-wrap: wrap; gap: 8px; padding-bottom: 2px; } .brand-name { font-size: 15px; } .button { min-height: 36px; padding: 8px 10px; } }"
    )


def inject_roleforge_branding(html_text: str) -> str:
    html_text = html_text.replace("<title>Job Agent - Upload Resume</title>", "<title>RoleForge - Upload Resume</title>")
    html_text = re.sub(
        r"<title>Latest Job Matches - ([^<]+)</title>",
        r"<title>RoleForge - Latest Job Matches - \1</title>",
        html_text,
        count=1,
    )
    html_text = re.sub(
        r'<span class="brand-mark">(?:JA|RF)</span><span(?: class="brand-name")?>[^<]*(?:Job Agent|RoleForge)[^<]*</span>',
        '<span class="brand-mark">RF</span><span class="brand-name">RoleForge</span>',
        html_text,
        count=1,
    )
    html_text = html_text.replace(
        "Job Agent will parse it, find matching jobs, and build your dashboard.",
        "RoleForge will parse it, find matching jobs, and build your dashboard.",
    )
    if ".brand-name {" not in html_text:
        old_brand_css = (
            ".brand-mark { width: 38px; height: 38px; border-radius: 8px; display: grid; place-items: center; "
            "background: var(--accent); color: #050505; font-weight: 950; }"
        )
        if old_brand_css in html_text:
            html_text = html_text.replace(old_brand_css, roleforge_brand_css().strip(), 1)
        else:
            html_text = html_text.replace("</style>", f"{roleforge_brand_css()}\n  </style>", 1)
    if "RoleForge mobile header" not in html_text and "flex-wrap: wrap; gap: 10px; padding: 10px 0;" not in html_text:
        html_text = html_text.replace(
            "</style>",
            """
    /* RoleForge mobile header */
    @media (max-width: 820px) {
      .nav { min-height: 62px; flex-wrap: wrap; gap: 10px; padding: 10px 0; }
      .nav-actions { width: 100%; flex-wrap: wrap; gap: 8px; padding-bottom: 2px; }
      .brand-name { font-size: 15px; }
      .button { min-height: 36px; padding: 8px 10px; }
    }
  </style>""",
            1,
        )
    return html_text


def clean_dashboard_text(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", str(value))).strip()


def dashboard_items(profile: dict, key: str, limit: int) -> list[str]:
    output = []
    seen = set()
    raw = profile.get(key, []) or []
    if isinstance(raw, str):
        raw = re.split(r"[\n,]", raw)
    for item in raw:
        value = clean_dashboard_text(item)
        marker = value.lower()
        if value and marker not in seen:
            seen.add(marker)
            output.append(value)
    return output[:limit]


def resume_insights_from_profile(profile: dict) -> dict:
    primary_skills = dashboard_items(profile, "primary_skills", 12)
    secondary_skills = dashboard_items(profile, "secondary_skills", 8)
    domains = dashboard_items(profile, "domains", 8)
    target_roles = dashboard_items(profile, "target_roles", 6)
    keywords = dashboard_items(profile, "keywords", 12)
    summary = clean_dashboard_text(profile.get("summary", ""))
    seniority = clean_dashboard_text(profile.get("seniority", ""))
    years = clean_dashboard_text(profile.get("total_years_experience", ""))
    position = clean_dashboard_text(profile.get("current_position", ""))
    company = clean_dashboard_text(profile.get("current_company", ""))

    if not summary:
        proof = ", ".join(primary_skills[:4] or keywords[:4])
        role = target_roles[0] if target_roles else position or "the target role"
        summary = (
            f"Profile aligns well with {role} opportunities, with strongest evidence around {proof}."
            if proof
            else "Upload a richer resume or enable OpenAI parsing to generate deeper profile insight."
        )

    overview_bits = []
    if years:
        overview_bits.append(f"{years} years of experience" if "year" not in years.lower() else years)
    if seniority:
        overview_bits.append(seniority)
    if position:
        overview_bits.append(position)
    if company:
        overview_bits.append(f"current company: {company}")

    return {
        "overview": " | ".join(overview_bits) or "Profile overview will improve after the resume parser extracts more structured details.",
        "summary": summary,
        "strong_skills": primary_skills or keywords[:10],
        "supporting_skills": secondary_skills,
        "domains": domains,
        "best_match_roles": target_roles[:5] or ([position] if position else []),
    }


def configured_profile_cache_path(config: dict) -> Path:
    return resolve_config_path(config.get("llm_resume_parser", {}).get("cache_path", "data/resume_profile.json"))


def load_resume_insights_for_dashboard() -> dict:
    latest_path = DATA_DIR / "latest_jobs.json"
    if latest_path.exists():
        try:
            payload = json.loads(latest_path.read_text(encoding="utf-8"))
            insights = payload.get("resume_insights")
            if isinstance(insights, dict):
                return insights
        except Exception:
            pass

    try:
        config = load_config()
        profile_path = configured_profile_cache_path(config)
        if profile_path.exists():
            payload = json.loads(profile_path.read_text(encoding="utf-8"))
            profile = payload.get("profile") if isinstance(payload, dict) else None
            if isinstance(profile, dict):
                return resume_insights_from_profile(profile)
    except Exception:
        pass
    return {}


def insight_chips(values: list[str], muted: bool = False, empty: str = "") -> str:
    items = []
    class_name = "insight-chip muted" if muted else "insight-chip"
    for value in values:
        items.append(f'<span class="{class_name}">{html_escape(str(value))}</span>')
    if not items and empty:
        return f'<span class="insight-empty">{html_escape(empty)}</span>'
    return "".join(items)


def render_resume_insights_html(insights: dict) -> str:
    strong = insight_chips(list(insights.get("strong_skills", []) or [])[:10], empty="No strong skills extracted yet.")
    supporting = insight_chips(list(insights.get("supporting_skills", []) or [])[:8], muted=True)
    domains = insight_chips(list(insights.get("domains", []) or [])[:8], muted=True)
    signals = supporting + domains or '<span class="insight-empty">No supporting signals extracted yet.</span>'
    roles = "".join(
        f"<li>{html_escape(str(role))}</li>" for role in list(insights.get("best_match_roles", []) or [])[:5]
    ) or "<li>Run resume parsing to infer recommended roles.</li>"
    return f"""
  <section class="resume-insights" id="resumeInsights">
    <div class="resume-insights-inner">
      <div class="insight-grid">
        <article class="insight-panel">
          <p class="eyebrow">Resume insights</p>
          <h2>Profile overview</h2>
          <p class="insight-overview">{html_escape(str(insights.get("overview") or ""))}</p>
          <p class="insight-copy">{html_escape(str(insights.get("summary") or ""))}</p>
        </article>
        <aside class="insight-panel insight-groups" aria-label="Resume skill and role insights">
          <div>
            <h3>Strong skill set</h3>
            <div class="insight-chip-row">{strong}</div>
          </div>
          <div>
            <h3>Supporting signals</h3>
            <div class="insight-chip-row">{signals}</div>
          </div>
          <div>
            <h3>Best match roles</h3>
            <ol class="role-list">{roles}</ol>
          </div>
        </aside>
      </div>
    </div>
  </section>"""


def inject_resume_insights(html_text: str) -> str:
    insights = load_resume_insights_for_dashboard()
    if not insights:
        return html_text
    if ".resume-insights {" not in html_text:
        css = """
    .resume-insights { border-bottom: 1px solid rgba(255,210,31,.14); background: #080808; }
    .resume-insights-inner { width: min(1240px, calc(100% - 32px)); margin: 0 auto; padding: 24px 0; }
    .insight-grid { display: grid; grid-template-columns: minmax(0, 1.15fr) minmax(260px, .85fr); gap: 16px; }
    .insight-panel { border: 1px solid var(--line); border-radius: 8px; background: var(--panel); padding: 16px; }
    .insight-panel h2, .insight-panel h3 { margin: 0; color: var(--accent); letter-spacing: 0; }
    .insight-panel h2 { font-size: 1.05rem; }
    .insight-panel h3 { font-size: .86rem; text-transform: uppercase; }
    .insight-overview { margin: 10px 0 0; color: rgba(255,255,255,.86); font-size: 1rem; }
    .insight-copy { margin: 10px 0 0; color: var(--muted); }
    .insight-groups { display: grid; gap: 14px; }
    .insight-chip-row { display: flex; flex-wrap: wrap; gap: 7px; margin-top: 9px; }
    .insight-chip { display: inline-flex; align-items: center; min-height: 28px; padding: 5px 9px; border: 1px solid rgba(255,210,31,.3); border-radius: 999px; color: var(--accent); background: rgba(255,210,31,.06); font-size: 12px; font-weight: 850; }
    .insight-chip.muted { color: rgba(255,255,255,.78); border-color: var(--line); background: rgba(255,255,255,.04); }
    .insight-empty { color: var(--muted); font-size: 12px; }
    .role-list { margin: 10px 0 0; padding-left: 18px; color: rgba(255,255,255,.86); }
    .role-list li { margin: 5px 0; }
"""
        if "    main {" in html_text:
            html_text = html_text.replace("    main {", f"{css}    main {{", 1)
        else:
            html_text = html_text.replace("</style>", f"{css}  </style>", 1)
    if "@media (max-width: 820px)" in html_text and ".insight-grid { grid-template-columns: 1fr; }" not in html_text:
        html_text = html_text.replace(
            "      .hero-stats { grid-template-columns: 1fr; }",
            "      .hero-stats { grid-template-columns: 1fr; }\n      .insight-grid { grid-template-columns: 1fr; }",
            1,
        )
    if 'id="resumeInsights"' not in html_text:
        section = render_resume_insights_html(insights)
        html_text = html_text.replace("  <main>", f"{section}\n  <main>", 1)
    return html_text


def load_candidate_profile_for_nav() -> dict:
    latest_path = DATA_DIR / "latest_jobs.json"
    if latest_path.exists():
        try:
            payload = json.loads(latest_path.read_text(encoding="utf-8"))
            profile = payload.get("candidate_profile")
            if isinstance(profile, dict):
                return {
                    "name": str(profile.get("name") or "Candidate"),
                    "experience": str(profile.get("experience") or "Experience unavailable"),
                    "position": str(profile.get("position") or "Current position unavailable"),
                    "company": str(profile.get("company") or "Current company unavailable"),
                }
        except Exception:
            pass
    return {
        "name": "Candidate",
        "experience": "Experience unavailable",
        "position": "Current position unavailable",
        "company": "Current company unavailable",
    }


def candidate_initials(name: str) -> str:
    parts = re.findall(r"[A-Za-z]+", name or "")
    if not parts or str(name or "").strip().lower() == "candidate":
        return "?"
    return "".join(part[0].upper() for part in parts[:2])


def inject_candidate_profile(html_text: str) -> str:
    profile = load_candidate_profile_for_nav()
    menu = (
        '<div class="profile-menu">'
        f'<button class="profile-button" id="profileButton" type="button" aria-label="Open profile" aria-expanded="false">{html_escape(candidate_initials(profile["name"]))}</button>'
        '<div class="profile-dropdown" id="profileDropdown" role="menu" aria-label="Candidate profile">'
        f'<div class="profile-dropdown-title">{html_escape(profile["name"])}</div>'
        f'<div class="profile-dropdown-row"><span>Experience</span><strong>{html_escape(profile["experience"])}</strong></div>'
        f'<div class="profile-dropdown-row"><span>Current role</span><strong>{html_escape(profile["position"])}</strong></div>'
        f'<div class="profile-dropdown-row"><span>Company</span><strong>{html_escape(profile["company"])}</strong></div>'
        '</div>'
        '</div>'
    )
    html_text = re.sub(
        r'\s*<div class="profile-chip" aria-label="Candidate profile">.*?</div>\s*(?=<div class="nav-actions">)',
        "\n      ",
        html_text,
        count=1,
        flags=re.S,
    )
    if "profileButton" not in html_text:
        loader_pattern = r'(<span class="run-loader" id="runLoader"[\s\S]*?<span class="run-loader-label"[^>]*>.*?</span>\s*</span>)'
        html_text, count = re.subn(loader_pattern, lambda match: f"{match.group(1)}\n        {menu}", html_text, count=1)
        if count == 0:
            html_text = html_text.replace("</div>\n    </nav>", f"        {menu}\n      </div>\n    </nav>", 1)
    if "Latest job matches for your next move." in html_text:
        name = profile.get("name") or "there"
        if name == "Candidate":
            name = "there"
        html_text = html_text.replace(
            "<h1>Latest job matches for your next move.</h1>",
            f'<h1><span class="hero-name">Hello {html_escape(name)}</span><span class="hero-title">Latest job matches for your next move.</span></h1>',
            1,
        )
        html_text = html_text.replace(
            f"Hello {html_escape(name)}, latest job matches for your next move.",
            f'<span class="hero-name">Hello {html_escape(name)}</span><span class="hero-title">Latest job matches for your next move.</span>',
            1,
        )
    if ".profile-button" not in html_text:
        profile_css = (
            "    .profile-menu { position: relative; display: inline-flex; align-items: center; }\n"
            "    .profile-button { width: 40px; height: 40px; display: grid; place-items: center; border: 1px solid rgba(255,210,31,.58); border-radius: 50%; background: var(--accent); color: #050505; font: inherit; font-size: 13px; font-weight: 950; cursor: pointer; box-shadow: 0 10px 26px rgba(255,210,31,.14); }\n"
            "    .profile-button:hover { background: var(--accent-strong); }\n"
            "    .profile-button:focus-visible { outline: 2px solid rgba(255,210,31,.62); outline-offset: 3px; }\n"
            "    .profile-dropdown { position: absolute; top: calc(100% + 10px); right: 0; width: min(320px, calc(100vw - 32px)); display: none; padding: 12px; border: 1px solid rgba(255,210,31,.26); border-radius: 8px; background: #101010; box-shadow: var(--shadow); }\n"
            "    .profile-dropdown.is-open { display: grid; gap: 10px; }\n"
            "    .profile-dropdown-title { color: var(--accent); font-size: 15px; font-weight: 950; line-height: 1.2; }\n"
            "    .profile-dropdown-row { display: grid; gap: 2px; padding-top: 8px; border-top: 1px solid var(--line); }\n"
            "    .profile-dropdown-row span { color: var(--muted); font-size: 10px; font-weight: 950; text-transform: uppercase; }\n"
            "    .profile-dropdown-row strong { display: block; color: var(--ink); font-size: 13px; line-height: 1.35; }"
        )
        old_brand_css = ".brand-mark { width: 38px; height: 38px; border-radius: 8px; display: grid; place-items: center; background: var(--accent); color: #050505; font-weight: 950; }"
        updated = html_text.replace(old_brand_css, f"{old_brand_css}\n{profile_css}", 1)
        if updated == html_text:
            updated = html_text.replace("</style>", f"{profile_css}\n  </style>", 1)
        html_text = updated
    if "serverInjectedProfileDropdown" not in html_text and "const profileButton = document.getElementById('profileButton')" not in html_text:
        script = """
  <script>
    const serverInjectedProfileButton = document.getElementById('profileButton');
    const serverInjectedProfileDropdown = document.getElementById('profileDropdown');
    if (serverInjectedProfileButton && serverInjectedProfileDropdown) {
      serverInjectedProfileButton.addEventListener('click', (event) => {
        event.stopPropagation();
        const shouldOpen = !serverInjectedProfileDropdown.classList.contains('is-open');
        serverInjectedProfileDropdown.classList.toggle('is-open', shouldOpen);
        serverInjectedProfileButton.setAttribute('aria-expanded', shouldOpen ? 'true' : 'false');
      });
      document.addEventListener('click', (event) => {
        if (!serverInjectedProfileDropdown.contains(event.target) && event.target !== serverInjectedProfileButton) {
          serverInjectedProfileDropdown.classList.remove('is-open');
          serverInjectedProfileButton.setAttribute('aria-expanded', 'false');
        }
      });
      document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
          serverInjectedProfileDropdown.classList.remove('is-open');
          serverInjectedProfileButton.setAttribute('aria-expanded', 'false');
        }
      });
    }
  </script>
"""
        html_text = html_text.replace("</body>", f"{script}</body>")
    if ".hero-name" not in html_text:
        html_text = html_text.replace(
            "h1::first-line { color: var(--accent); }",
            ".hero-name { display: block; color: var(--accent); }\n    .hero-title { display: block; }",
        )
        html_text = html_text.replace(
            "h1 { max-width: 900px; margin: 0; font-size: clamp(2.5rem, 7vw, 5.8rem); line-height: .96; letter-spacing: 0; }",
            "h1 { max-width: 900px; margin: 0; font-size: clamp(2.5rem, 7vw, 5.8rem); line-height: .96; letter-spacing: 0; }\n    .hero-name { display: block; color: var(--accent); }\n    .hero-title { display: block; }",
        )
    return html_text


def html_escape(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#x27;")
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Serve the job dashboard and on-demand resumes.")
    parser.add_argument("--host", default=None)
    parser.add_argument("--port", type=int, default=None)
    args = parser.parse_args()
    config = load_config()
    dashboard = config.get("dashboard", {})
    host = args.host or dashboard.get("host", "127.0.0.1")
    port = args.port if args.port is not None else int(dashboard.get("port", 8765))
    try:
        server = ReusableThreadingHTTPServer((host, port), DashboardHandler)
    except OSError as exc:
        if exc.errno == errno.EADDRINUSE:
            print(
                f"Dashboard port {host}:{port} is already in use. "
                f"The dashboard is probably already running at http://{host}:{port}. "
                "Use --port 0 for a random free port or --port <number> for another port.",
                file=sys.stderr,
                flush=True,
            )
            return 98
        raise
    actual_host, actual_port = server.server_address[:2]
    print(f"Job dashboard server running at http://{actual_host}:{actual_port}", flush=True)
    server.serve_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
