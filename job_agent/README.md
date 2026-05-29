# Job Search Agent

This folder contains the runnable local app. For the full project guide, see the root `README.md`.

## Quick Start

From the repo root:

```sh
./job_agent/Start\ Job\ Agent.command
```

The launcher creates `job_agent/.venv`, installs dependencies, starts the local dashboard, and opens your browser. It does not fetch jobs on startup.

On first launch, upload your resume in the browser. The app saves it under `data/uploaded_resumes/`, updates root `config.json`, runs a fresh job search, and opens the dashboard.

## Main Files

- `job_agent.py`: searches jobs, scores matches, writes JSON/HTML/XLSX outputs.
- `job_agent_server.py`: serves the dashboard, settings, resume upload, rerun, and tailored resume download endpoints.
- `build_ats_resume.py`: builds job-tailored DOCX resumes.
- `scripts/bootstrap_and_start.command`: one-click startup flow.
- `scripts/package_for_distribution.sh`: builds the shareable zip.

## Local Files Not To Commit

The app writes private or generated files during use. These are ignored:

- root `.env`
- root `config.json`
- `data/`
- `job_agent/data/`
- `job_agent/logs/`
- `job_agent/.venv/`
- generated resumes and uploaded resumes
