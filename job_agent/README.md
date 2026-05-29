# Job Match Agent

This local agent reads your resume, searches public web job feeds, and appends matching jobs posted on the run date to an Excel workbook. It stores seen job IDs in SQLite so the same job is not recorded twice.

## Setup

1. Copy the example config:

   ```sh
   cp config.example.json config.json
   cp .env.example .env
   ```

2. Edit root-level `config.json` and `.env`:

- Set `resume_path` to your resume file.
- Set `expected_salary.amount`, `currency`, and `period`.
- Salary values are normalized to INR in the workbook. USD salaries are converted using a live USD/INR rate when available, with `salary_conversion.usd_to_inr_fallback` as backup.
- Adjust roles, locations, required keywords, and excluded keywords.
- Optional: set `sources.adzuna` to `true` and provide `adzuna.app_id` / `adzuna.app_key`, or export `ADZUNA_APP_ID` and `ADZUNA_APP_KEY`.
- `llm_resume_parser.enabled` is `true` by default in the shared config. Put `OPENAI_API_KEY=...` in root-level `.env` to use OpenAI for cleaner resume skills, roles, domains, and keywords. If the key is missing or still a placeholder, the app falls back to local parsing and may report `local:low_quality`.
- For automatic LaunchAgent runs, keep `OPENAI_API_KEY=...` in root-level `.env`. This file is ignored by git.

3. Run once manually:

   ```sh
   zsh job_agent/scripts/run_agent_once.sh
   ```

4. Check what the resume parser sees:

   ```sh
   python3 job_agent/job_agent.py --config job_agent/config.json --print-resume-profile
   ```

5. Open the generated workbook:

   ```sh
   open job_agent/data/job_matches.xlsx
   ```

6. Open the latest jobs dashboard:

   ```sh
   zsh job_agent/scripts/start_server.sh
   ```

   Then open `http://127.0.0.1:8765`. The dashboard has search, source, and location filters, a `Run fresh report` button, apply links, and a `Download ATS friendly resume` link for each row. The resume is generated only when you click the download link.

## One-click start

Double-click this file in Finder:

```sh
job_agent/Start Job Agent.command
```

On first run it creates `job_agent/.venv`, installs Python dependencies from `requirements.txt`, creates root-level `config.json` and `.env` from examples if they do not exist, runs a fresh job search, starts the dashboard server, and opens it in the browser. If `8765` is already used by another copy, it automatically starts this copy on a free local port. The implementation lives in `job_agent/scripts/bootstrap_and_start.command`.

## Share with someone else

Create a clean distributable zip:

```sh
zsh job_agent/scripts/package_for_distribution.sh
```

Send either:

- `dist/job-agent-util/`
- `dist/job-agent-util.zip`

Both exclude your private `config.json`, `.env`, generated resumes, database, logs, and local virtualenv. The recipient should unzip it, double-click `job_agent/Start Job Agent.command` once to bootstrap, then edit root-level `config.json` and `.env`. Dependencies install automatically into `job_agent/.venv`.

Do not send your real root `.env` unless you intentionally want to share your OpenAI API key. The package includes `.env.example` only.

## Folder structure

- `config.json` - user-editable resume path, salary expectation, search settings, and source settings.
- `.env` - user-editable secrets such as `OPENAI_API_KEY`.
- `job_agent/job_agent.py` - searches jobs, scores matches, writes Excel/JSON/HTML outputs.
- `job_agent/job_agent_server.py` - serves the dashboard, reruns the agent on demand, and generates tailored DOCX resumes.
- `job_agent/build_ats_resume.py` - builds fact-preserving, job-tailored resumes.
- `job_agent/requirements.txt` - Python dependencies installed by the one-click launcher.
- `job_agent/scripts/` - bootstrap, packaging, one-click, and command-line launchers.
- `job_agent/data/` - generated workbook, dashboard, cache, database, and resumes.
- `job_agent/logs/` - dashboard and LaunchAgent logs.

## Run on laptop open/login

After `config.json` exists, install the macOS LaunchAgent:

```sh
zsh job_agent/install_launch_agent.sh
```

The LaunchAgent starts a small background daemon and a dashboard server. The search runs at login, detects laptop wake by checking for a sleep gap, and also runs every 6 hours while you are logged in so the workbook and dashboard stay fresh. The dashboard server runs at `http://127.0.0.1:8765`.

You can also refresh manually from the dashboard with `Run fresh report`. The server stays running, the agent reruns in the background request, and the page reloads with the new `latest_jobs.html` and `latest_jobs.json` output.

## Output columns

- Run Date
- Posted Date
- Source
- Title
- Company
- Location
- Salary Listed
- Meets Expected Salary
- Match Score
- Matched Resume Keywords
- Apply Link
- Job UID

The agent also writes `data/latest_jobs.html` and `data/latest_jobs.json` after every run. The HTML dashboard links to the local server for on-demand resume generation.

## Notes

- Sources: Jobicy, RemoteJobs.org, RemoteOK, Remotive, Arbeitnow, Career Nest, optional WorkAnywhere.pro, and optional Adzuna India.
- India-focused sources now include dedicated parsers for Jobsora India and Shine, which were reachable in live tests. Indeed India, Naukri, Instahyre, foundit, TimesJobs, Cutshort, Hirist, and Apna remain configurable but are disabled by default because they commonly block automated access or do not expose usable listing data.
- Salary is included when the source publishes it. Many job boards omit salary, so those rows show `Unknown`.
- The default freshness mode is `today`, so only jobs whose source posted date equals the agent run date are added.
- The PDF reader is best-effort without extra packages. A `.txt`, `.md`, or `.docx` resume gives cleaner keyword matching.
- The local parser now prefers macOS Spotlight text extraction and falls back to `strings` for PDFs. The LLM parser is more accurate for PDF resumes, especially when text extraction is noisy.
