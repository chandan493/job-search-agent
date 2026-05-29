# Job Search Agent

Job Search Agent is a local, private job-search dashboard. It reads your resume, finds matching jobs from public job feeds, estimates missing salaries, and generates job-specific tailored resumes on demand.

The app is wizard-first: on first launch it opens in your browser and asks you to upload a resume. After upload, it parses the resume, runs the job search, and shows your dashboard.

## What It Does

- Upload a resume from the browser: PDF, DOCX, or TXT.
- Parse the resume locally or with OpenAI when `OPENAI_API_KEY` is configured.
- Search enabled job sources and score jobs against your resume.
- Show a local dashboard with filters, profile summary, config, and change-resume controls.
- Estimate missing salary ranges with a visible disclaimer.
- Generate a tailored DOCX resume per job without inventing facts.
- Download tailored resumes with a blocking progress loader.
- Store generated data locally only.

## Clone And Run Locally

1. Clone the repo:

   ```sh
   git clone <your-repo-url>
   cd job-search-agent
   ```

2. Create your local env file:

   ```sh
   cp .env.example .env
   ```

3. Add your OpenAI key to `.env` if you want higher-quality resume parsing and tailored resume generation:

   ```sh
   OPENAI_API_KEY=your_key_here
   ```

4. Start the app:

   ```sh
   ./job_agent/Start\ Job\ Agent.command
   ```

   On first run this creates `job_agent/.venv`, installs Python dependencies, creates `config.json` if needed, starts the local dashboard server, and opens the browser.

5. Upload your resume in the first-launch wizard.

   After upload, the app updates `config.json`, runs a fresh job search, and reloads the dashboard with matching jobs.

## Daily Use

Start the app with:

```sh
./job_agent/Start\ Job\ Agent.command
```

The launcher does not fetch jobs immediately. It only starts the dashboard. From the dashboard you can:

- Click `Run fresh report` to fetch current jobs.
- Click `Change resume` to upload another resume, update config, rerun the search, and refresh the dashboard.
- Click `Config` to change salary, locations, roles, freshness, and enabled job sites.
- Click `Download` on a job row to generate a tailored resume for that job.

## Configuration

The app creates local `config.json` from `config.example.json`. Do not commit `config.json`.

Important settings:

- `resume_path`: set automatically when you upload a resume from the UI.
- `expected_salary`: used for salary checks.
- `search.roles`: fallback roles to search.
- `search.locations`: preferred locations.
- `search.freshness`: `today` or `all`.
- `sources`: enable or disable job sites.
- `llm_resume_parser.enabled`: use OpenAI parsing when an API key exists.

Secrets live in `.env`, especially:

```sh
OPENAI_API_KEY=...
```

## Dashboard Features

- Profile menu shows name, experience, current role, and company from the uploaded resume.
- Salary column shows posted salary when available.
- Missing salaries show an estimated range with an `i` marker. Hover, focus, or click it for the disclaimer.
- Resume downloads are generated on demand for the selected job.
- Tailored resumes preserve only supported facts from the resume and job description.

## Generated Files

The app writes runtime files under `data/`, including:

- `latest_jobs.html`
- `latest_jobs.json`
- `job_matches.xlsx`
- `seen_jobs.sqlite3`
- `uploaded_resumes/`
- `generated_resumes/`
- resume parsing caches

These are intentionally ignored by git.

## Package For Someone Else

Create a clean distributable zip:

```sh
zsh job_agent/scripts/package_for_distribution.sh
```

Send:

```text
dist/job-agent-util.zip
```

The package excludes your `.env`, `config.json`, uploaded resumes, generated resumes, data, logs, and virtualenv. The recipient can unzip it, add their own `.env`, run `Start Job Agent.command`, and upload their resume in the browser.

## Optional Background LaunchAgent

After local setup, you can install the macOS LaunchAgent:

```sh
zsh job_agent/install_launch_agent.sh
```

This starts the dashboard and background daemon at login. Manual dashboard controls still work.

## Git Safety

Private and generated files should not be committed:

- `.env`
- `config.json`
- `data/`
- `job_agent/data/`
- `job_agent/logs/`
- `job_agent/.venv/`
- `dist/`

Only commit example files such as `.env.example` and `config.example.json`.
