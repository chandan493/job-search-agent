#!/usr/bin/env python3
"""Find fresh matching jobs from public web job feeds and append them to Excel."""

from __future__ import annotations

import argparse
import base64
import datetime as dt
import hashlib
import html
import json
import os
import re
import sqlite3
import subprocess
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from xml.sax.saxutils import escape, quoteattr


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG = BASE_DIR / "config.json"
USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X) JobMatchAgent/1.0"
SKILL_STOPWORDS = {
    "and", "the", "with", "for", "from", "this", "that", "your", "you", "are",
    "was", "were", "will", "have", "has", "had", "not", "but", "all", "can",
    "our", "their", "resume", "experience", "project", "projects", "work",
    "working", "using", "used", "team", "teams", "role", "roles", "based",
    "skills", "skill", "professional", "summary", "education", "degree",
    "type", "fontdescriptor", "flatedecode", "subtype", "ascent", "avgwidth",
    "basefont", "capheight", "descent", "firstchar", "flags", "fontbbox",
    "fontfile", "fontfile2", "fontname", "italicangle", "lastchar", "length1",
    "maxwidth", "stemv", "truetype", "widths", "xheight", "encoding",
    "macromanencoding", "mediabox", "contents", "page", "parent", "resources",
    "boldmt", "colorspace", "procset", "rotate", "xref", "trailer",
}
PDF_ARTIFACT_WORDS = {
    "fontdescriptor", "flatedecode", "basefont", "fontbbox", "fontfile2",
    "truetype", "mediabox", "xref", "endobj", "stream", "obj", "procset",
}
TECH_SKILLS = [
    "python", "java", "javascript", "typescript", "react", "react.js", "next.js",
    "node", "node.js", "express", "django", "flask", "fastapi", "spring",
    "spring boot", "go", "golang", "ruby", "rails", "php", "laravel", "c#",
    ".net", "kotlin", "swift", "scala", "rust", "html", "css", "tailwind",
    "redux", "angular", "vue", "graphql", "rest", "microservices", "api",
    "aws", "azure", "gcp", "docker", "kubernetes", "terraform", "ansible",
    "jenkins", "github actions", "gitlab ci", "ci/cd", "linux", "nginx",
    "sql", "postgres", "postgresql", "mysql", "mongodb", "redis", "elasticsearch",
    "dynamodb", "snowflake", "bigquery", "spark", "airflow", "kafka",
    "machine learning", "data science", "pandas", "numpy", "pytorch",
    "tensorflow", "llm", "rag", "langchain", "security", "oauth", "saml",
    "salesforce", "sap", "tableau", "power bi",
]
ROLE_PHRASES = [
    "software engineer", "senior software engineer", "staff software engineer",
    "full stack developer", "backend engineer", "frontend engineer",
    "platform engineer", "devops engineer", "site reliability engineer",
    "data engineer", "machine learning engineer", "engineering manager",
    "technical lead", "tech lead", "architect", "product manager",
]
DOMAIN_PHRASES = [
    "fintech", "banking", "payments", "ecommerce", "healthcare", "saas",
    "enterprise", "marketplace", "security", "cloud", "ai", "analytics",
    "logistics", "supply chain", "edtech", "telecom",
]
CURRENCY_TO_INR_FALLBACK = {
    "CAD": 70.0,
    "EUR": 103.0,
    "GBP": 120.0,
}


class RedirectHandler(urllib.request.HTTPRedirectHandler):
    def http_error_308(self, req, fp, code, msg, headers):
        return self.http_error_301(req, fp, code, msg, headers)


URL_OPENER = urllib.request.build_opener(RedirectHandler)


def today_for_timezone(tz_name: str) -> dt.date:
    try:
        from zoneinfo import ZoneInfo

        return dt.datetime.now(ZoneInfo(tz_name)).date()
    except Exception:
        return dt.datetime.now().date()


def now_for_timezone(tz_name: str) -> dt.datetime:
    try:
        from zoneinfo import ZoneInfo

        return dt.datetime.now(ZoneInfo(tz_name))
    except Exception:
        return dt.datetime.now()


def human_datetime(value: dt.datetime) -> str:
    return value.strftime("%A, %B %d, %Y at %I:%M %p").replace(" 0", " ")


def candidate_profile_summary(profile: Dict[str, Any]) -> Dict[str, str]:
    name = clean_text(profile.get("candidate_name", ""))
    years = clean_text(profile.get("total_years_experience", ""))
    company = clean_text(profile.get("current_company", ""))
    position = clean_text(profile.get("current_position", ""))
    if years and "year" not in years.lower():
        years = f"{years} years"
    return {
        "name": name or "Candidate",
        "experience": years or "Experience unavailable",
        "company": company or "Current company unavailable",
        "position": position or "Current position unavailable",
    }


def candidate_initials(name: str) -> str:
    parts = re.findall(r"[A-Za-z]+", name or "")
    if not parts or clean_text(name).lower() == "candidate":
        return "?"
    return "".join(part[0].upper() for part in parts[:2])


def resolve_path(path_value: str, config_dir: Path) -> Path:
    path = Path(os.path.expanduser(path_value))
    if path.is_absolute():
        return path
    return config_dir / path


def load_config(config_path: Path) -> Dict[str, Any]:
    if not config_path.exists():
        raise FileNotFoundError(
            f"Missing {config_path}. Copy config.example.json to config.json and edit it."
        )
    try:
        config = json.loads(config_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"Invalid JSON in {config_path}: line {exc.lineno}, column {exc.colno}. "
            "Check for a missing quote or comma, especially around resume_path."
        ) from exc
    return config


def load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if is_placeholder_secret(value):
            continue
        if key and key not in os.environ:
            os.environ[key] = value


def is_placeholder_secret(value: str) -> bool:
    lowered = value.strip().lower()
    return not lowered or "replace" in lowered or lowered in {"your-openai-api-key", "your-api-key"}


def fetch_json(url: str) -> Any:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with URL_OPENER.open(req, timeout=30) as response:
        charset = response.headers.get_content_charset() or "utf-8"
        return json.loads(response.read().decode(charset, errors="replace"))


def fetch_text(url: str) -> str:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        },
    )
    with URL_OPENER.open(req, timeout=30) as response:
        charset = response.headers.get_content_charset() or "utf-8"
        return response.read().decode(charset, errors="replace")


def post_json(url: str, payload: Dict[str, Any], api_key: str) -> Any:
    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "User-Agent": USER_AGENT,
        },
        method="POST",
    )
    try:
        with URL_OPENER.open(req, timeout=90) as response:
            charset = response.headers.get_content_charset() or "utf-8"
            return json.loads(response.read().decode(charset, errors="replace"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        message = body
        try:
            payload = json.loads(body)
            message = payload.get("error", {}).get("message") or body
        except json.JSONDecodeError:
            pass
        if exc.code == 401:
            raise RuntimeError(
                "OpenAI API authentication failed. Check that OPENAI_API_KEY is current, not revoked, "
                "and belongs to a project with API access."
            ) from exc
        raise RuntimeError(f"OpenAI API request failed with HTTP {exc.code}: {message}") from exc


def read_resume_text(resume_path: Path) -> str:
    if not resume_path.exists():
        return ""
    suffix = resume_path.suffix.lower()
    try:
        if suffix in {".txt", ".md", ".rtf"}:
            return resume_path.read_text(encoding="utf-8", errors="replace")
        if suffix == ".docx":
            return read_docx_text(resume_path)
        if suffix == ".pdf":
            return read_pdf_text(resume_path)
    except Exception as exc:
        print(f"Warning: could not read resume text from {resume_path}: {exc}", file=sys.stderr)
    return resume_path.read_text(encoding="utf-8", errors="ignore")


def read_docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("word/document.xml").decode("utf-8", errors="replace")
    xml = re.sub(r"<[^>]+>", " ", xml)
    return html.unescape(xml)


def command_output(args: Sequence[str]) -> str:
    result = subprocess.run(args, check=False, capture_output=True, text=True, timeout=30)
    if result.returncode != 0:
        return ""
    return result.stdout


def read_pdf_text(path: Path) -> str:
    mdls_text = command_output(["/usr/bin/mdls", "-raw", "-name", "kMDItemTextContent", str(path)])
    if mdls_text and mdls_text.strip() not in {"(null)", "null"}:
        return clean_resume_text(mdls_text)

    strings_text = command_output(["/usr/bin/strings", "-n", "4", str(path)])
    if strings_text:
        return clean_resume_text(strings_text)

    data = path.read_bytes()
    text = data.decode("latin-1", errors="ignore")
    text = re.sub(r"\\[nrt]", " ", text)
    return clean_resume_text(text)


def clean_resume_text(text: str) -> str:
    text = html.unescape(text)
    text = re.sub(r"https?://\S+", " ", text)
    text = re.sub(r"\S+@\S+", " ", text)
    text = re.sub(r"[^A-Za-z0-9+#./,()&:% -]+", " ", text)
    text = re.sub(r"\b(obj|endobj|stream|endstream|xref|trailer|font|image|metadata)\b", " ", text, flags=re.I)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def resume_text_quality(text: str) -> float:
    words = re.findall(r"[a-zA-Z]{3,}", text.lower())
    if not words:
        return 0.0
    artifact_count = sum(1 for word in words if word in PDF_ARTIFACT_WORDS or word in SKILL_STOPWORDS)
    return 1.0 - (artifact_count / max(1, len(words)))


def extract_resume_keywords(text: str, limit: int = 40) -> List[str]:
    lowered = text.lower()
    phrase_hits = []
    for phrase in TECH_SKILLS + ROLE_PHRASES + DOMAIN_PHRASES:
        if re.search(rf"(?<![a-z0-9]){re.escape(phrase.lower())}(?![a-z0-9])", lowered):
            phrase_hits.append(phrase)
    words = re.findall(r"[A-Za-z][A-Za-z0-9+#.]{2,}", lowered)
    counts: Dict[str, int] = {}
    for word in words:
        if word in SKILL_STOPWORDS or len(word) < 3:
            continue
        if word in {"pdf", "obj", "endobj", "stream", "font", "width", "height", "filter", "length"}:
            continue
        counts[word] = counts.get(word, 0) + 1
    ranked = [word for word, _ in sorted(counts.items(), key=lambda item: (-item[1], item[0]))]
    merged: List[str] = []
    for item in phrase_hits + ranked:
        if item not in merged:
            merged.append(item)
        if len(merged) >= limit:
            break
    return merged


def infer_resume_profile_local(text: str) -> Dict[str, Any]:
    years_match = re.search(r"(\d{1,2}\+?)\s+years?", text, re.I)
    total_years = years_match.group(1) if years_match else ""
    name = ""
    for line in re.split(r"[\r\n]+", text):
        candidate = clean_text(line)
        if 2 <= len(candidate.split()) <= 4 and re.fullmatch(r"[A-Za-z][A-Za-z .'-]+", candidate):
            name = candidate.title()
            break
    quality = resume_text_quality(text)
    if quality < 0.72:
        print(
            "Warning: local resume text extraction looks noisy. Enable llm_resume_parser or use a .txt/.docx resume export.",
            file=sys.stderr,
        )
        return {
            "source": "local:low_quality",
            "summary": "",
            "target_roles": [],
            "primary_skills": [],
            "secondary_skills": [],
            "domains": [],
            "seniority": "",
            "keywords": [],
            "current_company": "",
            "current_position": "",
            "total_years_experience": total_years,
            "candidate_name": name,
        }
    keywords = extract_resume_keywords(text)
    lowered = text.lower()
    target_roles = [role for role in ROLE_PHRASES if role in lowered]
    skills = [skill for skill in TECH_SKILLS if re.search(rf"(?<![a-z0-9]){re.escape(skill.lower())}(?![a-z0-9])", lowered)]
    domains = [domain for domain in DOMAIN_PHRASES if domain in lowered]
    seniority = ""
    if re.search(r"\b(staff|principal|architect|lead)\b", lowered):
        seniority = "lead/staff"
    elif re.search(r"\b(senior|sr\.?)\b", lowered):
        seniority = "senior"
    elif re.search(r"\b(junior|entry|associate)\b", lowered):
        seniority = "junior"
    return {
        "source": "local",
        "summary": "",
        "target_roles": target_roles[:8],
        "primary_skills": skills[:20],
        "secondary_skills": [kw for kw in keywords if kw not in skills][:20],
        "domains": domains[:10],
        "seniority": seniority,
        "keywords": keywords,
        "current_company": "",
        "current_position": target_roles[0] if target_roles else "",
        "total_years_experience": total_years,
        "candidate_name": name,
    }


def cache_key_for_resume(path: Path, config: Dict[str, Any]) -> str:
    stat = path.stat() if path.exists() else None
    digest = ""
    if path.exists():
        hasher = hashlib.sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                hasher.update(chunk)
        digest = hasher.hexdigest()
    llm = config.get("llm_resume_parser", {})
    raw = {
        "parser_version": 5,
        "path": str(path.resolve() if path.exists() else path),
        "size": stat.st_size if stat else 0,
        "mtime_ns": stat.st_mtime_ns if stat else 0,
        "sha256": digest,
        "model": llm.get("model", ""),
        "enabled": bool(llm.get("enabled", False)),
    }
    return hashlib.sha256(json.dumps(raw, sort_keys=True).encode("utf-8")).hexdigest()


def load_cached_profile(cache_path: Path, key: str) -> Optional[Dict[str, Any]]:
    if not cache_path.exists():
        return None
    try:
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
    except Exception:
        return None
    if payload.get("cache_key") == key and isinstance(payload.get("profile"), dict):
        return payload["profile"]
    return None


def save_cached_profile(cache_path: Path, key: str, profile: Dict[str, Any]) -> None:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    cache_path.write_text(
        json.dumps({"cache_key": key, "profile": profile}, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def response_output_text(response: Dict[str, Any]) -> str:
    chunks: List[str] = []
    for item in response.get("output", []):
        for content in item.get("content", []):
            if content.get("type") == "output_text":
                chunks.append(content.get("text", ""))
    if chunks:
        return "".join(chunks)
    return response.get("output_text", "")


def extract_resume_profile_with_llm(resume_path: Path, resume_text: str, config: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    llm = config.get("llm_resume_parser", {})
    if not llm.get("enabled", False):
        return None
    api_key = llm.get("api_key") or os.environ.get(llm.get("api_key_env", "OPENAI_API_KEY"))
    if not api_key:
        print("Warning: LLM resume parser is enabled but no API key is configured", file=sys.stderr)
        return None

    model = llm.get("model", "gpt-5.4-mini")
    if is_placeholder_secret(api_key):
        print("Warning: LLM resume parser is enabled but OPENAI_API_KEY is missing or still a placeholder", file=sys.stderr)
        return None
    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "summary": {"type": "string"},
            "target_roles": {"type": "array", "items": {"type": "string"}},
            "primary_skills": {"type": "array", "items": {"type": "string"}},
            "secondary_skills": {"type": "array", "items": {"type": "string"}},
            "domains": {"type": "array", "items": {"type": "string"}},
            "seniority": {"type": "string"},
            "keywords": {"type": "array", "items": {"type": "string"}},
            "current_company": {"type": "string"},
            "current_position": {"type": "string"},
            "total_years_experience": {"type": "string"},
            "candidate_name": {"type": "string"},
        },
        "required": [
            "summary",
            "target_roles",
            "primary_skills",
            "secondary_skills",
            "domains",
            "seniority",
            "keywords",
            "current_company",
            "current_position",
            "total_years_experience",
            "candidate_name",
        ],
    }
    prompt = (
        "Extract a concise job-search profile from this resume. "
        "Return only accurate terms clearly supported by the resume. "
        "Identify candidate_name from the resume header. "
        "Identify current_company and current_position from the most recent/current experience entry only. "
        "Set total_years_experience to the resume-supported total experience, such as '13+' or '8'. "
        "Prefer canonical skills and job titles over random repeated words. "
        "Do not include contact info, education institution names, PDF artifacts, or generic verbs. "
        "Keep target_roles under 8, primary_skills under 20, secondary_skills under 20, "
        "domains under 10, and keywords under 40."
    )
    content: List[Dict[str, Any]] = [{"type": "input_text", "text": prompt}]
    if resume_path.exists() and resume_path.suffix.lower() == ".pdf" and llm.get("send_pdf_file", True):
        file_data = base64.b64encode(resume_path.read_bytes()).decode("ascii")
        content.append(
            {
                "type": "input_file",
                "filename": resume_path.name,
                "file_data": f"data:application/pdf;base64,{file_data}",
            }
        )
    else:
        content.append({"type": "input_text", "text": resume_text[:50000]})

    payload = {
        "model": model,
        "input": [{"role": "user", "content": content}],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "resume_profile",
                "strict": True,
                "schema": schema,
            }
        },
        "max_output_tokens": 2000,
    }
    try:
        response = post_json("https://api.openai.com/v1/responses", payload, api_key)
        profile = json.loads(response_output_text(response))
        profile["source"] = f"llm:{model}"
        return profile
    except Exception as exc:
        print(f"Warning: LLM resume parser failed, using local parser: {exc}", file=sys.stderr)
        return None


def extract_resume_profile(resume_path: Path, config: Dict[str, Any], config_dir: Path) -> Dict[str, Any]:
    load_env_file(config_dir / ".env")
    resume_text = read_resume_text(resume_path)
    if not resume_text:
        print(f"Warning: resume text is empty. Check resume_path: {resume_path}", file=sys.stderr)
    llm_enabled = config.get("llm_resume_parser", {}).get("enabled", False)
    cache_path = resolve_path(
        config.get("llm_resume_parser", {}).get("cache_path", "data/resume_profile.json"),
        config_dir,
    )
    key = cache_key_for_resume(resume_path, config)
    if config.get("llm_resume_parser", {}).get("cache", True):
        cached = load_cached_profile(cache_path, key)
        if cached:
            if llm_enabled and str(cached.get("source", "")).startswith("local:low_quality"):
                print("Warning: ignoring cached low-quality local resume profile and retrying LLM parser.", file=sys.stderr)
            else:
                return cached
    profile = extract_resume_profile_with_llm(resume_path, resume_text, config)
    if not profile:
        profile = infer_resume_profile_local(resume_text)
    if config.get("llm_resume_parser", {}).get("cache", True):
        save_cached_profile(cache_path, key, profile)
    return profile


def keywords_from_profile(profile: Dict[str, Any]) -> List[str]:
    merged: List[str] = []
    for key in ("primary_skills", "secondary_skills", "domains", "keywords"):
        for item in profile.get(key, []) or []:
            value = clean_text(item).lower()
            if value and value not in merged:
                merged.append(value)
    return merged[:60]


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    value = re.sub(r"<[^>]+>", " ", str(value))
    value = html.unescape(value)
    return re.sub(r"\s+", " ", value).strip()


def parse_date(value: Any) -> Optional[dt.date]:
    if not value:
        return None
    text = str(value).strip()
    lowered = text.lower()
    today = today_for_timezone("Asia/Kolkata")
    if any(token in lowered for token in ("just now", "today", "minute ago", "minutes ago", "hour ago", "hours ago")):
        return today
    days_match = re.search(r"\b(\d+)\s+days?\s+ago\b", lowered)
    if days_match:
        return today - dt.timedelta(days=int(days_match.group(1)))
    if "yesterday" in lowered:
        return today - dt.timedelta(days=1)
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    for candidate in (text, text[:10]):
        try:
            return dt.datetime.fromisoformat(candidate).date()
        except ValueError:
            pass
        try:
            return dt.datetime.strptime(candidate, "%a, %d %b %Y %H:%M:%S %z").date()
        except ValueError:
            pass
    match = re.search(r"\d{4}-\d{2}-\d{2}", text)
    if match:
        return dt.date.fromisoformat(match.group(0))
    return None


def normalize_salary(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        return ", ".join(clean_text(item) for item in value if item)
    return clean_text(value)


def flatten_text_parts(parts: Sequence[Any]) -> str:
    flattened: List[str] = []
    for part in parts:
        if part is None:
            continue
        if isinstance(part, list):
            flattened.extend(clean_text(item) for item in part if item)
        elif isinstance(part, dict):
            flattened.extend(clean_text(item) for item in part.values() if item)
        else:
            flattened.append(clean_text(part))
    return " ".join(part for part in flattened if part)


def get_usd_to_inr_rate(config: Dict[str, Any]) -> float:
    conversion = config.get("salary_conversion", {})
    fallback = float(conversion.get("usd_to_inr_fallback", 83.5))
    if not conversion.get("fetch_live_rate", True):
        return fallback
    try:
        payload = fetch_json("https://open.er-api.com/v6/latest/USD")
        rate = payload.get("rates", {}).get("INR")
        if rate:
            return float(rate)
    except Exception as exc:
        print(f"Warning: could not fetch live USD/INR rate, using fallback {fallback}: {exc}", file=sys.stderr)
    return fallback


def salary_currency(text: str) -> str:
    lowered = text.lower()
    if any(token in lowered for token in ("₹", "rs", "inr", "lakh", "lac", "lpa", "crore", "cr/")):
        return "INR"
    if any(token in lowered for token in ("$", "usd", "us$", "dollar")):
        return "USD"
    if "cad" in lowered:
        return "CAD"
    if any(token in lowered for token in ("eur", "€")):
        return "EUR"
    if any(token in lowered for token in ("gbp", "£")):
        return "GBP"
    return ""


def salary_numbers(text: str) -> List[float]:
    numbers: List[float] = []
    for match in re.finditer(r"(\d+(?:,\d{3})*(?:\.\d+)?)\s*(k|lakh|lakhs|lac|lacs|lpa|crore|crores|cr)?", text, re.I):
        amount = float(match.group(1).replace(",", ""))
        multiplier = (match.group(2) or "").lower()
        if multiplier == "k":
            amount *= 1000
        elif multiplier in {"lakh", "lakhs", "lac", "lacs", "lpa"}:
            amount *= 100000
        elif multiplier in {"crore", "crores", "cr"}:
            amount *= 10000000
        if amount >= 1:
            numbers.append(amount)
    return numbers


def salary_values_in_inr(salary_text: str, usd_to_inr: float) -> List[float]:
    if not salary_text:
        return []
    numbers = salary_numbers(salary_text)
    currency = salary_currency(salary_text)
    if not numbers:
        return []
    lowered = salary_text.lower()
    if currency == "INR" and any(token in lowered for token in ("lakh", "lakhs", "lac", "lacs", "lpa")):
        numbers = [value * 100000 if value < 1000 else value for value in numbers]
    elif currency == "INR" and any(token in lowered for token in ("crore", "crores", " cr", "cr/")):
        numbers = [value * 10000000 if value < 1000 else value for value in numbers]
    elif currency == "USD" and "k" in lowered:
        numbers = [value * 1000 if value < 1000 else value for value in numbers]
    if currency == "USD":
        return [value * usd_to_inr for value in numbers]
    if currency in CURRENCY_TO_INR_FALLBACK:
        return [value * CURRENCY_TO_INR_FALLBACK[currency] for value in numbers]
    if currency == "INR":
        return numbers
    return []


def format_inr(value: float) -> str:
    if value >= 10000000:
        return f"INR {value / 10000000:.2f} Cr/yr"
    if value >= 100000:
        return f"INR {value / 100000:.2f} LPA"
    return f"INR {round(value):,}/yr"


def salary_to_inr_display(salary_text: str, usd_to_inr: float) -> str:
    salary_text = normalize_salary(salary_text)
    original_match = re.search(r"\(converted from (.*?)\)$", salary_text, re.I)
    if original_match:
        salary_text = original_match.group(1)
    values = salary_values_in_inr(salary_text, usd_to_inr)
    if not salary_text:
        return ""
    if not values:
        return salary_text
    if len(values) == 1:
        converted = format_inr(values[0])
    else:
        converted = f"{format_inr(min(values))} - {format_inr(max(values))}"
    if salary_currency(salary_text) == "USD":
        return f"{converted} (converted from {salary_text})"
    return converted


def salary_meets_expectation(salary_text: str, expected: Dict[str, Any], usd_to_inr: float) -> str:
    amount = expected.get("amount")
    if not amount or not salary_text:
        return "Unknown"
    nums = salary_values_in_inr(salary_text, usd_to_inr)
    if not nums:
        return "Unknown"
    return "Yes" if max(nums) >= int(amount) else "No"


def resume_years(profile: Dict[str, Any]) -> Optional[float]:
    value = clean_text(profile.get("total_years_experience", ""))
    match = re.search(r"\d+(?:\.\d+)?", value)
    return float(match.group(0)) if match else None


def seniority_salary_band(title: str, years: Optional[float]) -> Tuple[float, float, str]:
    lowered = title.lower()
    if any(token in lowered for token in ("director", "head of", "vp ", "vice president")):
        return 8000000, 18000000, "director/head"
    if any(token in lowered for token in ("principal", "staff", "architect")):
        return 5500000, 14000000, "staff/principal"
    if any(token in lowered for token in ("lead", "manager", "engineering manager", "tech lead")):
        return 4000000, 9500000, "lead/manager"
    if any(token in lowered for token in ("senior", "sr.", "sr ")):
        return 2500000, 6500000, "senior"
    if any(token in lowered for token in ("junior", "associate", "entry", "fresher")):
        return 500000, 1500000, "early career"
    if years is not None:
        if years >= 12:
            return 4500000, 10500000, "12+ years"
        if years >= 8:
            return 3000000, 7500000, "8+ years"
        if years >= 5:
            return 2000000, 5000000, "5+ years"
        if years >= 2:
            return 1000000, 2800000, "2+ years"
    return 900000, 2400000, "market baseline"


def role_salary_multiplier(text: str) -> Tuple[float, str]:
    lowered = text.lower()
    if any(token in lowered for token in ("machine learning", "ml engineer", " ai", "ai ", "genai", "llm", "data scientist")):
        return 1.25, "AI/ML premium"
    if any(token in lowered for token in ("product manager", "program manager", "strategy")):
        return 1.15, "product/strategy role"
    if any(token in lowered for token in ("devops", "sre", "platform", "cloud", "security")):
        return 1.12, "cloud/platform premium"
    if any(token in lowered for token in ("data engineer", "analytics engineer")):
        return 1.10, "data engineering role"
    if any(token in lowered for token in ("qa", "quality", "test engineer", "support")):
        return 0.78, "QA/support role"
    if any(token in lowered for token in ("frontend", "front end", "mobile", "android", "ios")):
        return 0.95, "frontend/mobile role"
    return 1.0, "software role"


def location_salary_multiplier(location: str) -> Tuple[float, str]:
    lowered = location.lower()
    if any(token in lowered for token in ("united states", " usa", " us ", "new york", "san francisco", "california")):
        return 1.8, "US/global market"
    if any(token in lowered for token in ("europe", "germany", "netherlands", "uk", "london")):
        return 1.4, "Europe/UK market"
    if any(token in lowered for token in ("singapore", "australia")):
        return 1.45, "APAC premium market"
    if any(token in lowered for token in ("uae", "dubai", "abu dhabi")):
        return 1.25, "Gulf market"
    if any(token in lowered for token in ("bangalore", "bengaluru", "hyderabad", "gurgaon", "gurugram", "mumbai", "pune", "delhi", "noida")):
        return 1.08, "major India tech hub"
    if "remote" in lowered:
        return 1.03, "remote market"
    return 1.0, "location baseline"


def round_salary(value: float) -> int:
    step = 50000 if value < 3000000 else 100000
    return int(round(value / step) * step)


def estimated_salary_for_job(job: Dict[str, Any], resume_profile: Dict[str, Any]) -> Dict[str, str]:
    title = clean_text(job.get("title", ""))
    location = clean_text(job.get("location", ""))
    text = " ".join([title, clean_text(job.get("description", "")), clean_text(job.get("matched_keywords", ""))])
    low, high, seniority_basis = seniority_salary_band(title, resume_years(resume_profile))
    role_factor, role_basis = role_salary_multiplier(text)
    location_factor, location_basis = location_salary_multiplier(location)
    low = round_salary(low * role_factor * location_factor)
    high = round_salary(high * role_factor * location_factor)
    if high <= low:
        high = low + 500000
    return {
        "display": f"{format_inr(low)} - {format_inr(high)}",
        "detail": (
            f"Estimated market range based on role seniority ({seniority_basis}), {role_basis}, and {location_basis}. "
            "This is not a posted salary or offer. Validate against public compensation benchmarks such as Levels.fyi, "
            "Glassdoor, AmbitionBox, company career pages, and recruiter conversations."
        ),
    }


def job_id(source: str, title: str, company: str, url: str) -> str:
    key = f"{source}|{title}|{company}|{url}".lower()
    return hashlib.sha256(key.encode("utf-8")).hexdigest()[:24]


def salary_from_schema(value: Any) -> str:
    if not value:
        return ""
    if isinstance(value, str):
        return clean_text(value)
    if isinstance(value, list):
        return ", ".join(salary_from_schema(item) for item in value if item)
    if not isinstance(value, dict):
        return clean_text(value)
    currency = clean_text(value.get("currency") or value.get("salaryCurrency"))
    numeric = value.get("value") or value.get("baseSalary")
    if isinstance(numeric, dict):
        minimum = numeric.get("minValue") or numeric.get("minimumValue")
        maximum = numeric.get("maxValue") or numeric.get("maximumValue")
        unit = clean_text(numeric.get("unitText") or numeric.get("unit") or value.get("unitText"))
        if minimum or maximum:
            return clean_text(f"{currency} {minimum or ''} - {maximum or ''} {unit}".strip())
        if numeric.get("value"):
            return clean_text(f"{currency} {numeric.get('value')} {unit}".strip())
    return clean_text(json.dumps(value, ensure_ascii=False))


def company_from_schema(value: Any) -> str:
    if isinstance(value, dict):
        return clean_text(value.get("name") or value.get("legalName"))
    if isinstance(value, list) and value:
        return company_from_schema(value[0])
    return clean_text(value)


def location_from_schema(value: Any) -> str:
    if isinstance(value, list):
        return ", ".join(location_from_schema(item) for item in value if item)
    if not isinstance(value, dict):
        return clean_text(value)
    address = value.get("address") or {}
    if isinstance(address, dict):
        parts = [
            address.get("addressLocality"),
            address.get("addressRegion"),
            address.get("addressCountry"),
        ]
        return clean_text(", ".join(str(part) for part in parts if part))
    return clean_text(value.get("name") or value.get("address"))


def schema_nodes(node: Any) -> Iterable[Dict[str, Any]]:
    if isinstance(node, list):
        for item in node:
            yield from schema_nodes(item)
    elif isinstance(node, dict):
        if "@graph" in node:
            yield from schema_nodes(node["@graph"])
        node_type = node.get("@type")
        types = node_type if isinstance(node_type, list) else [node_type]
        if "JobPosting" in types:
            yield node


def parse_json_ld_jobs(source: str, page_url: str, page_text: str) -> List[Dict[str, Any]]:
    jobs: List[Dict[str, Any]] = []
    scripts = re.findall(
        r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        page_text,
        re.I | re.S,
    )
    for script in scripts:
        try:
            payload = json.loads(html.unescape(script.strip()))
        except json.JSONDecodeError:
            continue
        for item in schema_nodes(payload):
            title = clean_text(item.get("title"))
            company = company_from_schema(item.get("hiringOrganization"))
            apply_link = clean_text(item.get("url") or item.get("sameAs") or page_url)
            if not title or not apply_link:
                continue
            jobs.append(
                {
                    "source": source,
                    "title": title,
                    "company": company,
                    "location": location_from_schema(item.get("jobLocation") or item.get("applicantLocationRequirements")),
                    "posted_date": parse_date(item.get("datePosted") or item.get("validThrough")),
                    "salary": salary_from_schema(item.get("baseSalary") or item.get("estimatedSalary")),
                    "apply_link": apply_link,
                    "description": clean_text(
                        flatten_text_parts(
                            [
                                title,
                                company,
                                item.get("employmentType"),
                                item.get("industry"),
                                item.get("skills"),
                                item.get("description"),
                            ]
                        )
                    ),
                }
            )
    return jobs


def rss_jobs(source: str, feed_url: str) -> List[Dict[str, Any]]:
    text = fetch_text(feed_url)
    root = ET.fromstring(text)
    jobs: List[Dict[str, Any]] = []
    for item in root.findall(".//item"):
        title = clean_text(item.findtext("title"))
        link = clean_text(item.findtext("link") or item.findtext("guid"))
        description = clean_text(item.findtext("description"))
        jobs.append(
            {
                "source": source,
                "title": title,
                "company": "",
                "location": "",
                "posted_date": parse_date(item.findtext("pubDate") or item.findtext("published")),
                "salary": "",
                "apply_link": link,
                "description": clean_text(flatten_text_parts([title, description])),
            }
        )
    return jobs


def remoteok_jobs() -> List[Dict[str, Any]]:
    payload = fetch_json("https://remoteok.com/api")
    jobs = []
    for item in payload:
        if not isinstance(item, dict) or "position" not in item:
            continue
        tags = item.get("tags") or []
        title = clean_text(item.get("position"))
        company = clean_text(item.get("company"))
        url = clean_text(item.get("url") or f"https://remoteok.com/remote-jobs/{item.get('id', '')}")
        salary = normalize_salary(item.get("salary") or item.get("salary_min") or item.get("salary_max"))
        jobs.append(
            {
                "source": "RemoteOK",
                "title": title,
                "company": company,
                "location": clean_text(item.get("location") or "Remote"),
                "posted_date": parse_date(item.get("date")),
                "salary": salary,
                "apply_link": url,
                "description": clean_text(" ".join([title, company, " ".join(map(str, tags)), item.get("description") or ""])),
            }
        )
    return jobs


def remotive_jobs() -> List[Dict[str, Any]]:
    payload = fetch_json("https://remotive.com/api/remote-jobs")
    jobs = []
    for item in payload.get("jobs", []):
        title = clean_text(item.get("title"))
        company = clean_text(item.get("company_name"))
        jobs.append(
            {
                "source": "Remotive",
                "title": title,
                "company": company,
                "location": clean_text(item.get("candidate_required_location") or "Remote"),
                "posted_date": parse_date(item.get("publication_date")),
                "salary": normalize_salary(item.get("salary")),
                "apply_link": clean_text(item.get("url")),
                "description": clean_text(
                    " ".join([title, company, item.get("category") or "", item.get("description") or ""])
                ),
            }
        )
    return jobs


def arbeitnow_jobs() -> List[Dict[str, Any]]:
    payload = fetch_json("https://www.arbeitnow.com/api/job-board-api")
    jobs = []
    for item in payload.get("data", []):
        title = clean_text(item.get("title"))
        company = clean_text(item.get("company_name"))
        tags = item.get("tags") or []
        posted = item.get("created_at")
        if isinstance(posted, int):
            posted_date = dt.datetime.fromtimestamp(posted, tz=dt.timezone.utc).date()
        else:
            posted_date = parse_date(posted)
        jobs.append(
            {
                "source": "Arbeitnow",
                "title": title,
                "company": company,
                "location": clean_text(item.get("location") or ("Remote" if item.get("remote") else "")),
                "posted_date": posted_date,
                "salary": "",
                "apply_link": clean_text(item.get("url")),
                "description": clean_text(" ".join([title, company, " ".join(map(str, tags)), item.get("description") or ""])),
            }
        )
    return jobs


def jobicy_jobs() -> List[Dict[str, Any]]:
    payload = fetch_json("https://jobicy.com/api/v2/remote-jobs?count=100")
    items = payload.get("jobs") or payload.get("data") or []
    jobs = []
    for item in items:
        title = clean_text(item.get("jobTitle") or item.get("title"))
        company = clean_text(item.get("companyName") or item.get("company"))
        salary_parts = []
        if item.get("salaryMin"):
            salary_parts.append(str(item.get("salaryMin")))
        if item.get("salaryMax"):
            salary_parts.append(str(item.get("salaryMax")))
        salary = " - ".join(salary_parts)
        if salary:
            salary = " ".join(
                part for part in [item.get("salaryCurrency"), salary, item.get("salaryPeriod")] if part
            )
        jobs.append(
            {
                "source": "Jobicy",
                "title": title,
                "company": company,
                "location": clean_text(item.get("jobGeo") or item.get("location") or "Remote"),
                "posted_date": parse_date(item.get("pubDate") or item.get("posted_at")),
                "salary": normalize_salary(salary),
                "apply_link": clean_text(item.get("url")),
                "description": clean_text(
                    flatten_text_parts(
                        [
                            title,
                            company,
                            item.get("jobIndustry") or "",
                            item.get("jobType") or "",
                            item.get("jobLevel") or "",
                            item.get("jobExcerpt") or "",
                            item.get("jobDescription") or "",
                        ]
                    )
                ),
            }
        )
    return jobs


def remotejobsorg_jobs() -> List[Dict[str, Any]]:
    payload = fetch_json("https://remotejobs.org/api/v1/jobs?limit=50")
    jobs = []
    for item in payload.get("data", []):
        title = clean_text(item.get("title"))
        company_obj = item.get("company") or {}
        category_obj = item.get("category") or {}
        salary = item.get("salary_text")
        if not salary and (item.get("salary_min") or item.get("salary_max")):
            salary = f"{item.get('salary_min') or ''} - {item.get('salary_max') or ''}".strip(" -")
        jobs.append(
            {
                "source": "RemoteJobs.org",
                "title": title,
                "company": clean_text(company_obj.get("name") if isinstance(company_obj, dict) else company_obj),
                "location": clean_text(item.get("location") or "Remote"),
                "posted_date": parse_date(item.get("posted_at")),
                "salary": normalize_salary(salary),
                "apply_link": clean_text(item.get("apply_url") or item.get("url")),
                "description": clean_text(
                    " ".join(
                        [
                            title,
                            clean_text(category_obj.get("name") if isinstance(category_obj, dict) else category_obj),
                            item.get("type") or "",
                            item.get("description") or "",
                        ]
                    )
                ),
            }
        )
    return jobs


def adzuna_jobs(config: Dict[str, Any]) -> List[Dict[str, Any]]:
    adzuna = config.get("adzuna", {})
    app_id = adzuna.get("app_id") or os.environ.get("ADZUNA_APP_ID")
    app_key = adzuna.get("app_key") or os.environ.get("ADZUNA_APP_KEY")
    if not app_id or not app_key:
        print("Warning: adzuna enabled but ADZUNA_APP_ID/ADZUNA_APP_KEY are not configured", file=sys.stderr)
        return []
    search = config.get("search", {})
    roles = search.get("roles") or [""]
    locations = search.get("locations") or ["India"]
    jobs: List[Dict[str, Any]] = []
    seen_urls = set()
    for role in roles:
        for location in locations:
            params = {
                "app_id": app_id,
                "app_key": app_key,
                "what": role,
                "where": location,
                "sort_by": "date",
                "max_days_old": "1",
                "results_per_page": "50",
                "content-type": "application/json",
            }
            url = "https://api.adzuna.com/v1/api/jobs/in/search/1?" + urllib.parse.urlencode(params)
            payload = fetch_json(url)
            for item in payload.get("results", []):
                apply_link = clean_text(item.get("redirect_url") or item.get("adref"))
                if not apply_link or apply_link in seen_urls:
                    continue
                seen_urls.add(apply_link)
                title = clean_text(item.get("title"))
                company_obj = item.get("company") or {}
                location_obj = item.get("location") or {}
                salary = ""
                if item.get("salary_min") or item.get("salary_max"):
                    salary = f"INR {item.get('salary_min') or ''} - {item.get('salary_max') or ''}".strip(" -")
                jobs.append(
                    {
                        "source": "Adzuna India",
                        "title": title,
                        "company": clean_text(company_obj.get("display_name") if isinstance(company_obj, dict) else company_obj),
                        "location": clean_text(location_obj.get("display_name") if isinstance(location_obj, dict) else location_obj),
                        "posted_date": parse_date(item.get("created")),
                        "salary": normalize_salary(salary),
                        "apply_link": apply_link,
                        "description": clean_text(flatten_text_parts([title, item.get("description")])),
                    }
                )
    return jobs


def career_nest_jobs() -> List[Dict[str, Any]]:
    payload = fetch_json("https://careernest.cloud/api/jobs?limit=100")
    items = payload.get("jobs") or payload.get("data") or payload if isinstance(payload, list) else []
    jobs = []
    for item in items:
        title = clean_text(item.get("title") or item.get("job_title"))
        salary_obj = item.get("salary")
        salary = salary_from_schema(salary_obj) if salary_obj else item.get("salary_text")
        jobs.append(
            {
                "source": "Career Nest",
                "title": title,
                "company": clean_text(item.get("company") or item.get("company_name")),
                "location": clean_text(item.get("location") or item.get("job_location") or "Remote"),
                "posted_date": parse_date(item.get("posted_at") or item.get("pubDate") or item.get("date")),
                "salary": normalize_salary(salary),
                "apply_link": clean_text(item.get("job_url") or item.get("url") or item.get("apply_url")),
                "description": clean_text(flatten_text_parts([title, item.get("description")])),
            }
        )
    return [job for job in jobs if job["title"] and job["apply_link"]]


def workanywhere_jobs() -> List[Dict[str, Any]]:
    return rss_jobs("WorkAnywhere.pro", "https://workanywhere.pro/rss.xml")


def slugify_query(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")


def html_attr(block: str, attr: str) -> str:
    match = re.search(rf'{attr}=["\'](.*?)["\']', block, re.I | re.S)
    return html.unescape(match.group(1)) if match else ""


def jobsora_jobs(config: Dict[str, Any]) -> List[Dict[str, Any]]:
    search = config.get("search", {})
    roles = search.get("roles") or [""]
    locations = search.get("locations") or ["India"]
    jobs: List[Dict[str, Any]] = []
    seen_urls = set()
    location_aliases = {"bangalore": "bengaluru", "remote": "india"}
    for role in roles:
        role_slug = slugify_query(role)
        for location in locations:
            loc_slug = location_aliases.get(location.lower(), slugify_query(location))
            url = f"https://in.jobsora.com/jobs-{role_slug}-{loc_slug}?sort=date"
            page = fetch_text(url)
            for card in re.findall(r'<article class="js-listing-item.*?</article>', page, re.S):
                title_match = re.search(r'<h2 class="c-job-item__title">\s*<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', card, re.S)
                if not title_match:
                    continue
                apply_link = html.unescape(title_match.group(1))
                title = clean_text(title_match.group(2))
                if apply_link in seen_urls:
                    continue
                seen_urls.add(apply_link)
                info_items = [
                    clean_text(re.sub(r"<svg.*?</svg>", " ", item, flags=re.S))
                    for item in re.findall(r'<div class="c-job-item__info-item">(.*?)</div>', card, re.S)
                ]
                company = info_items[0] if info_items else ""
                job_location = info_items[1] if len(info_items) > 1 else location
                desc_match = re.search(r'<p class="c-job-item__description[^"]*">(.*?)</p>', card, re.S)
                date_match = re.search(r'<div class="c-job-item__date">\s*(.*?)\s*</div>', card, re.S)
                jobs.append(
                    {
                        "source": "Jobsora India",
                        "title": title,
                        "company": company,
                        "location": clean_text(job_location),
                        "posted_date": parse_date(clean_text(date_match.group(1) if date_match else "")),
                        "salary": "",
                        "apply_link": apply_link,
                        "description": clean_text(flatten_text_parts([title, company, desc_match.group(1) if desc_match else ""])),
                    }
                )
    return jobs


def shine_jobs(config: Dict[str, Any]) -> List[Dict[str, Any]]:
    search = config.get("search", {})
    roles = search.get("roles") or [""]
    locations = search.get("locations") or ["India"]
    jobs: List[Dict[str, Any]] = []
    seen_urls = set()
    for role in roles:
        role_slug = slugify_query(role)
        for location in locations:
            loc_slug = slugify_query(location)
            url = f"https://www.shine.com/job-search/{role_slug}-jobs-in-{loc_slug}?sort=1"
            page = fetch_text(url)
            match = re.search(r'<script id="__NEXT_DATA__" type="application/json">(.*?)</script>', page, re.S)
            if not match:
                continue
            payload = json.loads(html.unescape(match.group(1)))
            results = (
                payload.get("props", {})
                .get("pageProps", {})
                .get("initialState", {})
                .get("jsrp", {})
                .get("searchresult", {})
                .get("data", {})
                .get("results", [])
            )
            for item in results:
                slug = clean_text(item.get("jSlug"))
                apply_link = clean_text(item.get("jRUrl") or (f"https://www.shine.com/jobs/{slug}" if slug else ""))
                if not apply_link or apply_link in seen_urls:
                    continue
                seen_urls.add(apply_link)
                title = clean_text(item.get("jJT"))
                company = clean_text(item.get("jCName"))
                job_location = ", ".join(clean_text(loc) for loc in item.get("jLoc", []) if loc) if isinstance(item.get("jLoc"), list) else clean_text(item.get("jLoc"))
                jobs.append(
                    {
                        "source": "Shine",
                        "title": title,
                        "company": company,
                        "location": job_location,
                        "posted_date": parse_date(item.get("jPDate")),
                        "salary": normalize_salary(item.get("jSal")),
                        "apply_link": apply_link,
                        "description": clean_text(flatten_text_parts([title, company, item.get("jKwd"), item.get("jJD")])),
                    }
                )
    return jobs


def page_search_jobs(config: Dict[str, Any]) -> List[Dict[str, Any]]:
    sources = {
        "indeed_india": (
            "Indeed India",
            "https://in.indeed.com/jobs?q={query}&l={location}&fromage=1&sort=date",
        ),
        "naukri": (
            "Naukri",
            "https://www.naukri.com/jobs-in-{location}?k={query}&l={location}&jobAge=1",
        ),
        "instahyre": (
            "Instahyre",
            "https://www.instahyre.com/search-jobs/?search={query}&location={location}",
        ),
        "foundit": (
            "foundit",
            "https://www.foundit.in/srp/results?query={query}&locations={location}&sort=1",
        ),
        "timesjobs": (
            "TimesJobs",
            "https://www.timesjobs.com/candidate/job-search.html?searchType=personalizedSearch&txtKeywords={query}&txtLocation={location}&sequence=1",
        ),
        "cutshort": (
            "Cutshort",
            "https://cutshort.io/jobs/{query}-jobs?locations={location}",
        ),
        "hirist": (
            "Hirist",
            "https://www.hirist.tech/search/{query}-jobs-in-{location}.html",
        ),
        "apna": (
            "Apna",
            "https://apna.co/jobs?text={query}&location={location}",
        ),
    }
    enabled_sources = config.get("sources", {})
    search = config.get("search", {})
    roles = search.get("roles") or [""]
    locations = search.get("locations") or ["India"]
    jobs: List[Dict[str, Any]] = []
    seen_urls = set()
    failures: Dict[str, str] = {}
    for key, (source_name, template) in sources.items():
        if not enabled_sources.get(key, False):
            continue
        for role in roles:
            for location in locations:
                query = urllib.parse.quote_plus(role)
                loc = urllib.parse.quote_plus(location)
                slug_query = urllib.parse.quote_plus(role.replace(" ", "-"))
                url = template.format(query=query, location=loc, slug_query=slug_query)
                try:
                    parsed = parse_json_ld_jobs(source_name, url, fetch_text(url))
                except Exception as exc:
                    failures.setdefault(key, str(exc))
                    continue
                for job in parsed:
                    if job["apply_link"] in seen_urls:
                        continue
                    seen_urls.add(job["apply_link"])
                    jobs.append(job)
    for key, message in failures.items():
        print(f"Warning: {key} page search failed or was blocked: {message}", file=sys.stderr)
    return jobs


def collect_jobs(config: Dict[str, Any]) -> List[Dict[str, Any]]:
    enabled_sources = config.get("sources", {})
    collectors = {
        "jobicy": jobicy_jobs,
        "remotejobsorg": remotejobsorg_jobs,
        "remoteok": remoteok_jobs,
        "remotive": remotive_jobs,
        "arbeitnow": arbeitnow_jobs,
        "adzuna": lambda: adzuna_jobs(config),
        "careernest": career_nest_jobs,
        "workanywhere": workanywhere_jobs,
        "jobsora": lambda: jobsora_jobs(config),
        "shine": lambda: shine_jobs(config),
    }
    all_jobs: List[Dict[str, Any]] = []
    for name, collector in collectors.items():
        if not enabled_sources.get(name, False):
            continue
        try:
            all_jobs.extend(collector())
        except Exception as exc:
            print(f"Warning: {name} search failed: {exc}", file=sys.stderr)
    all_jobs.extend(page_search_jobs(config))
    return all_jobs


def score_job(job: Dict[str, Any], roles: Sequence[str], keywords: Sequence[str]) -> Tuple[int, List[str]]:
    haystack = " ".join(
        [job.get("title", ""), job.get("company", ""), job.get("location", ""), job.get("description", "")]
    ).lower()
    matched: List[str] = []
    score = 0
    for role in roles:
        role_l = role.lower()
        if role_l and role_l in haystack:
            score += 3
            matched.append(role)
    for keyword in keywords:
        kw_l = keyword.lower()
        if kw_l and kw_l in haystack:
            score += 1
            matched.append(keyword)
    return score, matched[:20]


def filter_jobs(
    jobs: Iterable[Dict[str, Any]],
    config: Dict[str, Any],
    resume_profile: Dict[str, Any],
    run_date: dt.date,
    usd_to_inr: float,
) -> List[Dict[str, Any]]:
    search = config.get("search", {})
    configured_roles = search.get("roles") or []
    profile_roles = resume_profile.get("target_roles", []) or []
    roles = list(dict.fromkeys(configured_roles + profile_roles))
    resume_keywords = keywords_from_profile(resume_profile)
    required = [kw.lower() for kw in search.get("required_keywords", [])]
    excluded = [kw.lower() for kw in search.get("excluded_keywords", [])]
    locations = [loc.lower() for loc in search.get("locations", []) if loc]
    min_score = int(search.get("min_score", 2))
    strict_today = search.get("freshness", "today") == "today"
    expected = config.get("expected_salary", {})
    filtered = []
    for job in jobs:
        if strict_today and job.get("posted_date") != run_date:
            continue
        haystack = " ".join([job.get("title", ""), job.get("location", ""), job.get("description", "")]).lower()
        if excluded and any(word in haystack for word in excluded):
            continue
        if required and not all(word in haystack for word in required):
            continue
        if locations and not any(loc in haystack for loc in locations):
            continue
        score, matched = score_job(job, roles, resume_keywords)
        if score < min_score:
            continue
        row = dict(job)
        row["match_score"] = score
        row["matched_keywords"] = ", ".join(matched)
        row["salary_inr"] = salary_to_inr_display(row.get("salary", ""), usd_to_inr)
        row["salary_meets_expectation"] = salary_meets_expectation(row.get("salary", ""), expected, usd_to_inr)
        salary_estimate = estimated_salary_for_job(row, resume_profile)
        row["salary_estimate"] = salary_estimate["display"]
        row["salary_estimate_detail"] = salary_estimate["detail"]
        row["salary_is_estimated"] = "yes" if not row["salary_inr"] else ""
        row["job_uid"] = job_id(row["source"], row["title"], row["company"], row["apply_link"])
        filtered.append(row)
    return sorted(filtered, key=lambda item: (-item["match_score"], item["source"], item["title"]))


def init_db(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS seen_jobs (
          job_uid TEXT PRIMARY KEY,
          first_seen_at TEXT NOT NULL
        )
        """
    )
    conn.commit()
    return conn


def unseen_jobs(conn: sqlite3.Connection, jobs: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    output = []
    for job in jobs:
        exists = conn.execute("SELECT 1 FROM seen_jobs WHERE job_uid = ?", (job["job_uid"],)).fetchone()
        if not exists:
            output.append(job)
    return output


def mark_seen(conn: sqlite3.Connection, jobs: Sequence[Dict[str, Any]]) -> None:
    now = dt.datetime.now(dt.timezone.utc).isoformat()
    conn.executemany(
        "INSERT OR IGNORE INTO seen_jobs (job_uid, first_seen_at) VALUES (?, ?)",
        [(job["job_uid"], now) for job in jobs],
    )
    conn.commit()


HEADERS = [
    "Run Date",
    "Posted Date",
    "Source",
    "Title",
    "Company",
    "Location",
    "Salary Listed (INR)",
    "Meets Expected Salary",
    "Match Score",
    "Matched Resume Keywords",
    "Apply Link",
    "Job UID",
]


def row_for_job(job: Dict[str, Any], run_date: dt.date) -> List[str]:
    posted = job.get("posted_date")
    return [
        run_date.isoformat(),
        posted.isoformat() if isinstance(posted, dt.date) else "",
        job.get("source", ""),
        job.get("title", ""),
        job.get("company", ""),
        job.get("location", ""),
        job.get("salary_inr") or job.get("salary", ""),
        job.get("salary_meets_expectation", ""),
        str(job.get("match_score", "")),
        job.get("matched_keywords", ""),
        job.get("apply_link", ""),
        job.get("job_uid", ""),
    ]


def job_to_dashboard_record(job: Dict[str, Any], run_date: dt.date) -> Dict[str, Any]:
    posted = job.get("posted_date")
    return {
        "run_date": run_date.isoformat(),
        "posted_date": posted.isoformat() if isinstance(posted, dt.date) else "",
        "source": job.get("source", ""),
        "title": job.get("title", ""),
        "company": job.get("company", ""),
        "location": job.get("location", ""),
        "salary_inr": job.get("salary_inr") or job.get("salary", ""),
        "salary_estimate": job.get("salary_estimate", ""),
        "salary_estimate_detail": job.get("salary_estimate_detail", ""),
        "salary_is_estimated": job.get("salary_is_estimated", ""),
        "salary_meets_expectation": job.get("salary_meets_expectation", ""),
        "match_score": job.get("match_score", ""),
        "matched_keywords": job.get("matched_keywords", ""),
        "apply_link": job.get("apply_link", ""),
        "job_uid": job.get("job_uid", ""),
        "description": job.get("description", ""),
    }


def write_latest_jobs_outputs(
    jobs: Sequence[Dict[str, Any]],
    run_date: dt.date,
    config: Dict[str, Any],
    config_dir: Path,
    resume_profile: Dict[str, Any],
) -> Tuple[Path, Path]:
    dashboard = config.get("dashboard", {})
    port = int(dashboard.get("port", 8765))
    run_at = now_for_timezone(config.get("timezone", "Asia/Kolkata"))
    run_timestamp = human_datetime(run_at)
    data_dir = resolve_path("data", config_dir)
    data_dir.mkdir(parents=True, exist_ok=True)
    json_path = data_dir / "latest_jobs.json"
    html_path = data_dir / "latest_jobs.html"
    records = [job_to_dashboard_record(job, run_date) for job in jobs]
    candidate_profile = candidate_profile_summary(resume_profile)
    json_path.write_text(
        json.dumps(
            {
                "run_date": run_date.isoformat(),
                "run_timestamp": run_timestamp,
                "candidate_profile": candidate_profile,
                "server_url": f"http://127.0.0.1:{port}",
                "jobs": records,
            },
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    sources = sorted({job["source"] for job in records if job["source"]})
    locations = sorted({job["location"] for job in records if job["location"]})
    source_options = '<option value="">All sources</option>' + ''.join(
        f'<option value="{html.escape(source)}">{html.escape(source)}</option>' for source in sources
    )
    location_options = '<option value="">All locations</option>' + ''.join(
        f'<option value="{html.escape(location)}">{html.escape(location)}</option>' for location in locations
    )
    top_score = max([int(job["match_score"]) for job in records if str(job["match_score"]).isdigit()] or [0])
    resume_count = len(records)
    india_count = sum(1 for job in records if any(token in job["location"].lower() for token in ("india", "bangalore", "bengaluru", "hyderabad")))
    profile_name = html.escape(candidate_profile["name"])
    profile_experience = html.escape(candidate_profile["experience"])
    profile_position = html.escape(candidate_profile["position"])
    profile_company = html.escape(candidate_profile["company"])
    profile_initials = html.escape(candidate_initials(candidate_profile["name"]))
    hero_name = profile_name if candidate_profile["name"] != "Candidate" else "there"
    html_rows = []
    for job in records:
        uid = urllib.parse.quote(job["job_uid"])
        resume_url = f"/resume/{uid}"
        apply = html.escape(job["apply_link"])
        resume_label = html.escape(f"Download tailored resume for {job['title']} at {job['company']}")
        if job.get("salary_inr"):
            salary_html = html.escape(job["salary_inr"])
        elif job.get("salary_estimate"):
            salary_html = (
                f'<span class="salary-estimate">Est. {html.escape(job["salary_estimate"])} '
                '<span class="salary-info" tabindex="0" aria-label="Salary estimate details">i'
                f'<span class="salary-tooltip">{html.escape(job["salary_estimate_detail"])}</span>'
                '</span></span>'
            )
        else:
            salary_html = "Not listed"
        html_rows.append(
            f"<tr data-source=\"{html.escape(job['source'])}\" data-location=\"{html.escape(job['location'])}\" "
            f"data-search=\"{html.escape(' '.join([job['title'], job['company'], job['location'], job['source'], job['matched_keywords']]).lower())}\">"
            f"<td><span class=\"score\">{html.escape(str(job['match_score']))}</span></td>"
            f"<td><span class=\"date\">{html.escape(job['posted_date'])}</span><span class=\"source-chip\">{html.escape(job['source'])}</span></td>"
            f"<td><strong>{html.escape(job['title'])}</strong><br><span>{html.escape(job['company'])}</span></td>"
            f"<td>{html.escape(job['location'])}</td>"
            f"<td>{salary_html}</td>"
            f"<td><span class=\"keywords\">{html.escape(job['matched_keywords'])}</span></td>"
            f"<td><a class=\"button secondary\" href=\"{apply}\" target=\"_blank\" rel=\"noopener\">Apply</a></td>"
            f"<td><a class=\"button primary download-button\" href=\"{resume_url}\" aria-label=\"{resume_label}\">Download</a></td>"
            "</tr>"
        )
    html_doc = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Latest Job Matches - {html.escape(run_date.isoformat())}</title>
  <style>
    :root {{
      --ink: #ffffff;
      --muted: #b8b8b8;
      --paper: #050505;
      --panel: #111111;
      --panel-soft: #171717;
      --line: #2c2c2c;
      --accent: #ffd21f;
      --accent-strong: #f2b705;
      --shadow: 0 18px 50px rgba(0, 0, 0, 0.45);
      --radius: 8px;
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }}
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; color: var(--ink); background: var(--paper); line-height: 1.5; }}
    a {{ color: inherit; text-decoration: none; }}
    .site-header {{
      position: sticky; top: 0; z-index: 10; background: rgba(0, 0, 0, 0.92);
      border-bottom: 1px solid rgba(255, 210, 31, 0.24); backdrop-filter: blur(16px);
    }}
    .nav {{ width: min(1240px, calc(100% - 32px)); min-height: 70px; margin: 0 auto; display: flex; align-items: center; justify-content: space-between; gap: 20px; }}
    .brand {{ display: inline-flex; align-items: center; gap: 12px; font-weight: 900; }}
    .brand-mark {{ width: 38px; height: 38px; border-radius: 8px; display: grid; place-items: center; background: var(--accent); color: #050505; font-weight: 950; }}
    .nav-meta {{ color: var(--muted); font-size: 0.92rem; }}
    .nav-actions {{ display: inline-flex; align-items: center; gap: 10px; }}
    .profile-menu {{ position: relative; display: inline-flex; align-items: center; }}
    .profile-button {{ width: 40px; height: 40px; display: grid; place-items: center; border: 1px solid rgba(255,210,31,.58); border-radius: 50%; background: var(--accent); color: #050505; font: inherit; font-size: 13px; font-weight: 950; cursor: pointer; box-shadow: 0 10px 26px rgba(255,210,31,.14); }}
    .profile-button:hover {{ background: var(--accent-strong); }}
    .profile-button:focus-visible {{ outline: 2px solid rgba(255,210,31,.62); outline-offset: 3px; }}
    .profile-dropdown {{ position: absolute; top: calc(100% + 10px); right: 0; width: min(320px, calc(100vw - 32px)); display: none; padding: 12px; border: 1px solid rgba(255,210,31,.26); border-radius: 8px; background: #101010; box-shadow: var(--shadow); }}
    .profile-dropdown.is-open {{ display: grid; gap: 10px; }}
    .profile-dropdown-title {{ color: var(--accent); font-size: 15px; font-weight: 950; line-height: 1.2; }}
    .profile-dropdown-row {{ display: grid; gap: 2px; padding-top: 8px; border-top: 1px solid var(--line); }}
    .profile-dropdown-row span {{ color: var(--muted); font-size: 10px; font-weight: 950; text-transform: uppercase; }}
    .profile-dropdown-row strong {{ display: block; color: var(--ink); font-size: 13px; line-height: 1.35; }}
    .run-status {{ min-width: 96px; color: var(--muted); font-size: .82rem; }}
    .run-loader {{ --progress-angle: 3.6deg; display: none; align-items: center; gap: 8px; color: var(--accent); font-size: .82rem; font-weight: 950; }}
    .run-loader.is-active {{ display: inline-flex; }}
    .run-loader-ring {{ position: relative; width: 34px; height: 34px; border-radius: 50%; display: grid; place-items: center; background: conic-gradient(var(--accent) var(--progress-angle), rgba(255,255,255,.14) 0); box-shadow: 0 0 0 1px rgba(255,210,31,.2); }}
    .run-loader-ring::after {{ content: ""; width: 24px; height: 24px; border-radius: 50%; background: #050505; position: absolute; }}
    .run-loader-percent {{ position: relative; z-index: 1; color: var(--ink); font-size: 9px; line-height: 1; }}
    .hero {{ position: relative; overflow: hidden; border-bottom: 1px solid rgba(255, 210, 31, .18); }}
    .hero::before {{
      content: ""; position: absolute; inset: 0;
      background: linear-gradient(90deg, rgba(0,0,0,.98), rgba(0,0,0,.72) 52%, rgba(255,210,31,.12)),
                  radial-gradient(circle at 78% 8%, rgba(255,210,31,.22), transparent 34%);
    }}
    .hero-content {{ position: relative; width: min(1240px, calc(100% - 32px)); margin: 0 auto; padding: 72px 0 34px; }}
    .eyebrow {{ margin: 0 0 10px; color: var(--accent); font-size: .78rem; font-weight: 950; text-transform: uppercase; letter-spacing: 0; }}
    h1 {{ max-width: 900px; margin: 0; font-size: clamp(2.5rem, 7vw, 5.8rem); line-height: .96; letter-spacing: 0; }}
    .hero-name {{ display: block; color: var(--accent); }}
    .hero-title {{ display: block; }}
    .hero-copy {{ max-width: 720px; margin: 20px 0 0; color: rgba(255,255,255,.82); font-size: 1.08rem; }}
    .hero-stats {{ display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 1px; width: min(760px, 100%); margin: 30px 0 0; background: rgba(255,210,31,.25); border: 1px solid rgba(255,210,31,.28); border-radius: var(--radius); overflow: hidden; }}
    .hero-stats div {{ padding: 18px; background: rgba(0,0,0,.72); }}
    .hero-stats dt {{ color: var(--accent); font-size: 1.9rem; font-weight: 950; line-height: 1; }}
    .hero-stats dd {{ margin: 6px 0 0; color: var(--muted); font-size: .86rem; }}
    main {{ width: min(1240px, calc(100% - 32px)); margin: 0 auto; padding: 28px 0 48px; }}
    .controls {{ display: grid; grid-template-columns: minmax(220px, 1.4fr) minmax(180px, .7fr) minmax(180px, .7fr) auto; gap: 12px; margin-bottom: 18px; }}
    input, select {{ width: 100%; min-height: 46px; border: 1px solid rgba(255,210,31,.24); border-radius: 8px; background: var(--panel); color: var(--ink); padding: 0 13px; font: inherit; }}
    input:focus, select:focus {{ outline: 2px solid rgba(255,210,31,.5); outline-offset: 0; }}
    .button {{ min-height: 40px; display: inline-flex; align-items: center; justify-content: center; border: 1px solid transparent; border-radius: 8px; padding: 10px 13px; font-weight: 950; transition: transform 160ms ease, background 160ms ease, border-color 160ms ease; white-space: nowrap; }}
    .button:hover {{ transform: translateY(-1px); }}
    .button.primary {{ color: #050505; background: var(--accent); box-shadow: 0 12px 28px rgba(255,210,31,.16); }}
    .button.primary:hover {{ background: var(--accent-strong); }}
    .button.secondary {{ color: var(--ink); border-color: rgba(255,210,31,.42); background: rgba(255,255,255,.06); }}
    .button:disabled {{ cursor: wait; opacity: .7; transform: none; }}
    .download-button {{ min-height: 34px; padding: 7px 10px; font-size: 12px; }}
    body.is-blocked {{ overflow: hidden; }}
    .resume-download-overlay {{ position: fixed; inset: 0; z-index: 80; display: none; place-items: center; padding: 20px; background: rgba(0,0,0,.78); backdrop-filter: blur(10px); }}
    .resume-download-overlay.is-active {{ display: grid; }}
    .resume-download-panel {{ width: min(360px, 100%); display: grid; justify-items: center; gap: 14px; padding: 24px; border: 1px solid rgba(255,210,31,.28); border-radius: 8px; background: #101010; box-shadow: var(--shadow); text-align: center; }}
    .resume-download-ring {{ --download-progress-angle: 3.6deg; position: relative; width: 92px; height: 92px; border-radius: 50%; display: grid; place-items: center; background: conic-gradient(var(--accent) var(--download-progress-angle), rgba(255,255,255,.14) 0); box-shadow: 0 0 0 1px rgba(255,210,31,.18), 0 18px 48px rgba(0,0,0,.45); }}
    .resume-download-ring::after {{ content: ""; position: absolute; width: 68px; height: 68px; border-radius: 50%; background: #050505; }}
    .resume-download-percent {{ position: relative; z-index: 1; color: var(--ink); font-size: 18px; font-weight: 950; }}
    .resume-download-title {{ color: var(--accent); font-size: 12px; font-weight: 950; text-transform: uppercase; }}
    .resume-download-message {{ color: var(--ink); font-size: 16px; font-weight: 900; }}
    .resume-download-sub {{ max-width: 280px; color: var(--muted); font-size: 12px; }}
    .table-wrap {{ overflow: auto; border: 1px solid var(--line); border-radius: var(--radius); background: var(--panel); box-shadow: var(--shadow); }}
    table {{ width: 100%; min-width: 1080px; border-collapse: collapse; }}
    th, td {{ border-bottom: 1px solid var(--line); padding: 13px 12px; vertical-align: top; font-size: 13px; }}
    th {{ position: sticky; top: 0; z-index: 2; text-align: left; color: var(--accent); background: #0b0b0b; font-size: 11px; text-transform: uppercase; letter-spacing: 0; }}
    tr:hover {{ background: rgba(255,255,255,.035); }}
    strong {{ color: var(--ink); font-size: 14px; }}
    span {{ color: var(--muted); }}
    .score {{ display: inline-grid; place-items: center; width: 34px; height: 34px; border-radius: 8px; background: var(--accent); color: #050505; font-weight: 950; }}
    .source-chip {{ display: inline-flex; margin-top: 7px; padding: 3px 7px; border: 1px solid rgba(255,210,31,.28); border-radius: 999px; color: var(--accent); font-size: 11px; }}
    .salary-estimate {{ display: inline-flex; align-items: center; gap: 6px; color: var(--ink); font-weight: 850; }}
    .salary-info {{ position: relative; width: 18px; height: 18px; display: inline-grid; place-items: center; flex: 0 0 auto; border: 1px solid rgba(255,210,31,.5); border-radius: 50%; color: var(--accent); background: rgba(255,210,31,.08); font-size: 11px; font-weight: 950; cursor: help; }}
    .salary-info:focus-visible {{ outline: 2px solid rgba(255,210,31,.55); outline-offset: 2px; }}
    .salary-tooltip {{ position: absolute; right: 0; bottom: calc(100% + 10px); z-index: 8; width: min(320px, calc(100vw - 40px)); display: none; padding: 10px 11px; border: 1px solid rgba(255,210,31,.28); border-radius: 8px; background: #101010; color: var(--ink); box-shadow: var(--shadow); font-size: 12px; font-weight: 600; line-height: 1.45; text-align: left; }}
    .salary-info:hover .salary-tooltip, .salary-info:focus .salary-tooltip, .salary-info.is-open .salary-tooltip {{ display: block; }}
    .keywords {{ color: rgba(255,255,255,.76); }}
    .empty {{ display: none; padding: 26px; text-align: center; color: var(--muted); border: 1px solid var(--line); border-radius: var(--radius); background: var(--panel); }}
    .count {{ color: var(--accent); font-weight: 950; }}
    @media (max-width: 820px) {{
      .nav {{ min-height: 62px; }}
      .nav-actions {{ gap: 8px; }}
      .run-status {{ display: none; }}
      .run-loader-label {{ display: none; }}
      .hero-content {{ padding-top: 46px; }}
      .hero-stats {{ grid-template-columns: 1fr; }}
      .controls {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <header class="site-header">
    <nav class="nav" aria-label="Dashboard">
      <a class="brand" href="/"><span class="brand-mark">JA</span><span>Job Agent</span></a>
      <div class="nav-actions">
        <button class="button secondary" id="runAgentButton" type="button">Run fresh report</button>
        <span class="run-status" id="runAgentStatus"></span>
        <span class="run-loader" id="runLoader" aria-live="polite" aria-label="Fresh report progress">
          <span class="run-loader-ring"><span class="run-loader-percent" id="runProgressPercent">1%</span></span>
          <span class="run-loader-label" id="runProgressLabel">Starting</span>
        </span>
        <div class="profile-menu">
          <button class="profile-button" id="profileButton" type="button" aria-label="Open profile" aria-expanded="false">{profile_initials}</button>
          <div class="profile-dropdown" id="profileDropdown" role="menu" aria-label="Candidate profile">
            <div class="profile-dropdown-title">{profile_name}</div>
            <div class="profile-dropdown-row"><span>Experience</span><strong>{profile_experience}</strong></div>
            <div class="profile-dropdown-row"><span>Current role</span><strong>{profile_position}</strong></div>
            <div class="profile-dropdown-row"><span>Company</span><strong>{profile_company}</strong></div>
          </div>
        </div>
      </div>
    </nav>
  </header>
  <section class="hero">
    <div class="hero-content">
      <p class="eyebrow">Fresh jobs. Tailored resumes.</p>
      <h1><span class="hero-name">Hello {hero_name}</span><span class="hero-title">Latest job matches for your next move.</span></h1>
      <p class="hero-copy">Filter the run, open the posting, and generate an ATS-friendly resume only when a role is worth applying to.</p>
      <dl class="hero-stats">
        <div><dt>{resume_count}</dt><dd>matched jobs today</dd></div>
        <div><dt>{top_score}</dt><dd>highest match score</dd></div>
        <div><dt>{india_count}</dt><dd>India-focused listings</dd></div>
      </dl>
    </div>
  </section>
  <main>
    <p class="eyebrow">Last run {html.escape(run_timestamp)} · <span class="count" id="visibleCount">{len(records)}</span> visible</p>
    <div class="controls" aria-label="Job filters">
      <input id="searchFilter" type="search" placeholder="Search title, company, keyword..." autocomplete="off">
      <select id="sourceFilter" aria-label="Filter by source">{source_options}</select>
      <select id="locationFilter" aria-label="Filter by location">{location_options}</select>
      <button class="button secondary" id="clearFilters" type="button">Clear</button>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Score</th><th>Posted / Source</th><th>Job</th><th>Location</th>
            <th>Salary</th><th>Matched keywords</th><th>Apply</th><th>Download tailored resume</th>
          </tr>
        </thead>
        <tbody id="jobsBody">
          {''.join(html_rows)}
        </tbody>
      </table>
    </div>
    <div class="empty" id="emptyState">No jobs match those filters.</div>
  </main>
  <div class="resume-download-overlay" id="resumeDownloadOverlay" aria-live="polite" aria-modal="true" role="dialog" aria-label="Preparing tailored resume">
    <div class="resume-download-panel">
      <div class="resume-download-ring" id="resumeDownloadRing"><span class="resume-download-percent" id="resumeDownloadPercent">1%</span></div>
      <div class="resume-download-title">Tailored resume</div>
      <div class="resume-download-message" id="resumeDownloadMessage">Connecting to OpenAI</div>
      <div class="resume-download-sub">Please keep this window open while your job-specific resume is prepared.</div>
    </div>
  </div>
  <script>
    const sourceFilter = document.getElementById('sourceFilter');
    const locationFilter = document.getElementById('locationFilter');
    const searchFilter = document.getElementById('searchFilter');
    const clearFilters = document.getElementById('clearFilters');
    const rows = Array.from(document.querySelectorAll('#jobsBody tr'));
    const visibleCount = document.getElementById('visibleCount');
    const emptyState = document.getElementById('emptyState');
    const runAgentButton = document.getElementById('runAgentButton');
    const runAgentStatus = document.getElementById('runAgentStatus');
    const runLoader = document.getElementById('runLoader');
    const runProgressPercent = document.getElementById('runProgressPercent');
    const runProgressLabel = document.getElementById('runProgressLabel');
    const profileButton = document.getElementById('profileButton');
    const profileDropdown = document.getElementById('profileDropdown');
    const resumeDownloadOverlay = document.getElementById('resumeDownloadOverlay');
    const resumeDownloadRing = document.getElementById('resumeDownloadRing');
    const resumeDownloadPercent = document.getElementById('resumeDownloadPercent');
    const resumeDownloadMessage = document.getElementById('resumeDownloadMessage');
    const salaryInfoMarkers = Array.from(document.querySelectorAll('.salary-info'));
    let runProgressTimer = null;
    let runProgressValue = 1;
    let resumeDownloadTimer = null;
    let resumeDownloadMessageTimer = null;
    let resumeDownloadProgress = 1;
    const resumeDownloadMessages = [
      'Connecting to OpenAI',
      'Reading the job description',
      'Creating your tailored resume',
      'Publishing the download'
    ];

    function applyFilters() {{
      const source = sourceFilter.value;
      const location = locationFilter.value;
      const query = searchFilter.value.trim().toLowerCase();
      let count = 0;
      rows.forEach((row) => {{
        const sourceOk = !source || row.dataset.source === source;
        const locationOk = !location || row.dataset.location === location;
        const searchOk = !query || row.dataset.search.includes(query);
        const visible = sourceOk && locationOk && searchOk;
        row.style.display = visible ? '' : 'none';
        if (visible) count += 1;
      }});
      visibleCount.textContent = count;
      emptyState.style.display = count ? 'none' : 'block';
    }}

    [sourceFilter, locationFilter, searchFilter].forEach((el) => el.addEventListener('input', applyFilters));
    clearFilters.addEventListener('click', () => {{
      sourceFilter.value = '';
      locationFilter.value = '';
      searchFilter.value = '';
      applyFilters();
    }});

    async function runFreshReport() {{
      runAgentButton.disabled = true;
      runAgentButton.textContent = 'Running...';
      runAgentStatus.textContent = 'Searching jobs';
      startRunProgress('Searching jobs');
      try {{
        const response = await fetch('/run-agent', {{ method: 'POST' }});
        const payload = await response.json();
        if (!response.ok || !payload.ok) {{
          throw new Error(payload.message || payload.stderr || 'Run failed');
        }}
        runAgentStatus.textContent = 'Refreshing';
        completeRunProgress('Refreshing');
        window.setTimeout(() => window.location.reload(), 450);
      }} catch (error) {{
        stopRunProgress();
        runAgentStatus.textContent = error.message || 'Run failed';
        runAgentButton.disabled = false;
        runAgentButton.textContent = 'Run fresh report';
      }}
    }}

    function setRunProgress(value, label) {{
      runProgressValue = Math.max(1, Math.min(100, Math.round(value)));
      if (runLoader) {{
        runLoader.classList.add('is-active');
        runLoader.style.setProperty('--progress-angle', `${{runProgressValue * 3.6}}deg`);
      }}
      if (runProgressPercent) runProgressPercent.textContent = `${{runProgressValue}}%`;
      if (runProgressLabel && label) runProgressLabel.textContent = label;
    }}

    function startRunProgress(label) {{
      stopRunProgress(false);
      setRunProgress(1, label || 'Starting');
      runProgressTimer = window.setInterval(() => {{
        const step = runProgressValue < 70 ? 3 : runProgressValue < 92 ? 2 : 1;
        if (runProgressValue < 98) setRunProgress(runProgressValue + step, label || 'Running');
      }}, 650);
    }}

    function completeRunProgress(label) {{
      stopRunProgress(false);
      setRunProgress(100, label || 'Done');
    }}

    function stopRunProgress(hide = true) {{
      if (runProgressTimer) {{
        window.clearInterval(runProgressTimer);
        runProgressTimer = null;
      }}
      if (hide && runLoader) runLoader.classList.remove('is-active');
    }}

    function setResumeDownloadProgress(value, message) {{
      resumeDownloadProgress = Math.max(1, Math.min(100, Math.round(value)));
      if (resumeDownloadRing) resumeDownloadRing.style.setProperty('--download-progress-angle', `${{resumeDownloadProgress * 3.6}}deg`);
      if (resumeDownloadPercent) resumeDownloadPercent.textContent = `${{resumeDownloadProgress}}%`;
      if (resumeDownloadMessage && message) resumeDownloadMessage.textContent = message;
    }}

    function startResumeDownloadLoader() {{
      stopResumeDownloadLoader(false);
      document.body.classList.add('is-blocked');
      if (resumeDownloadOverlay) resumeDownloadOverlay.classList.add('is-active');
      setResumeDownloadProgress(1, resumeDownloadMessages[0]);
      let messageIndex = 0;
      resumeDownloadMessageTimer = window.setInterval(() => {{
        messageIndex = Math.min(messageIndex + 1, resumeDownloadMessages.length - 1);
        setResumeDownloadProgress(Math.max(resumeDownloadProgress, 20 + messageIndex * 22), resumeDownloadMessages[messageIndex]);
      }}, 1700);
      resumeDownloadTimer = window.setInterval(() => {{
        const step = resumeDownloadProgress < 55 ? 4 : resumeDownloadProgress < 88 ? 2 : 1;
        if (resumeDownloadProgress < 96) setResumeDownloadProgress(resumeDownloadProgress + step);
      }}, 420);
    }}

    function stopResumeDownloadLoader(hide = true) {{
      if (resumeDownloadTimer) {{
        window.clearInterval(resumeDownloadTimer);
        resumeDownloadTimer = null;
      }}
      if (resumeDownloadMessageTimer) {{
        window.clearInterval(resumeDownloadMessageTimer);
        resumeDownloadMessageTimer = null;
      }}
      if (hide) {{
        if (resumeDownloadOverlay) resumeDownloadOverlay.classList.remove('is-active');
        document.body.classList.remove('is-blocked');
      }}
    }}

    function filenameFromDisposition(value) {{
      const match = /filename\\*?=(?:UTF-8''|")?([^";]+)/i.exec(value || '');
      if (!match) return '';
      return decodeURIComponent(match[1].replace(/"/g, '').trim());
    }}

    async function downloadTailoredResume(event) {{
      event.preventDefault();
      const link = event.currentTarget;
      if (!link || link.dataset.loading === 'true') return;
      link.dataset.loading = 'true';
      startResumeDownloadLoader();
      try {{
        const response = await fetch(link.href);
        if (!response.ok) {{
          const message = await response.text();
          throw new Error(message || 'Resume generation failed');
        }}
        setResumeDownloadProgress(92, 'Publishing the download');
        const blob = await response.blob();
        const fileName = filenameFromDisposition(response.headers.get('Content-Disposition')) || 'tailored_resume.docx';
        const downloadUrl = URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = downloadUrl;
        anchor.download = fileName;
        document.body.appendChild(anchor);
        setResumeDownloadProgress(100, 'Download ready');
        anchor.click();
        anchor.remove();
        window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 30000);
        window.setTimeout(() => stopResumeDownloadLoader(), 650);
      }} catch (error) {{
        setResumeDownloadProgress(100, error.message || 'Resume generation failed');
        window.setTimeout(() => stopResumeDownloadLoader(), 2400);
      }} finally {{
        link.dataset.loading = 'false';
      }}
    }}

    runAgentButton.addEventListener('click', runFreshReport);
    document.querySelectorAll('.download-button').forEach((link) => link.addEventListener('click', downloadTailoredResume));
    salaryInfoMarkers.forEach((marker) => {{
      marker.addEventListener('click', (event) => {{
        event.stopPropagation();
        const shouldOpen = !marker.classList.contains('is-open');
        salaryInfoMarkers.forEach((item) => item.classList.remove('is-open'));
        marker.classList.toggle('is-open', shouldOpen);
      }});
    }});
    document.addEventListener('click', () => salaryInfoMarkers.forEach((marker) => marker.classList.remove('is-open')));
    document.addEventListener('keydown', (event) => {{
      if (event.key === 'Escape') salaryInfoMarkers.forEach((marker) => marker.classList.remove('is-open'));
    }});
    if (profileButton && profileDropdown) {{
      profileButton.addEventListener('click', (event) => {{
        event.stopPropagation();
        const shouldOpen = !profileDropdown.classList.contains('is-open');
        profileDropdown.classList.toggle('is-open', shouldOpen);
        profileButton.setAttribute('aria-expanded', shouldOpen ? 'true' : 'false');
      }});
      document.addEventListener('click', (event) => {{
        if (!profileDropdown.contains(event.target) && event.target !== profileButton) {{
          profileDropdown.classList.remove('is-open');
          profileButton.setAttribute('aria-expanded', 'false');
        }}
      }});
      document.addEventListener('keydown', (event) => {{
        if (event.key === 'Escape') {{
          profileDropdown.classList.remove('is-open');
          profileButton.setAttribute('aria-expanded', 'false');
        }}
      }});
    }}
  </script>
</body>
</html>
"""
    html_path.write_text(html_doc, encoding="utf-8")
    return json_path, html_path


def read_existing_xlsx_rows(path: Path) -> List[List[str]]:
    if not path.exists():
        return [HEADERS]
    try:
        with zipfile.ZipFile(path) as archive:
            shared = []
            if "xl/sharedStrings.xml" in archive.namelist():
                xml = archive.read("xl/sharedStrings.xml").decode("utf-8", errors="replace")
                shared = [html.unescape(re.sub(r"<[^>]+>", "", match)) for match in re.findall(r"<si>(.*?)</si>", xml, re.S)]
            sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8", errors="replace")
        rows = []
        for row_xml in re.findall(r"<row[^>]*>(.*?)</row>", sheet, re.S):
            values = []
            cells = re.findall(r"<c([^>]*)>(.*?)</c>", row_xml, re.S)
            for attrs, cell_xml in cells:
                value_match = re.search(r"<v>(.*?)</v>", cell_xml, re.S)
                if not value_match:
                    values.append("")
                    continue
                value = html.unescape(value_match.group(1))
                if 't="s"' in attrs:
                    idx = int(value)
                    values.append(shared[idx] if idx < len(shared) else "")
                else:
                    values.append(value)
            rows.append(values)
        return rows or [HEADERS]
    except Exception:
        backup = path.with_suffix(".corrupt-backup.xlsx")
        path.rename(backup)
        print(f"Warning: existing workbook could not be read and was moved to {backup}", file=sys.stderr)
        return [HEADERS]


def column_letter(index: int) -> str:
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def write_xlsx(path: Path, rows: Sequence[Sequence[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    shared_strings: List[str] = []
    shared_index: Dict[str, int] = {}
    hyperlinks: List[Tuple[str, str, str]] = []

    def sst_index(value: str) -> int:
        if value not in shared_index:
            shared_index[value] = len(shared_strings)
            shared_strings.append(value)
        return shared_index[value]

    sheet_rows = []
    for r_idx, row in enumerate(rows, start=1):
        cells = []
        for c_idx, value in enumerate(row, start=1):
            value = "" if value is None else str(value)
            ref = f"{column_letter(c_idx)}{r_idx}"
            if value.startswith("http://") or value.startswith("https://"):
                rel_id = f"rIdLink{len(hyperlinks) + 1}"
                hyperlinks.append((ref, rel_id, value))
                style = ' s="2"' if r_idx != 1 else ' s="1"'
                cells.append(f'<c r="{ref}" t="s"{style}><v>{sst_index(value)}</v></c>')
            else:
                style = ' s="1"' if r_idx == 1 else ""
                cells.append(f'<c r="{ref}" t="s"{style}><v>{sst_index(value)}</v></c>')
        sheet_rows.append(f'<row r="{r_idx}">{"".join(cells)}</row>')

    hyperlink_xml = ""
    if hyperlinks:
        hyperlink_xml = "<hyperlinks>" + "".join(
            f'<hyperlink ref="{ref}" r:id="{rel_id}"/>' for ref, rel_id, _ in hyperlinks
        ) + "</hyperlinks>"
    sheet_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(
            f'<Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target={quoteattr(url)} TargetMode="External"/>'
            for _, rel_id, url in hyperlinks
        )
        + "</Relationships>"
    )
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        '<cols>'
        '<col min="1" max="2" width="12" customWidth="1"/>'
        '<col min="3" max="6" width="24" customWidth="1"/>'
        '<col min="7" max="10" width="22" customWidth="1"/>'
        '<col min="11" max="11" width="58" customWidth="1"/>'
        '<col min="12" max="12" width="24" customWidth="1"/>'
        '</cols>'
        f'<sheetData>{"".join(sheet_rows)}</sheetData>'
        f'<autoFilter ref="A1:L{max(1, len(rows))}"/>'
        f"{hyperlink_xml}"
        '</worksheet>'
    )
    shared_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">'
        + "".join(f"<si><t>{escape(value)}</t></si>" for value in shared_strings)
        + "</sst>"
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Job Matches" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/></font>'
        '<font><u/><color rgb="FF0563C1"/><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        '</Relationships>'
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        '</Types>'
    )

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", content_types_xml)
            zf.writestr("_rels/.rels", rels_xml)
            zf.writestr("xl/workbook.xml", workbook_xml)
            zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
            zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
            if hyperlinks:
                zf.writestr("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels_xml)
            zf.writestr("xl/styles.xml", styles_xml)
            zf.writestr("xl/sharedStrings.xml", shared_xml)
        tmp_path.replace(path)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def run(config_path: Path) -> int:
    config = load_config(config_path)
    config_dir = config_path.parent
    run_date = today_for_timezone(config.get("timezone", "Asia/Kolkata"))
    resume_path = resolve_path(config.get("resume_path", ""), config_dir)
    output_path = resolve_path(config.get("output_xlsx", "data/job_matches.xlsx"), config_dir)
    db_path = resolve_path(config.get("database_path", "data/seen_jobs.sqlite3"), config_dir)

    resume_profile = extract_resume_profile(resume_path, config, config_dir)
    usd_to_inr = get_usd_to_inr_rate(config)
    jobs = collect_jobs(config)
    filtered = filter_jobs(jobs, config, resume_profile, run_date, usd_to_inr)
    latest_json_path, latest_html_path = write_latest_jobs_outputs(filtered, run_date, config, config_dir, resume_profile)

    conn = init_db(db_path)
    new_jobs = unseen_jobs(conn, filtered)
    rows = read_existing_xlsx_rows(output_path)
    if rows and rows[0] != HEADERS and rows[0][:6] == HEADERS[:6]:
        rows[0] = HEADERS
    if len(rows) > 1 and rows[1][:6] == HEADERS[:6]:
        rows = [rows[0]] + rows[2:]
    if not rows or rows[0] != HEADERS:
        rows = [HEADERS] + rows
    rows.extend(row_for_job(job, run_date) for job in new_jobs)
    write_xlsx(output_path, rows)
    mark_seen(conn, new_jobs)

    print(
        json.dumps(
            {
                "run_date": run_date.isoformat(),
                "searched_jobs": len(jobs),
                "matching_jobs_today": len(filtered),
                "new_jobs_added": len(new_jobs),
                "output_xlsx": str(output_path),
                "latest_jobs_json": str(latest_json_path),
                "latest_jobs_html": str(latest_html_path),
                "dashboard_url": f"http://127.0.0.1:{int(config.get('dashboard', {}).get('port', 8765))}",
                "usd_to_inr": usd_to_inr,
                "resume_profile_source": resume_profile.get("source", "unknown"),
                "resume_keywords": keywords_from_profile(resume_profile)[:15],
            },
            indent=2,
        )
    )
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Find today's matching jobs and append them to Excel.")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG), help="Path to config.json")
    parser.add_argument(
        "--print-resume-profile",
        action="store_true",
        help="Only parse the resume and print the extracted profile without searching jobs",
    )
    args = parser.parse_args()
    try:
        config_path = Path(args.config).expanduser().resolve()
        if args.print_resume_profile:
            config = load_config(config_path)
            resume_path = resolve_path(config.get("resume_path", ""), config_path.parent)
            profile = extract_resume_profile(resume_path, config, config_path.parent)
            print(json.dumps(profile, indent=2, ensure_ascii=False))
            return 0
        return run(config_path)
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
