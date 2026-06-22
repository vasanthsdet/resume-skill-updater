"""
Extracts specific technical skills, tools, and technologies from raw recruiter
requirement text using Claude Haiku (fast + cheap for short extraction tasks).
"""

import os
import re
import json
import httpx
import anthropic
from dotenv import load_dotenv

load_dotenv()


def _client() -> anthropic.Anthropic:
    key = os.getenv("ANTHROPIC_API_KEY")
    if not key:
        raise RuntimeError("ANTHROPIC_API_KEY not set — add it to .env")
    return anthropic.Anthropic(api_key=key, http_client=httpx.Client(verify=False))


def extract_skills(recruiter_text: str) -> list[str]:
    """
    Parse raw recruiter requirements and return a clean list of
    specific skills, tools, languages, and platforms mentioned.
    """
    prompt = f"""You are a technical recruiter assistant. Extract every specific technical skill,
tool, framework, language, platform, and methodology mentioned in the job requirements below.

JOB REQUIREMENTS:
{recruiter_text}

Rules:
- Include specific tools/tech only: "Playwright", "Cypress", "TypeScript", "Azure DevOps", "DataDog", "SQL", "CI/CD", "Agile", "NPM", "Docker", etc.
- Exclude vague phrases like "6+ years of experience", "proven track record", "strong communication".
- If a tool or language is mentioned with alternatives (e.g. "JavaScript and/or TypeScript"), include both.
- Deduplicate — list each skill only once.
- Keep names as they appear in the JD (e.g. "DataDog" not "datadog").

Return ONLY a valid JSON array of strings. No explanation, no markdown:
["skill1", "skill2", ...]"""

    resp = _client().messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=600,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = resp.content[0].text.strip()
    raw = re.sub(r"^```(?:json)?\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)
    return json.loads(raw)
