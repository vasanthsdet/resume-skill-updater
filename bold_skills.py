"""
Post-process the rewritten resume to apply bold formatting:
  1. Skills section  : category label bold, values normal font
  2. Summary bullets : technology / tool names bold inline
  3. All job bullets : technology / tool names bold inline

Usage:
  python bold_skills.py
  python bold_skills.py --input resume/Revathi_Battina_Resume.docx --output resume/Revathi_Battina_Resume.docx
"""

import re, shutil, argparse
from docx import Document

NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

SUMMARY_KEYWORDS    = {"professional summary", "summary", "profile", "objective"}
SKILLS_KEYWORDS     = {"technical skills", "skills", "core competencies", "expertise"}
EXPERIENCE_KEYWORDS = {"professional experience", "work experience", "employment history"}

# Extra skill terms that may appear in bullets but live outside the skills section
EXTRA_TERMS = [
    "CI/CD", "BDD", "DDD", "TDD", "REST", "gRPC", "SOAP", "SQL", "NoSQL",
    "Linux", "AWS", "GCP", "Azure", "Docker", "Kubernetes", "Git", "GitHub",
    "GitLab", "JIRA", "X-Ray", "Agile", "Scrum", "Swagger", "OAuth", "JWT",
    "HL7", "EDI", "Playwright", "Cypress", "Selenium", "Appium", "TestNG",
    "JUnit", "Maven", "Gradle", "Jenkins", "Python", "Java", "TypeScript",
    "JavaScript", "Groovy", "Postman", "Bruno", "JMeter", "LoadRunner",
    "Dynatrace", "Splunk", "ELK Stack", "Prometheus", "Grafana", "CloudWatch",
    "Salesforce", "Espresso", "XCUITest", "Perfecto", "WebDriverIO",
    "SonarQube", "IntelliJ", "VS Code", "Confluence", "Cucumber", "Gherkin",
    "Spock", "Mocha", "Gauge", "Robot Framework", "Karate", "REST Assured",
    "Axios", "Charles Proxy", "SoapUI", "DataDog", "Azure DevOps",
    "GitHub Actions", "GitLab CI", "AWS Lambda", "API Gateway", "DynamoDB",
    "CodeCommit", "S3", "Xcode", "Android Studio", "GitHub Copilot", "ChatGPT",
    "Claude", "NPM", "NodeJS", "Go", "Ruby", "Groovy", "Oracle", "PostgreSQL",
    "MySQL", "MongoDB", "Couchbase", "Selenium WebDriver", "WebDriverIO",
    "Insomnia", "Karate DSL", "Gauge", "Perfecto Mobile",
]


# ── Document structure helpers ─────────────────────────────────────────────────

def _is_heading(para) -> bool:
    text = para.text.strip()
    if not text:
        return False
    if "heading" in para.style.name.lower():
        return True
    if text.rstrip().endswith(":") and len(text) < 60:
        return True
    has_bold = para.runs and all(r.bold for r in para.runs if r.text.strip())
    if has_bold and len(text) < 80:
        colon_idx = text.find(":")
        if 0 < colon_idx < len(text) - 2:
            return False
        return True
    return False


def _find_section(doc, keywords):
    paras = doc.paragraphs
    for i, para in enumerate(paras):
        lower = para.text.strip().lower()
        if any(kw in lower for kw in keywords) and _is_heading(para):
            content = []
            j = i + 1
            while j < len(paras):
                if _is_heading(paras[j]) and paras[j].text.strip():
                    break
                if paras[j].text.strip():
                    content.append(j)
                j += 1
            return i, content
    return -1, []


def _find_all_job_blocks(doc, exp_heading_idx):
    paras = doc.paragraphs
    jobs = []
    label, bullets, in_job = "", [], False
    i = exp_heading_idx + 1
    while i < len(paras):
        text = paras[i].text.strip()
        if not text:
            i += 1
            continue
        lower = text.lower()
        if lower.rstrip().endswith(":") and any(
            kw in lower for kw in {"education", "certification", "award"}
        ):
            break
        if " | " in text:
            if in_job:
                jobs.append((label, bullets[:]))
            label, bullets, in_job = text, [], True
            i += 1
            if i < len(paras):
                nxt = paras[i].text.strip()
                if nxt and len(nxt) < 70 and " | " not in nxt and not re.search(r"\d{4}", nxt):
                    i += 1
        elif in_job:
            bullets.append(i)
            i += 1
        else:
            i += 1
    if in_job:
        jobs.append((label, bullets[:]))
    return jobs


# ── Run-level helpers ──────────────────────────────────────────────────────────

def _get_base_fmt(para) -> dict:
    """Capture font name and size from the first non-empty run."""
    for r in para.runs:
        if r.text.strip():
            return {"name": r.font.name, "size": r.font.size}
    return {}


def _clear_runs(para):
    """Remove every <w:r> element from the paragraph XML."""
    for r_elem in para._element.findall(f"{{{NS}}}r"):
        para._element.remove(r_elem)


def _write_segments(para, segments: list[tuple[str, bool]], base_fmt: dict):
    """Add (text, is_bold) as new runs into para (must be already cleared)."""
    for text, is_bold in segments:
        if not text:
            continue
        run = para.add_run(text)
        run.bold = is_bold
        if base_fmt.get("name"):
            run.font.name = base_fmt["name"]
        if base_fmt.get("size"):
            run.font.size = base_fmt["size"]


# ── Skill-term regex ───────────────────────────────────────────────────────────

def _build_pattern(skill_terms: list[str]) -> re.Pattern:
    """Regex matching any skill term; longest terms checked first."""
    sorted_terms = sorted(set(skill_terms), key=len, reverse=True)
    parts = []
    for t in sorted_terms:
        if not t:
            continue
        prefix = r"\b" if re.match(r"\w", t[0]) else ""
        suffix = r"\b" if re.match(r"\w", t[-1]) else ""
        parts.append(prefix + re.escape(t) + suffix)
    if not parts:
        return re.compile(r"(?!)")   # never matches
    return re.compile("|".join(parts), re.IGNORECASE)


# ── Formatting functions ───────────────────────────────────────────────────────

def format_skills_line(para):
    """Skills section: bold the category key, normal for the values."""
    text = para.text
    colon_idx = text.find(":")
    if colon_idx < 0:
        return
    fmt = _get_base_fmt(para)
    key_part = text[:colon_idx + 1]   # e.g. "Languages:"
    val_part = text[colon_idx + 1:]   # e.g. " Java, Python, ..."
    _clear_runs(para)
    _write_segments(para, [(key_part, True), (val_part, False)], fmt)


def bold_skills_in_bullet(para, pattern: re.Pattern):
    """Inline-bold every skill/tool occurrence in a bullet paragraph."""
    text = para.text
    if not text.strip():
        return
    segments: list[tuple[str, bool]] = []
    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            segments.append((text[pos:m.start()], False))
        segments.append((m.group(), True))
        pos = m.end()
    if pos < len(text):
        segments.append((text[pos:], False))
    if not any(is_bold for _, is_bold in segments):
        return   # nothing to change
    fmt = _get_base_fmt(para)
    _clear_runs(para)
    _write_segments(para, segments, fmt)


# ── Main ───────────────────────────────────────────────────────────────────────

def apply_bold_formatting(input_path: str, output_path: str):
    if input_path != output_path:
        shutil.copy2(input_path, output_path)
    doc = Document(output_path)

    _, skills_idxs  = _find_section(doc, SKILLS_KEYWORDS)
    _, summary_idxs = _find_section(doc, SUMMARY_KEYWORDS)
    exp_idx, _       = _find_section(doc, EXPERIENCE_KEYWORDS)
    jobs = _find_all_job_blocks(doc, exp_idx) if exp_idx >= 0 else []

    # ── Skills section: key bold, values normal ────────────────────────────────
    print(f"Skills section: formatting {len(skills_idxs)} lines (key bold, values normal)...")
    for idx in skills_idxs:
        format_skills_line(doc.paragraphs[idx])

    # ── Build term list from skills section values ─────────────────────────────
    skill_terms: list[str] = list(EXTRA_TERMS)
    for idx in skills_idxs:
        line = doc.paragraphs[idx].text
        c = line.find(":")
        if c >= 0:
            for term in re.split(r",\s*", line[c + 1:].strip()):
                term = term.strip().strip("()")
                if term and len(term) > 1:
                    skill_terms.append(term)

    pattern = _build_pattern(skill_terms)

    # ── Summary bullets ────────────────────────────────────────────────────────
    print(f"Summary: bolding skills in {len(summary_idxs)} bullets...")
    for idx in summary_idxs:
        bold_skills_in_bullet(doc.paragraphs[idx], pattern)

    # ── Job experience bullets ─────────────────────────────────────────────────
    total = sum(len(b) for _, b in jobs)
    print(f"Experience: bolding skills in {total} bullets across {len(jobs)} jobs...")
    for _, bullet_idxs in jobs:
        for idx in bullet_idxs:
            bold_skills_in_bullet(doc.paragraphs[idx], pattern)

    doc.save(output_path)
    print(f"Saved -> {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Bold skill/tech terms across resume sections")
    parser.add_argument("--input",  default="resume/Revathi_Battina_Resume.docx")
    parser.add_argument("--output", default="resume/Revathi_Battina_Resume.docx")
    args = parser.parse_args()
    print(f"Reading: {args.input}\n")
    apply_bold_formatting(args.input, args.output)
    print("\nDone. Open in Word to verify bold formatting.")


if __name__ == "__main__":
    main()
