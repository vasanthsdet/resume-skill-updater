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

EXTRA_TERMS = [
    "Playwright", "Cypress", "Selenium WebDriver", "Selenium", "Appium",
    "WebDriverIO", "Perfecto", "REST Assured", "Karate", "Postman", "Bruno",
    "Insomnia", "JMeter", "LoadRunner", "gRPC", "SOAP", "SoapUI", "Axios",
    "Charles Proxy", "TestNG", "JUnit", "Cucumber", "Gherkin", "Spock",
    "Mocha", "Gauge", "Robot Framework", "Jenkins", "GitHub Actions",
    "GitLab CI", "Azure DevOps", "Maven", "Gradle", "Docker", "Kubernetes",
    "AWS Lambda", "API Gateway", "DynamoDB", "CodeCommit", "CloudWatch",
    "AWS", "GCP", "Azure", "Git", "GitHub", "GitLab", "JIRA", "X-Ray",
    "SonarQube", "Confluence", "IntelliJ", "VS Code", "Dynatrace", "Splunk",
    "ELK Stack", "Prometheus", "Grafana", "Salesforce", "Espresso", "XCUITest",
    "Xcode", "Android Studio", "GitHub Copilot", "ChatGPT", "Claude",
    "Java", "Python", "TypeScript", "JavaScript", "Ruby", "Groovy", "Go",
    "NodeJS", "NPM", "C#", ".NET", "Oracle", "PostgreSQL", "MySQL",
    "MongoDB", "Couchbase", "SQL", "NoSQL", "Linux",
    "CI/CD", "BDD", "DDD", "TDD", "REST", "Swagger", "OAuth", "JWT",
    "HL7", "EDI", "Agile", "Scrum", "S3",
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



def _remove_ppr_bold(para):
    """Strip bold from paragraph-level run defaults so run-level bold controls it."""
    ppr = para._element.find(f"{{{NS}}}pPr")
    if ppr is None:
        return
    rpr = ppr.find(f"{{{NS}}}rPr")
    if rpr is None:
        return
    for tag in (f"{{{NS}}}b", f"{{{NS}}}bCs"):
        for elem in rpr.findall(tag):
            rpr.remove(elem)


def _rewrite_runs(para, segments: list[tuple[str, bool | None]]):
    """Replace paragraph runs with new segments, setting bold via python-docx API."""
    # Capture font properties before clearing
    first_run = para.runs[0] if para.runs else None
    font_name = first_run.font.name if first_run else None
    font_size = first_run.font.size if first_run else None

    # Remove all existing runs at the XML level
    p_elem = para._element
    for r_elem in list(p_elem.findall(f"{{{NS}}}r")):
        p_elem.remove(r_elem)

    # Re-add runs using python-docx native API — handles bold XML correctly
    for text, is_bold in segments:
        if not text:
            continue
        run = para.add_run(text)
        run.bold = is_bold
        if font_name:
            run.font.name = font_name
        if font_size:
            run.font.size = font_size


# ── Skill-term pattern ─────────────────────────────────────────────────────────

def _build_pattern(skill_terms: list[str]) -> re.Pattern:
    sorted_terms = sorted({t for t in skill_terms if t}, key=len, reverse=True)
    parts = []
    for t in sorted_terms:
        pre  = r"\b" if re.match(r"\w", t[0])  else ""
        post = r"\b" if re.match(r"\w", t[-1]) else ""
        parts.append(pre + re.escape(t) + post)
    return re.compile("|".join(parts), re.IGNORECASE) if parts else re.compile(r"(?!)")


def _split_segments(text: str, pattern: re.Pattern) -> list[tuple[str, bool]]:
    segs, pos = [], 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            segs.append((text[pos:m.start()], False))
        segs.append((m.group(), True))
        pos = m.end()
    if pos < len(text):
        segs.append((text[pos:], False))
    return segs


# ── Public formatting functions ────────────────────────────────────────────────

def format_skills_line(para):
    """Skills section: category key bold, values normal."""
    text = para.text
    c = text.find(":")
    if c < 0:
        return
    _remove_ppr_bold(para)
    _rewrite_runs(para, [(text[:c + 1], True), (text[c + 1:], False)])


def bold_skills_in_bullet(para, pattern: re.Pattern):
    """Inline-bold every matched skill/tool in a bullet paragraph."""
    text = para.text
    if not text.strip():
        return
    segs = _split_segments(text, pattern)
    if not any(bold for _, bold in segs):
        return
    # Use None (inherit style) for non-matched text — avoids explicit <w:b val="0"/>
    # conflicting with the paragraph's default non-bold style
    segs = [(t, True if b else None) for t, b in segs]
    _rewrite_runs(para, segs)


# ── Orchestrator ───────────────────────────────────────────────────────────────

def apply_bold_formatting(input_path: str, output_path: str):
    if input_path != output_path:
        shutil.copy2(input_path, output_path)
    doc = Document(output_path)

    _, skills_idxs  = _find_section(doc, SKILLS_KEYWORDS)
    _, summary_idxs = _find_section(doc, SUMMARY_KEYWORDS)
    exp_idx, _       = _find_section(doc, EXPERIENCE_KEYWORDS)
    jobs = _find_all_job_blocks(doc, exp_idx) if exp_idx >= 0 else []

    print(f"Detected  summary={len(summary_idxs)} bullets  "
          f"skills={len(skills_idxs)} lines  "
          f"jobs={len(jobs)}")

    # 1. Skills section
    print(f"\nSkills: key bold, values normal ({len(skills_idxs)} lines)...")
    for idx in skills_idxs:
        format_skills_line(doc.paragraphs[idx])

    # 2. Build term list from skills section
    skill_terms = list(EXTRA_TERMS)
    for idx in skills_idxs:
        line = doc.paragraphs[idx].text
        c = line.find(":")
        if c >= 0:
            for term in re.split(r",\s*", line[c + 1:].strip()):
                t = term.strip().strip("()")
                if t and len(t) > 1:
                    skill_terms.append(t)
    pattern = _build_pattern(skill_terms)

    # 3. Summary
    print(f"Summary: bolding skill terms in {len(summary_idxs)} bullets...")
    for idx in summary_idxs:
        bold_skills_in_bullet(doc.paragraphs[idx], pattern)

    # 4. Job bullets
    total = sum(len(b) for _, b in jobs)
    print(f"Experience: bolding skill terms in {total} bullets across {len(jobs)} jobs...")
    for _, bullet_idxs in jobs:
        for idx in bullet_idxs:
            bold_skills_in_bullet(doc.paragraphs[idx], pattern)

    doc.save(output_path)
    print(f"\nSaved -> {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Bold skill/tech terms across resume sections")
    parser.add_argument("--input",  default="resume/Revathi_Battina_Resume.docx")
    parser.add_argument("--output", default="resume/Revathi_Battina_Resume.docx")
    args = parser.parse_args()
    print(f"Reading: {args.input}\n")
    apply_bold_formatting(args.input, args.output)
    print("Done. Open in Word to verify bold formatting.")


if __name__ == "__main__":
    main()
