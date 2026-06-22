"""
Sends the tailored resume as an email attachment via Gmail SMTP.
Recipients can be passed directly or fall back to EMAIL_RECIPIENTS in .env.
The email body includes a detailed change log: which section, which bullet,
and which skills were added at each location.
"""

import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv

load_dotenv()


def _format_change_log(change_log: list[dict]) -> str:
    """
    Render a human-readable change report grouped by section.

    change_log items: {"section": str, "bullet": int, "skills_added": [str], "text": str}
    """
    if not change_log:
        return "  (No changes detected — all sections matched original content.)"

    # Group by section
    sections: dict[str, list[dict]] = {}
    for entry in change_log:
        sections.setdefault(entry["section"], []).append(entry)

    lines = []
    for section, entries in sections.items():
        lines.append(f"  [{section}]")
        for e in entries:
            bullet_label = f"Bullet {e['bullet']}"
            skills = e.get("skills_added", [])
            skill_note = f"  >> added: {', '.join(skills)}" if skills else ""
            # Truncate long bullet text for readability
            preview = e["text"][:130].rstrip()
            if len(e["text"]) > 130:
                preview += "..."
            lines.append(f"    {bullet_label}:{skill_note}")
            lines.append(f"      \"{preview}\"")
        lines.append("")  # blank line between sections

    return "\n".join(lines).rstrip()


def send_resume_email(
    resume_path: str,
    job_title: str,
    skills: list[str],
    change_log: list[dict],
    to_addresses: list[str] | None = None,
) -> None:
    """
    Email the updated resume to `to_addresses`.
    Falls back to EMAIL_RECIPIENTS env var if to_addresses is None or empty.
    Raises RuntimeError if sender credentials are missing.
    """
    sender   = os.getenv("EMAIL_SENDER", "").strip()
    password = os.getenv("EMAIL_APP_PASSWORD", "").strip()

    if not sender or not password:
        raise RuntimeError("EMAIL_SENDER and EMAIL_APP_PASSWORD must be set in .env")

    # Resolve recipients — explicit arg takes priority over env var
    if to_addresses:
        recipients = [r.strip() for r in to_addresses if r.strip()]
    else:
        raw = os.getenv("EMAIL_RECIPIENTS", "").strip()
        recipients = [r.strip() for r in raw.split(",") if r.strip()]

    if not recipients:
        raise RuntimeError(
            "No recipients found. Pass --email-to or set EMAIL_RECIPIENTS in .env"
        )

    # ── Build message ────────────────────────────────────────────
    msg = MIMEMultipart()
    msg["From"]    = sender
    msg["To"]      = ", ".join(recipients)
    msg["Subject"] = f"Tailored Resume — {job_title}"

    skills_list  = "\n".join(f"  • {s}" for s in skills)
    changes_text = _format_change_log(change_log)

    body = f"""\
Hi,

Please find attached the updated resume tailored for: {job_title}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SKILLS INCORPORATED FROM RECRUITER REQUIREMENTS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{skills_list}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DETAILED CHANGE LOG — WHERE EACH SKILL WAS ADDED:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{changes_text}

Sections updated:
  ✓ Professional Summary
  ✓ Technical Skills
  ✓ Toyota | Plano, TX  (Job 1 — most recent)
  ✓ Equifax | Trivandrum, India  (Job 2)

Everything else (contact info, other jobs, education) is untouched.

Best regards,
Resume Updater Bot
"""

    msg.attach(MIMEText(body, "plain"))

    # ── Attach the .docx ────────────────────────────────────────
    filename = os.path.basename(resume_path)
    with open(resume_path, "rb") as fh:
        part = MIMEBase(
            "application",
            "vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        part.set_payload(fh.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    msg.attach(part)

    # ── Send via Gmail SSL ───────────────────────────────────────
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, password)
        server.sendmail(sender, recipients, msg.as_string())

    print(f"  Email sent → {', '.join(recipients)}")
