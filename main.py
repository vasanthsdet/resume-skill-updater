"""
Resume Skill Updater — main entry point.

Flow:
  1. Read recruiter requirements (text or file)
  2. Extract skills using Claude Haiku
  3. Update Summary, Skills, Job1, Job2 in the .docx using Claude Sonnet
  4. Optionally email the updated resume with a full change log

Usage examples:
  # Paste recruiter text inline, email to custom recipients
  python main.py --requirements "Playwright, TypeScript, Azure DevOps..." \\
                 --email-to recruiter@company.com you@gmail.com --send-email

  # Point to a requirements file
  python main.py --requirements-file recruiter_jd.txt --job-title "SDET" --send-email

  # Skip extraction and supply skills directly
  python main.py --skills "Playwright" "TypeScript" "DataDog" --send-email
"""

import argparse
import os
import re
import sys
from dotenv import load_dotenv

from skill_extractor import extract_skills
from resume_updater import update_resume
from email_sender import send_resume_email

load_dotenv()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Tailor a resume to recruiter requirements and optionally email it",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    # ── Input: recruiter requirements OR direct skills ────────────
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument(
        "--requirements",
        metavar="TEXT",
        help="Recruiter requirements pasted inline as a string",
    )
    source.add_argument(
        "--requirements-file",
        metavar="FILE",
        help="Path to a .txt file containing recruiter requirements",
    )
    source.add_argument(
        "--skills",
        nargs="+",
        metavar="SKILL",
        help='Skip extraction — supply skills directly: --skills "Playwright" "TypeScript"',
    )

    # ── Resume options ────────────────────────────────────────────
    parser.add_argument(
        "--job-title",
        default="QA_Engineer",
        help="Job title used in the output filename and email subject",
    )
    parser.add_argument(
        "--base",
        default="resume/base_resume.docx",
        help="Path to the base resume .docx (default: resume/base_resume.docx)",
    )
    parser.add_argument(
        "--output-dir",
        default="tailored_resumes",
        help="Directory to save the updated resume (default: tailored_resumes)",
    )

    # ── Email options ─────────────────────────────────────────────
    parser.add_argument(
        "--send-email",
        action="store_true",
        help="Email the updated resume after generation",
    )
    parser.add_argument(
        "--email-to",
        nargs="+",
        metavar="ADDRESS",
        help=(
            "Email recipient(s) — overrides EMAIL_RECIPIENTS in .env. "
            "e.g. --email-to hr@company.com you@gmail.com"
        ),
    )

    args = parser.parse_args()

    # ── Step 1: Resolve skills ────────────────────────────────────
    if args.skills:
        skills = args.skills
        print(f"[1/3] Using {len(skills)} provided skills: {', '.join(skills)}")
    else:
        if args.requirements_file:
            with open(args.requirements_file, encoding="utf-8") as f:
                recruiter_text = f.read()
            print(f"[1/3] Read requirements from: {args.requirements_file}")
        else:
            recruiter_text = args.requirements
            print("[1/3] Using inline requirements text")

        print("[2/3] Extracting skills with Claude Haiku...")
        skills = extract_skills(recruiter_text)
        print(f"      Found {len(skills)} skills: {', '.join(skills)}")

    # ── Step 2: Update resume ─────────────────────────────────────
    safe_title  = re.sub(r"[^\w\s-]", "", args.job_title).strip().replace(" ", "_")[:50]
    output_path = os.path.join(args.output_dir, f"Resume_{safe_title}.docx")
    os.makedirs(args.output_dir, exist_ok=True)

    step = "3" if args.skills else "3"
    print(f"[{step}/3] Updating resume → {output_path}")
    output_path, change_log = update_resume(args.base, skills, output_path)
    print(f"      Saved: {output_path}")
    print(f"      {len(change_log)} bullets changed across all sections")

    # ── Step 3: Email ─────────────────────────────────────────────
    if args.send_email:
        print("[4/4] Sending email...")
        send_resume_email(
            resume_path  = output_path,
            job_title    = args.job_title,
            skills       = skills,
            change_log   = change_log,
            to_addresses = args.email_to,   # None falls back to .env
        )

    print("\nDone.")


if __name__ == "__main__":
    main()
