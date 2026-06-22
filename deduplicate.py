"""
Scans every job section in the resume and removes duplicate bullet points.

Strategy:
  - Process jobs in order (Toyota → Equifax → UST → Availity → Sana Technos)
  - Keep the FIRST occurrence of a bullet; delete all later copies
  - Also catches near-duplicates (>82% word overlap via Jaccard similarity)
  - Blank separators / company headers / education are never touched

Usage:
  python deduplicate.py                            # reads & overwrites resume/base_resume.docx
  python deduplicate.py --input in.docx --output out.docx
  python deduplicate.py --dry-run                  # shows what would be removed, no file change
"""

import re
import argparse
from docx import Document

EXPERIENCE_KEYWORDS  = {"professional experience", "work experience", "employment history"}
SIMILARITY_THRESHOLD = 0.82   # Jaccard word-overlap ratio to flag near-duplicates


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip().lower())


def _jaccard(a: str, b: str) -> float:
    wa, wb = set(a.split()), set(b.split())
    if not wa or not wb:
        return 0.0
    return len(wa & wb) / len(wa | wb)


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


def _find_experience_heading(doc: Document) -> int:
    for i, para in enumerate(doc.paragraphs):
        lower = para.text.strip().lower()
        if any(kw in lower for kw in EXPERIENCE_KEYWORDS) and _is_heading(para):
            return i
    return -1


def _find_all_job_blocks(doc: Document, exp_heading_idx: int) -> list[tuple[str, list[int]]]:
    """
    Return [(job_label, [bullet_para_idxs]), ...] for every job found
    after the experience section heading.
    """
    paras  = doc.paragraphs
    jobs: list[tuple[str, list[int]]] = []
    label  = ""
    bullets: list[int] = []
    in_job = False
    i = exp_heading_idx + 1

    while i < len(paras):
        text = paras[i].text.strip()
        if not text:
            i += 1
            continue

        lower = text.lower()
        if lower.rstrip().endswith(":") and any(
            kw in lower for kw in {"education", "certification", "award", "publication"}
        ):
            break

        # "Company | Location   Date" — pipe identifies a company header line
        if " | " in text:
            if in_job:
                jobs.append((label, bullets[:]))
            label   = text
            bullets = []
            in_job  = True
            i += 1
            # Consume the job-title line (short, no pipe, no year)
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


def _remove_para(para) -> None:
    elem = para._element
    elem.getparent().remove(elem)


# ---------------------------------------------------------------------------
# Core deduplication
# ---------------------------------------------------------------------------

def deduplicate(input_path: str, output_path: str, dry_run: bool = False) -> int:
    """
    Remove duplicate bullets across all job blocks.
    Returns the number of paragraphs removed.
    """
    doc   = Document(input_path)
    paras = doc.paragraphs

    exp_idx = _find_experience_heading(doc)
    if exp_idx < 0:
        print("  [dedup] Experience section not found — nothing to do.")
        if not dry_run:
            doc.save(output_path)
        return 0

    jobs = _find_all_job_blocks(doc, exp_idx)
    print(f"  [dedup] {len(jobs)} job block(s) found")

    seen_exact: set[str]  = set()
    seen_list:  list[str] = []     # kept for Jaccard comparison
    to_remove:  list[int] = []

    for job_label, bullet_idxs in jobs:
        removed_here: list[str] = []

        for idx in bullet_idxs:
            text = paras[idx].text.strip()
            if not text:
                continue
            norm = _normalize(text)

            # Exact duplicate
            if norm in seen_exact:
                to_remove.append(idx)
                removed_here.append(text[:80])
                continue

            # Near-duplicate
            if any(_jaccard(norm, s) >= SIMILARITY_THRESHOLD for s in seen_list):
                to_remove.append(idx)
                removed_here.append(text[:80])
                continue

            # First occurrence — mark as seen
            seen_exact.add(norm)
            seen_list.append(norm)

        if removed_here:
            short_label = job_label.split("|")[0].strip()[:30]
            print(f"  [dedup] {short_label}: {len(removed_here)} duplicate(s) removed")
            for t in removed_here:
                print(f"          - {t}...")

    if not to_remove:
        print("  [dedup] No duplicates found.")
        if not dry_run:
            doc.save(output_path)
        return 0

    if dry_run:
        print(f"\n  [dry-run] Would remove {len(to_remove)} bullet(s). No file written.")
        return len(to_remove)

    # Collect paragraph objects before removal (indices shift after each delete)
    paras_to_remove = [doc.paragraphs[i] for i in sorted(set(to_remove))]

    for para in paras_to_remove:
        _remove_para(para)

    doc.save(output_path)
    print(f"\n  [dedup] Done — {len(to_remove)} duplicate bullet(s) removed.")
    print(f"  [dedup] Saved -> {output_path}")
    return len(to_remove)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Remove duplicate resume bullet points")
    parser.add_argument("--input",   default="resume/base_resume.docx", help="Input .docx")
    parser.add_argument("--output",  default="resume/base_resume.docx", help="Output .docx (default: overwrite input)")
    parser.add_argument("--dry-run", action="store_true",               help="Preview changes without writing")
    args = parser.parse_args()

    print(f"Reading: {args.input}")
    removed = deduplicate(args.input, args.output, dry_run=args.dry_run)
    if not args.dry_run:
        print(f"Saved:   {args.output}  ({removed} bullets removed)")


if __name__ == "__main__":
    main()
