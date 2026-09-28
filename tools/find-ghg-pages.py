#!/usr/bin/env python3
"""Locate pages in PDF reports that likely contain Scope 1/2 GHG emissions figures.

Usage: python tools/find-ghg-pages.py <file.pdf> [more.pdf ...]

Wraps `pdftotext -layout` (local, free). For each PDF prints candidate pages with
matched snippets and a verdict:
  DATA          - scope keywords AND numeric values with CO2e-style units on the page
  KEYWORDS_ONLY - scope keywords found but no numeric/unit context (image table or
                  targets-only text) -> visually Read those pages to confirm
  NONE          - no scope mentions in the text layer (wrong document, or scanned PDF)
"""
import re
import subprocess
import sys

# Windows consoles default to cp1252; PDF text routinely contains characters outside
# it. Without this, printing a snippet can raise UnicodeEncodeError MID-SCAN, killing
# the run after only the early pages — a silently incomplete scan (this happened: a
# subsidiary breakout on p.35 went unflagged because the scan died at p.17).
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

SCOPE_RE = re.compile(r"scope\s*[12]", re.IGNORECASE)
# a number (1,234 / 1.234 / 12 345 / 0.1) near a CO2e-style unit on the same page
UNIT_RE = re.compile(
    r"(tCO2e?|tCO₂e?|CO2e|CO₂e|CO2-?eq|CO₂-?eq|tonnes?\s+(of\s+)?CO2|metric\s+tons?|ktCO2e?|MtCO2e?|million\s+tons?)",
    re.IGNORECASE,
)
NUM_RE = re.compile(r"\d[\d,.  ]{0,14}\d|\d")


def snippet(text: str, m: re.Match, radius: int = 90) -> str:
    s = max(0, m.start() - radius)
    e = min(len(text), m.end() + radius)
    return re.sub(r"\s+", " ", text[s:e]).strip()


def scan(pdf_path: str) -> None:
    try:
        raw = subprocess.run(
            ["pdftotext", "-layout", pdf_path, "-"],
            capture_output=True, timeout=120,
        )
    except FileNotFoundError:
        sys.exit("ERROR: pdftotext not found on PATH")
    if raw.returncode != 0:
        print(f"=== {pdf_path}\n  ERROR: pdftotext failed: {raw.stderr.decode(errors='replace')[:200]}")
        return

    pages = raw.stdout.decode("utf-8", errors="replace").split("\f")
    data_pages, keyword_pages = [], []
    print(f"=== {pdf_path} ({len(pages)} pages)")

    for i, page in enumerate(pages, start=1):
        matches = list(SCOPE_RE.finditer(page))
        if not matches:
            continue
        has_units = bool(UNIT_RE.search(page)) and bool(NUM_RE.search(page))
        (data_pages if has_units else keyword_pages).append(i)
        label = "DATA?" if has_units else "kw   "
        for m in matches[:3]:
            print(f"  p.{i:<4} [{label}] {snippet(page, m)[:170]}")

    if data_pages:
        print(f"  VERDICT: DATA  pages: {', '.join(map(str, data_pages))}")
    elif keyword_pages:
        print(f"  VERDICT: KEYWORDS_ONLY  pages: {', '.join(map(str, keyword_pages))}"
              "  -> visually Read these pages")
    else:
        note = " (no text layer - scanned PDF?)" if not any(p.strip() for p in pages) else ""
        print(f"  VERDICT: NONE{note}")
    print()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    for path in sys.argv[1:]:
        scan(path)
