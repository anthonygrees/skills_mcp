#!/usr/bin/env python3
"""Validate skills against Airia Skills-over-MCP repo requirements.

Rules (invalid skills are skipped silently by the Gateway, so check here):
  - Every folder containing SKILL.md is a skill.
  - Frontmatter `name` must match the folder name.
  - Frontmatter keys limited to: name, description, license, compatibility,
    metadata, allowed-tools.
  - SKILL.md must be <= 256 KB (larger files are skipped).
  - Only the first 20 skills per server are scanned.
"""
import re
import sys
from pathlib import Path

ALLOWED_KEYS = {"name", "description", "license", "compatibility", "metadata", "allowed-tools"}
MAX_BYTES = 256 * 1024
SCAN_LIMIT = 20
ROOT = Path(__file__).resolve().parent.parent / "skills"


def parse_frontmatter(text):
    m = re.match(r"^---\r?\n(.*?)\r?\n---\r?\n", text, re.DOTALL)
    if not m:
        return None
    keys = {}
    for line in m.group(1).splitlines():
        # top-level keys only: no leading whitespace, not a list item or comment
        km = re.match(r"^([A-Za-z0-9_-]+):\s*(.*)$", line)
        if km:
            keys[km.group(1)] = km.group(2).strip().strip("\"'")
    return keys


def main():
    skills = sorted(p.parent for p in ROOT.rglob("SKILL.md"))
    errors, warnings = [], []

    for d in skills:
        rel = d.relative_to(ROOT.parent)
        f = d / "SKILL.md"
        if f.stat().st_size > MAX_BYTES:
            errors.append(f"{rel}: SKILL.md exceeds 256 KB")
        fm = parse_frontmatter(f.read_text(encoding="utf-8"))
        if fm is None:
            errors.append(f"{rel}: missing or malformed frontmatter")
            continue
        extra = set(fm) - ALLOWED_KEYS
        if extra:
            errors.append(f"{rel}: disallowed frontmatter keys: {sorted(extra)}")
        if not fm.get("name"):
            errors.append(f"{rel}: missing `name`")
        elif fm["name"] != d.name:
            errors.append(f"{rel}: name '{fm['name']}' != folder '{d.name}'")
        if not fm.get("description"):
            errors.append(f"{rel}: missing `description`")

    if len(skills) > SCAN_LIMIT:
        warnings.append(f"{len(skills)} skills: only {SCAN_LIMIT} per server are security-scanned")

    for w in warnings:
        print(f"WARN  {w}")
    for e in errors:
        print(f"ERROR {e}")
    print(f"{len(skills)} skill(s) checked, {len(errors)} error(s)")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
