#!/usr/bin/env python3
"""Outline a .docx and optionally check it for common problems.

Usage:
    python inspect_docx.py doc.docx            # outline in document order
    python inspect_docx.py doc.docx --check    # outline + lint; exit 1 on findings

Requires: pip install python-docx
"""
import argparse
import re
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

PLACEHOLDER = re.compile(
    r"lorem|ipsum|xxxx|\bTBD\b|\bTODO\b|\[(insert|your|name|date|company)[^\]]*\]|\{\{.*?\}\}",
    re.IGNORECASE,
)
FAKE_BULLET = re.compile(r"^\s*[•●◦‣⁃\-\*]\s+\S")


def iter_blocks(doc):
    """Yield paragraphs and tables in document order."""
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def para_text(p):
    """Paragraph text as it reads with tracked changes shown as inserted.

    python-docx's Paragraph.text skips runs inside <w:ins>; read w:t directly
    (deleted text lives in w:delText and is correctly excluded).
    """
    return "".join(t.text or "" for t in p._p.iter(qn("w:t")))


def heading_level(p):
    name = p.style.name if p.style is not None else ""
    if name == "Title":
        return 0
    m = re.match(r"Heading (\d)", name)
    return int(m.group(1)) if m else None


def count_tracked(doc):
    body = doc.element.body
    return len(body.xpath(".//w:ins")), len(body.xpath(".//w:del"))


def outline(doc):
    ins, dele = count_tracked(doc)
    s = doc.sections[0]
    print(f"{len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables, "
          f"{len(doc.inline_shapes)} inline images, {len(doc.sections)} section(s)")
    print(f"page: {s.page_width.inches:.2f}in x {s.page_height.inches:.2f}in, "
          f"margins L{s.left_margin.inches:.2f} R{s.right_margin.inches:.2f} "
          f"T{s.top_margin.inches:.2f} B{s.bottom_margin.inches:.2f}")
    print(f"tracked changes: {ins} insertions, {dele} deletions; comments: "
          f"{len(list(doc.comments)) if hasattr(doc, 'comments') else 'n/a'}")
    print()
    for block in iter_blocks(doc):
        if isinstance(block, Table):
            print(f"[TABLE {len(block.rows)}x{len(block.columns)}] "
                  f"first row: {[c.text.strip()[:20] for c in block.rows[0].cells]}")
            continue
        text = para_text(block).strip()
        lvl = heading_level(block)
        if lvl is not None and text:
            print(f"{'  ' * lvl}{'#' * (lvl + 1)} {text}")
        elif text:
            print(f"    {text[:110]}")


def check(doc):
    findings = []
    last_level = 0
    seen_heading = False
    for n, block in enumerate(iter_blocks(doc), 1):
        if isinstance(block, Table):
            continue
        text = para_text(block).strip()
        lvl = heading_level(block)
        if lvl is not None:
            seen_heading = True
            if not text:
                findings.append(f"block {n}: empty heading")
            if lvl > last_level + 1 and lvl > 1:
                findings.append(f"block {n}: heading level jumps to {lvl} after {last_level}: '{text[:40]}'")
            last_level = lvl
        if text and PLACEHOLDER.search(text):
            findings.append(f"block {n}: leftover placeholder text: '{text[:60]}'")
        if text and FAKE_BULLET.match(text) and "List" not in (block.style.name or ""):
            findings.append(f"block {n}: typed bullet character instead of a list style: '{text[:40]}'")
        if "\n" in para_text(block):
            findings.append(f"block {n}: manual line breaks in a paragraph (use separate paragraphs)")
    if not seen_heading and len(doc.paragraphs) > 15:
        findings.append("no headings used in a long document (navigation and TOC need heading styles)")

    for i, shape in enumerate(doc.inline_shapes, 1):
        descr = shape._inline.docPr.get("descr") or ""
        if not descr.strip():
            findings.append(f"image {i}: no alt text")

    for i, t in enumerate(doc.tables, 1):
        if len(t.rows) > 1 and not t.rows[0]._tr.xpath("./w:trPr/w:tblHeader"):
            findings.append(f"table {i}: header row is not marked to repeat across pages")
        for row in t.rows:
            if any(not c.text.strip() for c in row.cells) and len(row.cells) == len(t.columns):
                if all(not c.text.strip() for c in row.cells):
                    findings.append(f"table {i}: fully empty row")
                    break

    for i, s in enumerate(doc.sections, 1):
        if s.page_width and s.page_height:
            w, h = s.page_width.inches, s.page_height.inches
            if abs(min(w, h) - 8.27) < 0.05 and abs(max(w, h) - 11.69) < 0.05:
                findings.append(f"section {i}: A4 page size; confirm that is intended")
    ins, dele = count_tracked(doc)
    if ins or dele:
        findings.append(f"{ins} insertions and {dele} deletions are still tracked (accept or reject if a clean copy is expected)")
    return findings


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--check", action="store_true", help="lint; exit 1 on findings")
    args = ap.parse_args()
    doc = Document(args.path)
    outline(doc)
    if args.check:
        findings = check(doc)
        print("\n=== Check ===")
        for f in findings:
            print(f"  - {f}")
        print(f"{len(findings)} finding(s)")
        sys.exit(1 if findings else 0)


if __name__ == "__main__":
    main()
