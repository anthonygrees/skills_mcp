#!/usr/bin/env python3
"""Outline a .pptx and optionally check it for common problems.

Usage:
    python inspect_pptx.py deck.pptx            # outline
    python inspect_pptx.py deck.pptx --check    # outline + lint; exit 1 if findings

Requires: pip install python-pptx
"""
import argparse
import re
import sys

from pptx import Presentation

PLACEHOLDER_PATTERNS = re.compile(
    r"lorem|ipsum|click to (add|edit)|xxxx|\bTBD\b|\bTODO\b|your (title|text|name) here|\{\{.*?\}\}",
    re.IGNORECASE,
)
MIN_FONT_PT = 12
# Rough average glyph width as a fraction of font size; deliberately conservative.
CHAR_WIDTH_FACTOR = 0.5
LINE_HEIGHT_FACTOR = 1.2


def iter_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:  # group
            yield from iter_shapes(sh.shapes)


def frame_font_size(tf, default=18.0):
    sizes = [r.font.size.pt for p in tf.paragraphs for r in p.runs if r.font.size]
    return max(sizes) if sizes else default


def estimate_overflow(sh):
    """Return True if text likely needs more room than the box provides."""
    tf = sh.text_frame
    text = tf.text.strip()
    if not text or not sh.width or not sh.height:
        return False
    size = frame_font_size(tf)
    width_pt = sh.width / 12700 - (
        ((tf.margin_left or 91440) + (tf.margin_right or 91440)) / 12700
    )
    height_pt = sh.height / 12700 - (
        ((tf.margin_top or 45720) + (tf.margin_bottom or 45720)) / 12700
    )
    chars_per_line = max(1, int(width_pt / (size * CHAR_WIDTH_FACTOR)))
    lines = 0
    for para in text.split("\n"):
        lines += max(1, -(-len(para) // chars_per_line))
    needed = lines * size * LINE_HEIGHT_FACTOR
    return needed > height_pt * 1.05


def outline(prs):
    for n, slide in enumerate(prs.slides, 1):
        title = slide.shapes.title.text_frame.text if slide.shapes.title is not None else "(no title)"
        print(f"\n--- Slide {n}: {title!r}  [layout: {slide.slide_layout.name}]")
        for sh in iter_shapes(slide.shapes):
            kind = str(sh.shape_type).split(".")[-1].split(" ")[0] if sh.shape_type else "UNKNOWN"
            if sh.has_text_frame and sh.text_frame.text.strip():
                text = sh.text_frame.text.strip().replace("\n", " / ")
                print(f"  [{kind}] {sh.name}: {text[:140]}")
            elif getattr(sh, "has_chart", False) and sh.has_chart:
                print(f"  [CHART] {sh.name}: {sh.chart.chart_type}")
            elif getattr(sh, "has_table", False) and sh.has_table:
                t = sh.table
                print(f"  [TABLE] {sh.name}: {len(t.rows)}x{len(t.columns)}")
            elif sh.shape_type == 13:
                print(f"  [PICTURE] {sh.name}")
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            print(f"  [NOTES] {slide.notes_slide.notes_text_frame.text.strip()[:140]}")


def check(prs):
    findings = []
    sw, sh_ = prs.slide_width, prs.slide_height
    for n, slide in enumerate(prs.slides, 1):
        if slide.shapes.title is None or not slide.shapes.title.text_frame.text.strip():
            findings.append(f"slide {n}: no title")
        for sh in iter_shapes(slide.shapes):
            where = f"slide {n} '{sh.name}'"
            if sh.is_placeholder and sh.has_text_frame and not sh.text_frame.text.strip():
                findings.append(f"{where}: empty placeholder")
            if sh.has_text_frame:
                text = sh.text_frame.text
                m = PLACEHOLDER_PATTERNS.search(text)
                if m:
                    findings.append(f"{where}: leftover placeholder text '{m.group(0)}'")
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.size and r.font.size.pt < MIN_FONT_PT and r.text.strip():
                            findings.append(f"{where}: font {r.font.size.pt:g}pt < {MIN_FONT_PT}pt")
                            break
                if estimate_overflow(sh):
                    findings.append(f"{where}: text may overflow its box")
            if None not in (sh.left, sh.top, sh.width, sh.height):
                if sh.left < 0 or sh.top < 0 or sh.left + sh.width > sw or sh.top + sh.height > sh_:
                    findings.append(f"{where}: extends beyond the slide")
            if sh.shape_type == 13:
                descr = sh._element.xpath(".//p:cNvPr/@descr")
                if not descr or not descr[0].strip():
                    findings.append(f"{where}: picture has no alt text")
    return findings


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--check", action="store_true", help="lint the deck; exit 1 on findings")
    args = ap.parse_args()

    prs = Presentation(args.path)
    print(f"{args.path}: {len(prs.slides)} slides, "
          f"{prs.slide_width / 914400:.2f}in x {prs.slide_height / 914400:.2f}in")
    outline(prs)

    if args.check:
        findings = check(prs)
        print("\n=== Check ===")
        for f in findings:
            print(f"  - {f}")
        print(f"{len(findings)} finding(s)")
        sys.exit(1 if findings else 0)


if __name__ == "__main__":
    main()
