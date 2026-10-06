#!/usr/bin/env python3
"""Tracked-change helpers for .docx files (python-docx has no API for these).

CLI:
    python tracked_changes.py replace in.docx out.docx --old "30 days" --new "60 days" [--author NAME]
    python tracked_changes.py accept  in.docx out.docx
    python tracked_changes.py reject  in.docx out.docx

Library:
    from tracked_changes import tracked_replace, accept_all, reject_all

Limits: tracked_replace matches text inside a single run. If Word split the phrase
across runs, widen --old to a single-run fragment, or join the runs first.
Accept/reject act on the main body only (not headers, footers, or comments).

Requires: pip install python-docx
"""
import argparse
import copy
import datetime
import itertools

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

_ids = itertools.count(9000)


def _now():
    return datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _mark(tag, author):
    el = OxmlElement(tag)
    el.set(qn("w:id"), str(next(_ids)))
    el.set(qn("w:author"), author)
    el.set(qn("w:date"), _now())
    return el


def _run_with_text(template_r, text, deleted=False):
    """Copy a run (keeping its formatting) and give it new text."""
    r = copy.deepcopy(template_r)
    for child in list(r):
        if child.tag != qn("w:rPr"):
            r.remove(child)
    t = OxmlElement("w:delText" if deleted else "w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def _find_match(doc, old):
    for r in doc.element.body.iter(qn("w:r")):
        if r.getparent().tag in (qn("w:del"), qn("w:ins")):
            continue
        t_nodes = r.findall(qn("w:t"))
        if len(t_nodes) == 1 and old in (t_nodes[0].text or ""):
            return r, t_nodes[0].text
    return None, None


def tracked_replace(doc, old, new, author="Reviewer", count=1):
    """Replace `old` with `new` as a tracked deletion + insertion.

    Returns the number of replacements made (at most `count`; use count=0 for all).
    Inserted text is never rescanned, so `new` may contain `old`.
    """
    done = 0
    while not count or done < count:
        r, text = _find_match(doc, old)
        if r is None:
            break
        before, _, after = text.partition(old)
        parent = r.getparent()
        idx = parent.index(r)
        pieces = []
        if before:
            pieces.append(_run_with_text(r, before))
        d = _mark("w:del", author)
        d.append(_run_with_text(r, old, deleted=True))
        pieces.append(d)
        if new:
            i = _mark("w:ins", author)
            i.append(_run_with_text(r, new))
            pieces.append(i)
        if after:
            pieces.append(_run_with_text(r, after))
        parent.remove(r)
        for off, el in enumerate(pieces):
            parent.insert(idx + off, el)
        done += 1
    return done


def _unwrap(el):
    parent = el.getparent()
    idx = parent.index(el)
    for off, child in enumerate(list(el)):
        parent.insert(idx + off, child)
    parent.remove(el)


def _drop_empty_paragraphs(body):
    # Paragraphs whose mark was tracked-deleted and that are now empty.
    for p in list(body.iter(qn("w:p"))):
        ppr_del = p.xpath("./w:pPr/w:rPr/w:del")
        if ppr_del:
            if not p.xpath(".//w:t | .//w:drawing"):
                p.getparent().remove(p)
            else:
                for d in ppr_del:
                    d.getparent().remove(d)


def accept_all(doc):
    body = doc.element.body
    _drop_empty_paragraphs(body)
    for d in body.xpath(".//w:del[not(parent::w:rPr)]"):
        d.getparent().remove(d)
    for i in body.xpath(".//w:ins[not(parent::w:rPr)]"):
        _unwrap(i)


def reject_all(doc):
    body = doc.element.body
    for i in body.xpath(".//w:ins[not(parent::w:rPr)]"):
        i.getparent().remove(i)
    for d in body.xpath(".//w:del[not(parent::w:rPr)]"):
        for dt in list(d.iter(qn("w:delText"))):
            dt.tag = qn("w:t")
        _unwrap(d)
    for m in body.xpath(".//w:pPr/w:rPr/w:ins | .//w:pPr/w:rPr/w:del"):
        m.getparent().remove(m)


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("replace")
    r.add_argument("src"); r.add_argument("dst")
    r.add_argument("--old", required=True); r.add_argument("--new", required=True)
    r.add_argument("--author", default="Reviewer")
    r.add_argument("--all", action="store_true", help="replace every occurrence")
    for name in ("accept", "reject"):
        s = sub.add_parser(name)
        s.add_argument("src"); s.add_argument("dst")
    a = ap.parse_args()

    doc = Document(a.src)
    if a.cmd == "replace":
        n = tracked_replace(doc, a.old, a.new, a.author, count=0 if a.all else 1)
        print(f"{n} replacement(s)")
        if n == 0:
            raise SystemExit("no match inside a single run; nothing changed")
    elif a.cmd == "accept":
        accept_all(doc)
    else:
        reject_all(doc)
    doc.save(a.dst)


if __name__ == "__main__":
    main()
