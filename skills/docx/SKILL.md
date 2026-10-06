---
name: docx
description: "Create, read, edit, or review Word documents (.docx). Use whenever a .docx is an input or output, or the user asks for a Word document, report, memo, letter, proposal, or template with real formatting: headings, lists, tables, images, headers and footers, page numbers, a table of contents. Also covers extracting text, find-and-replace that keeps formatting, tracked changes, comments, and checking a finished document. Not for PDFs, spreadsheets, or Google Docs."
license: Apache-2.0
---

# DOCX

Work with Word files using `python-docx`. This skill is self-contained: everything it refers to is in this folder.

## Setup

```bash
pip install python-docx
```

- `python-docx` 1.2 or later is needed for comments.
- Optional: `pandoc` for a quick text or Markdown dump, including tracked changes.
- Optional: LibreOffice and Poppler for a visual check.

## Choose a path

| Task | Go to |
|------|-------|
| Read or summarize a document | [Reading](#reading) |
| Build a new document | [Creating](#creating) and `references/recipes.md` |
| Change an existing document | [Editing](#editing) |
| Tracked changes or comments | [Review mode](#review-mode-tracked-changes-and-comments) and `references/review.md` |
| Check a finished document | [Checking](#checking-required) |

## Reading

```bash
python scripts/inspect_docx.py doc.docx             # outline in document order
pandoc --track-changes=all doc.docx -t markdown     # alternative dump, shows tracked changes
```

Report what is in the document. Do not modify the file when the task is only to read it. Legacy `.doc` files must first be converted to `.docx` (for example with LibreOffice: `soffice --headless --convert-to docx file.doc`); tell the user if that tool is unavailable.

## Creating

1. **Clarify the brief.** Audience, purpose, length, required sections, and whether a company template or style applies. If none is given, choose sensible defaults and say what you chose.
2. **Outline first** for anything longer than a page or two.
3. **Set the page up explicitly** (size, margins, orientation) and **define styles once** rather than formatting each paragraph by hand. See `references/recipes.md`.
4. **Use real structure:** heading styles for headings, list styles for lists, table objects for tabular data, fields for page numbers and the table of contents.
5. **Save, then run the checks** below.

Defaults unless told otherwise:

- Page: US Letter (8.5in x 11in) for US users, A4 elsewhere. `python-docx`'s built-in template is US Letter; set the size explicitly anyway.
- Margins: 1in.
- Body: a widely available font (Arial or Calibri) at 11pt, left-aligned.
- Headings: Heading 1 to 3 in the built-in styles, so navigation and the table of contents work.

## Editing

1. Run `scripts/inspect_docx.py` on the original to learn its structure and styles.
2. **Work on a copy.** Never overwrite the user's original unless asked.
3. Match the document's existing styles. Apply an existing style by name instead of formatting directly.
4. Change only what was asked.
5. To replace text without losing formatting, edit **runs**, not `paragraph.text`. Setting `paragraph.text` drops run formatting. Text split across runs needs the runs joined first (see `references/recipes.md`).
6. Save under a new filename and run the checks below.

## Review mode: tracked changes and comments

If the user wants edits they can accept or reject, or asks for review comments, use tracked changes and comments instead of silently editing. `references/review.md` has the details; the helper script does the common cases:

```bash
python scripts/tracked_changes.py replace in.docx out.docx --old "30 days" --new "60 days" --author "Reviewer"
python scripts/tracked_changes.py accept in.docx out.docx
python scripts/tracked_changes.py reject in.docx out.docx
```

Use the author name the user gives; otherwise use a neutral name such as "Reviewer", and say which you used. Make minimal edits: change the words that change, not the whole sentence.

## Checking (required)

```bash
python scripts/inspect_docx.py output.docx --check
```

`--check` flags empty or skipped heading levels, leftover placeholder text, typed bullet characters instead of list styles, manual line breaks, images without alt text, tables whose header row does not repeat, fully empty table rows, A4 pages, and remaining tracked changes. Fix every finding or explain why it is intentional, then rerun until clean.

Then verify the content: read the outline for order, spelling, numbers, and consistency.

**Optional visual check.** If LibreOffice and Poppler are installed:

```bash
soffice --headless --convert-to pdf output.docx
pdftoppm -jpeg -r 100 output.pdf page
```

Look at the pages for clipped tables, awkward page breaks, headings stranded at the bottom of a page, and wrong fonts. A table of contents built from a field shows placeholder text until Word updates it; a LibreOffice render may show it unupdated. If the tools are missing, say the visual check was skipped; do not claim the layout was visually verified.

## Rules of thumb

- Never type bullet characters or numbers by hand; use list styles.
- Never use blank paragraphs or repeated spaces for spacing; use paragraph spacing and indents.
- Never use `\n` inside a run to start a new paragraph; add a paragraph.
- Never use a table for layout or as a divider line; use paragraph borders or tab stops.
- Set widths on table cells, not only on the table.
- Keep headings in order (Heading 1, then 2, then 3).
- Give meaningful images alt text.
- Do not invent facts, figures, or citations. Mark missing information clearly and tell the user.
- Do not put credentials, internal URLs, or personal data into a document unless the user supplied them for that purpose.
