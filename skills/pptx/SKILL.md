---
name: pptx
description: "Create, read, edit, or restructure PowerPoint (.pptx) files. Use whenever a .pptx is an input or output, or the user mentions a deck, slides, presentation, pitch deck, speaker notes, or a template. Covers building decks from scratch, extracting text and notes, editing existing decks, filling templates, merging or reordering slides, and checking the result for layout problems."
license: Apache-2.0
---

# PPTX

Work with PowerPoint files using `python-pptx`. This skill is self-contained: everything it refers to is in this folder.

## Setup

```bash
pip install python-pptx "markitdown[pptx]"
```

- `python-pptx` creates and edits decks.
- `markitdown` is optional; it gives a quick Markdown view of a deck.

## Choose a path

| Task | Go to |
|------|-------|
| Read or summarize a deck | [Reading](#reading) |
| Build a new deck | [Creating](#creating) and `references/design.md` |
| Change an existing deck or fill a template | [Editing](#editing) and `references/recipes.md` |
| Check a finished deck | [Checking](#checking-required) |

## Reading

```bash
python scripts/inspect_pptx.py deck.pptx            # outline: per-slide layout, titles, text, notes
python -m markitdown deck.pptx                      # alternative Markdown dump
```

Report what is in the deck (slide count, titles, key points, notes). Do not modify the file when the task is only to read it.

## Creating

1. **Clarify the brief.** Audience, purpose, length, and any required content. If the user gave none, pick a sensible default and say what you chose.
2. **Outline first.** One message per slide, written as a sentence (a claim, not a topic label). Show the outline for approval when the deck is long or the brief is vague.
3. **Pick a design system** before writing code: palette, font pair, and one repeated visual motif. Follow `references/design.md`.
4. **Build with layouts.** Use the layouts and placeholders in the presentation's slide master so titles, outline view, and accessibility work. Fall back to free shapes only for custom visuals. Code patterns are in `references/recipes.md`.
5. **Add speaker notes** when the deck will be presented.
6. **Save, then run the checks** below.

Use 16:9 (`13.333in x 7.5in`) unless told otherwise. `Presentation()` defaults to 4:3, so set the size explicitly.

## Editing

1. Run `scripts/inspect_pptx.py` on the original to learn its layouts, placeholders, and existing content.
2. **Work on a copy.** Never overwrite the user's original unless they ask.
3. Reuse the deck's existing layouts, fonts, and colors. Match what is already there instead of imposing a new style.
4. Change only what was asked. Preserve formatting by editing runs, not by replacing whole text frames (see `references/recipes.md`).
5. For templates: map each piece of content to a layout, fill the placeholders, and delete any slides or placeholders left unused. No sample text may remain.
6. Save under a new filename and run the checks below.

## Checking (required)

Treat the first build as a draft that probably has problems.

```bash
python scripts/inspect_pptx.py output.pptx --check
```

`--check` flags empty placeholders, leftover template text, shapes off the slide, text likely to overflow its box, fonts below 12pt, and pictures without alt text. Fix every finding or explain why it is intentional, then rerun until clean.

Then verify the content: read the outline and confirm slide order, spelling, numbers, and that every slide supports its message.

**Optional visual check.** If LibreOffice and Poppler are installed, render and look at the slides:

```bash
soffice --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 100 output.pdf slide
```

Inspect the images for overlap, clipping, low contrast, and uneven spacing. If the tools are missing, say that the visual check was skipped; do not claim the layout was visually verified.

## Rules of thumb

- One idea per slide. Titles state the takeaway.
- No text-only slides when a chart, diagram, or image would carry the point better.
- Left-align body text. Keep at least 0.5in margins.
- Body text 14pt or larger; titles 32pt or larger.
- Never invent statistics, quotes, or sources. Mark placeholders clearly if data is missing, and tell the user.
- Do not embed credentials, internal URLs, or personal data in a deck unless the user supplied them for that purpose.
