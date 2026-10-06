# Review mode: tracked changes and comments

Use review mode when the user wants edits they can accept or reject, wants to see what changed, or asks for feedback on a document. If they want a clean final copy, edit directly instead.

Rules:

- Keep the original untouched; save to a new filename.
- Use the author name the user gives. If none, use a neutral name such as "Reviewer" and say so.
- Make **minimal** changes: mark only the words that change.
- Keep the original run formatting on inserted and deleted text.

## Comments (python-docx 1.2+)

```python
from docx import Document

doc = Document("input.docx")
para = doc.paragraphs[3]
comment = doc.add_comment(
    runs=para.runs,                 # a run or a list of runs the comment is anchored to
    text="Can we cite a source for this figure?",
    author="Reviewer",
    initials="R",
)
comment.add_paragraph("Second paragraph of the same comment.")   # optional
doc.save("reviewed.docx")

for c in doc.comments:              # reading
    print(c.comment_id, c.author, c.text)
```

Limits: `python-docx` creates top-level comments only. It has no API for threaded replies or for marking a comment resolved. If the user needs replies, say so and add a new comment that quotes the one it answers.

Comments are anchored to runs. Text inside a table cell works the same way (`cell.paragraphs[0].runs`).

## Tracked replacements (helper script)

```bash
python scripts/tracked_changes.py replace in.docx out.docx --old "30 days" --new "60 days" --author "Reviewer"
python scripts/tracked_changes.py replace in.docx out.docx --old "colour" --new "color" --all
python scripts/tracked_changes.py accept  in.docx clean.docx
python scripts/tracked_changes.py reject  in.docx original_again.docx
```

From Python:

```python
import sys; sys.path.insert(0, "scripts")
from tracked_changes import tracked_replace, accept_all, reject_all

doc = Document("in.docx")
n = tracked_replace(doc, "30 days", "60 days", author="Reviewer", count=1)   # count=0 replaces all
if n == 0:
    ...   # the phrase is not inside a single run; see below
doc.save("out.docx")
```

What it does: splits the matching run into the text before the match, a tracked deletion of the old text, a tracked insertion of the new text, and the text after. Run formatting is copied to each piece.

Limits:

- A match must sit inside one run. If Word split the phrase across runs, use a shorter fragment that lies in one run, or merge identical-format runs first (`merge_adjacent_runs` in `recipes.md`).
- It works on the document body, not headers, footers, footnotes, or text boxes.
- `accept` and `reject` act on the body only. Accepting removes deletions and keeps insertions; rejecting does the opposite. Tracked formatting changes and moves are not handled.
- Pass an empty `--new ""` to track a pure deletion.

## What the XML looks like

Needed only for cases the helper does not cover. Edit the XML through `python-docx` elements (`paragraph._p`), or unpack the `.docx` (it is a ZIP) and edit `word/document.xml`, then rezip with `[Content_Types].xml` first and the same relative paths.

Insertion:

```xml
<w:ins w:id="1" w:author="Reviewer" w:date="2025-01-01T00:00:00Z">
  <w:r><w:t>inserted text</w:t></w:r>
</w:ins>
```

Deletion (note `w:delText`, not `w:t`):

```xml
<w:del w:id="2" w:author="Reviewer" w:date="2025-01-01T00:00:00Z">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
```

Minimal change from "30 days" to "60 days":

```xml
<w:r><w:t xml:space="preserve">The term is </w:t></w:r>
<w:del w:id="1" w:author="Reviewer" w:date="..."><w:r><w:delText>30</w:delText></w:r></w:del>
<w:ins w:id="2" w:author="Reviewer" w:date="..."><w:r><w:t>60</w:t></w:r></w:ins>
<w:r><w:t xml:space="preserve"> days.</w:t></w:r>
```

Rules for hand-written XML:

- `w:id` values must be unique across all tracked changes in the document.
- Put `xml:space="preserve"` on any `w:t` or `w:delText` with leading or trailing spaces.
- Insert `w:ins` and `w:del` as siblings of runs inside the paragraph, not inside a run.
- Copy the original run's `<w:rPr>` into the new runs to keep bold, size, and font.
- To delete a whole paragraph, also add `<w:del .../>` inside `<w:pPr><w:rPr>` so accepting the change does not leave an empty paragraph.
- To reject someone else's insertion, nest your `w:del` inside their `w:ins`. To restore someone else's deletion, add a new `w:ins` after it; do not edit their deletion.
- In `<w:pPr>` the child order matters: `pStyle`, `numPr`, `spacing`, `ind`, `jc`, `rPr` last.

## Verifying

```bash
pandoc --track-changes=all out.docx -t markdown     # shows insertions, deletions, and comments
python scripts/inspect_docx.py out.docx             # counts of tracked changes and comments
```

`inspect_docx.py --check` reports remaining tracked changes, which is expected when you are producing a reviewed copy and a finding when the user wants a clean one.

If Word reports the file as corrupt, the cause is almost always hand-written XML: duplicate `w:id`, a `w:t` inside a `w:del`, a tracked-change element inside a run, or an invalid child order. Revert to the output before the hand edit and redo it with the helper.
