# python-docx recipes

All snippets were run against `python-docx` 1.2. They assume:

```python
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
```

Units: `Inches()`, `Pt()`. Internally lengths are EMU (914400 per inch); table and page XML uses twentieths of a point (1440 per inch).

## New document and page setup

```python
doc = Document()                       # built-in template, US Letter
sec = doc.sections[0]
sec.page_width, sec.page_height = Inches(8.5), Inches(11)      # A4: Inches(8.27), Inches(11.69)
sec.left_margin = sec.right_margin = Inches(1)
sec.top_margin = sec.bottom_margin = Inches(1)

doc.core_properties.title = "Quarterly report"
doc.core_properties.author = "Team name"
```

Opening an existing file keeps its template, styles, and headers:

```python
doc = Document("input.docx")           # then save to a NEW filename
```

## Styles: define once, apply by name

```python
normal = doc.styles["Normal"]
normal.font.name = "Arial"
normal.font.size = Pt(11)
normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")   # make the font stick for all scripts
normal.paragraph_format.space_after = Pt(6)

for name, size in (("Heading 1", 18), ("Heading 2", 14), ("Heading 3", 12)):
    st = doc.styles[name]
    st.font.name = "Arial"
    st.font.size = Pt(size)
    st.font.bold = True
    st.font.color.rgb = RGBColor(0, 0, 0)
    st.paragraph_format.space_before = Pt(12)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.keep_with_next = True
```

Custom style:

```python
from docx.enum.style import WD_STYLE_TYPE
callout = doc.styles.add_style("Callout", WD_STYLE_TYPE.PARAGRAPH)
callout.base_style = doc.styles["Normal"]
callout.font.italic = True
callout.paragraph_format.left_indent = Inches(0.4)
doc.add_paragraph("Important note.", style="Callout")
```

## Headings, paragraphs, runs

```python
doc.add_heading("Document title", 0)            # level 0 = Title style
doc.add_heading("Introduction", 1)

p = doc.add_paragraph("Normal text with ")
p.add_run("bold").bold = True
p.add_run(" and ")
p.add_run("italic").italic = True
p.alignment = WD_ALIGN_PARAGRAPH.LEFT           # left-align body text; justify rarely
p.paragraph_format.space_after = Pt(8)
```

Keep formatting in styles. Direct run formatting is for emphasis only.

## Lists

Use the built-in list styles. Never type "•" or "1." by hand.

```python
doc.add_paragraph("First point", style="List Bullet")
doc.add_paragraph("Sub point", style="List Bullet 2")
doc.add_paragraph("Step one", style="List Number")
doc.add_paragraph("Step two", style="List Number")
```

All `List Number` paragraphs share one numbering sequence, so a second numbered list in the same document continues counting (3, 4, ...) instead of restarting. Restarting numbering needs a new numbering instance (`w:num`) in the numbering part, which `python-docx` has no API for. If a document needs several independent numbered lists, tell the user the limitation or fall back to a different list style per list.

## Tables

```python
rows = [("Item", "Qty", "Notes"), ("Widget", "4", "Blue"), ("Gadget", "2", "Red")]
widths = (Inches(3), Inches(1), Inches(2.5))             # sum <= text width (page width - margins)

table = doc.add_table(rows=len(rows), cols=len(widths))
table.style = "Table Grid"
table.autofit = False
table.alignment = WD_TABLE_ALIGNMENT.CENTER

for r, row in enumerate(rows):
    for c, value in enumerate(row):
        cell = table.cell(r, c)
        cell.width = widths[c]                           # set on EVERY cell, not just the table
        cell.text = value
        if r == 0:
            cell.paragraphs[0].runs[0].bold = True
            shd = OxmlElement("w:shd")                   # header shading: clear, never "solid"
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), "D9E2F3")
            cell._tc.get_or_add_tcPr().append(shd)

# repeat the header row on every page
hdr = OxmlElement("w:tblHeader")
hdr.set(qn("w:val"), "true")
table.rows[0]._tr.get_or_add_trPr().append(hdr)
```

Right-align numeric columns with `cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT`. Merge cells with `table.cell(0, 0).merge(table.cell(0, 2))`.

Do not use tables for page layout or as horizontal rules.

## Images

```python
doc.add_picture("chart.png", width=Inches(5))            # give width OR height to keep aspect ratio
inline = doc.inline_shapes[-1]._inline
inline.docPr.set("descr", "Bar chart of revenue by quarter")   # alt text
inline.docPr.set("title", "Revenue by quarter")
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
```

Add a caption as a normal paragraph in the `Caption` style below the picture.

## Page breaks, sections, orientation, columns

```python
doc.add_page_break()                                     # or: run.add_break(WD_BREAK.PAGE)
para.paragraph_format.page_break_before = True           # start this paragraph on a new page

# New landscape section
new = doc.add_section(WD_SECTION.NEW_PAGE)
new.orientation = WD_ORIENT.LANDSCAPE
new.page_width, new.page_height = Inches(11), Inches(8.5)   # swap explicitly; python-docx does not

# Two columns for the current section (edit the existing w:cols element)
cols = new._sectPr.xpath("./w:cols")[0]
cols.set(qn("w:num"), "2")
cols.set(qn("w:space"), "720")                           # gap in twips (720 = 0.5in)
```

Sections created by `add_section` inherit the previous section's header and footer ("link to previous"). Set `section.header.is_linked_to_previous = False` to give a section its own.

## Headers, footers, page numbers

```python
def add_field(paragraph, instruction, placeholder="1"):
    run = paragraph.add_run()
    for kind in ("begin", "instr", "separate", "text", "end"):
        if kind == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = instruction
        elif kind == "text":
            el = OxmlElement("w:t")
            el.text = placeholder
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        run._r.append(el)

sec.header.paragraphs[0].text = "Company report"
footer = sec.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
footer.add_run("Page ")
add_field(footer, "PAGE")
footer.add_run(" of ")
add_field(footer, "NUMPAGES")

sec.different_first_page_header_footer = True            # no header/footer on the title page
```

## Table of contents

A TOC is a field; Word fills it in when the document is opened or the field is updated.

```python
add_field(doc.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u', placeholder="Right-click and choose Update Field")

# ask Word to refresh fields on open
flag = OxmlElement("w:updateFields")
flag.set(qn("w:val"), "true")
doc.settings.element.append(flag)
```

The TOC only lists paragraphs that use heading styles. Word may show a prompt about updating fields on open; mention that to the user. Do not hand-type a table of contents.

## Hyperlinks

```python
def add_hyperlink(paragraph, url, text):
    rid = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), rid)
    run = OxmlElement("w:r")
    props = OxmlElement("w:rPr")
    underline = OxmlElement("w:u"); underline.set(qn("w:val"), "single"); props.append(underline)
    colour = OxmlElement("w:color"); colour.set(qn("w:val"), "0563C1"); props.append(colour)
    run.append(props)
    t = OxmlElement("w:t"); t.text = text; run.append(t)
    link.append(run)
    paragraph._p.append(link)

p = doc.add_paragraph("Details at ")
add_hyperlink(p, "https://example.com", "example.com")
```

## Tab stops (right-aligned date, dot leaders)

```python
from docx.enum.text import WD_TAB_LEADER
p = doc.add_paragraph()
p.paragraph_format.tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
p.add_run("Introduction\t3")
```

`\t` inside a run becomes a tab; this is the one place `\t` in text is correct. Never use `\n`; add another paragraph (or `run.add_break()` for a deliberate soft line break such as an address block).

## Reading content in order

`doc.paragraphs` and `doc.tables` are separate lists. To keep document order:

```python
from docx.table import Table
from docx.text.paragraph import Paragraph

def iter_blocks(doc):
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)
```

`paragraph.text` leaves out text inside tracked insertions (`<w:ins>`). For the text as it reads with changes shown, join the `w:t` elements: `"".join(t.text or "" for t in p._p.iter(qn("w:t")))`.

Headers and footers: `section.header.paragraphs`, `section.footer.paragraphs`. Text boxes, footnotes, and comments live outside the main body and need separate handling.

## Find and replace that keeps formatting

Replace within runs; do not assign `paragraph.text` (it discards run formatting).

```python
def replace_in_runs(doc, old, new):
    count = 0
    def paragraphs():
        yield from doc.paragraphs
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    yield from cell.paragraphs
        for s in doc.sections:
            yield from s.header.paragraphs
            yield from s.footer.paragraphs
    for p in paragraphs():
        for r in p.runs:
            if old in r.text:
                r.text = r.text.replace(old, new)
                count += 1
    return count
```

Word often splits a phrase across runs (spell-check, edits, mixed formatting). When a replacement does not take effect, merge the neighbouring runs that share identical formatting:

```python
def merge_adjacent_runs(paragraph):
    runs = paragraph.runs
    i = 0
    while i < len(runs) - 1:
        a, b = runs[i], runs[i + 1]
        same = (a.bold, a.italic, a.underline, a.font.name, a.font.size, a.font.color.rgb if a.font.color and a.font.color.type else None) == \
               (b.bold, b.italic, b.underline, b.font.name, b.font.size, b.font.color.rgb if b.font.color and b.font.color.type else None)
        if same:
            a.text += b.text
            b._r.getparent().remove(b._r)
            runs = paragraph.runs
        else:
            i += 1
```

Replacement text that spans runs of different formatting needs a judgement call about which formatting wins; keep the first run's.

## Template filling

For `{{placeholders}}` or `[Bracketed]` fields, run the replacement helper over body, tables, headers, and footers; then run `scripts/inspect_docx.py --check`, which flags leftover placeholder text.

## Gotchas

- `paragraph.text = "..."` resets run formatting; edit runs.
- Table width comes from cell widths. Set `cell.width` on every cell and `table.autofit = False`.
- `doc.add_heading(text, 0)` is the Title style, not Heading 1.
- Saving over a file that is open in Word fails on Windows; write a new filename.
- Fields (page numbers, TOC) show placeholder text until Word renders them.
- python-docx cannot render or paginate; page count and layout are only known to Word or a converter.
