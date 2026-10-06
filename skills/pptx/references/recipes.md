# python-pptx recipes

All snippets assume:

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
```

## New deck at 16:9

```python
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
```

The default template ships layouts indexed 0-10 (0 Title, 1 Title and Content, 5 Title Only, 6 Blank). Inspect `prs.slide_layouts` rather than assuming; templates differ.

```python
for i, l in enumerate(prs.slide_layouts):
    print(i, l.name, [(p.placeholder_format.idx, p.placeholder_format.type) for p in l.placeholders])
```

## Add a slide from a layout

```python
slide = prs.slides.add_slide(prs.slide_layouts[5])   # Title Only
slide.shapes.title.text = "Revenue grew 40% in the second half"
```

Prefer a layout with a title placeholder so the slide has a real title (outline view, navigation, screen readers).

## Background color

```python
fill = slide.background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(0x1B, 0x26, 0x3B)
```

## Text box

```python
tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(6), Inches(3))
tf = tb.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_right = 0
tf.vertical_anchor = MSO_ANCHOR.TOP

p = tf.paragraphs[0]
r = p.add_run()
r.text = "Key point"
r.font.size = Pt(18)
r.font.bold = True
r.font.name = "Calibri"
r.font.color.rgb = RGBColor(0x1B, 0x26, 0x3B)
p.alignment = PP_ALIGN.LEFT

p2 = tf.add_paragraph()
p2.text = "Supporting detail"
p2.space_before = Pt(6)
```

## Bullets

Real bullets come from placeholders or from paragraph levels in a body placeholder. In a free text box, bullets are not automatic; use a body placeholder when you need them:

```python
body = slide.placeholders[1]          # Title and Content layout
tf = body.text_frame
tf.text = "First point"
for t in ("Second point", "Third point"):
    tf.add_paragraph().text = t
tf.paragraphs[1].level = 1            # indent one level
```

## Shapes, cards, icon circles

```python
card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(2), Inches(3.8), Inches(3.5))
card.fill.solid(); card.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
card.line.fill.background()            # no outline
card.shadow.inherit = False

badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.8), Inches(2.3), Inches(0.7), Inches(0.7))
badge.fill.solid(); badge.fill.fore_color.rgb = RGBColor(0xE0, 0x6C, 0x3C)
badge.line.fill.background()
badge.text_frame.text = "1"
```

## Pictures

```python
pic = slide.shapes.add_picture("photo.jpg", Inches(7), Inches(1.5), width=Inches(5.5))
pic._element.nvPicPr.cNvPr.set("descr", "Describe what the image shows")   # alt text
```

Give width or height, not both, to keep the aspect ratio. Always set alt text for meaningful images.

## Native chart

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

cd = CategoryChartData()
cd.categories = ["Q1", "Q2", "Q3", "Q4"]
cd.add_series("Revenue", (12.1, 14.3, 17.8, 21.0))
gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.6), Inches(8), Inches(5), cd)
chart = gf.chart
chart.has_legend = False
chart.plots[0].has_data_labels = True
chart.plots[0].series[0].format.fill.solid()
chart.plots[0].series[0].format.fill.fore_color.rgb = RGBColor(0xE0, 0x6C, 0x3C)
```

Native charts stay editable in PowerPoint; prefer them over pasted images.

## Table

```python
rows, cols = 4, 3
gf = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(2), Inches(8), Inches(2.5))
tbl = gf.table
tbl.cell(0, 0).text = "Metric"
tbl.columns[0].width = Inches(3)
```

## Speaker notes

```python
slide.notes_slide.notes_text_frame.text = "What to say on this slide."
```

## Reading content

```python
for n, slide in enumerate(prs.slides, 1):
    print(n, slide.slide_layout.name)
    for sh in slide.shapes:
        if sh.has_text_frame:
            print("  ", sh.shape_type, sh.name, repr(sh.text_frame.text[:80]))
        if sh.has_chart if hasattr(sh, "has_chart") else False:
            print("   chart", sh.chart.chart_type)
    if slide.has_notes_slide:
        print("  notes:", slide.notes_slide.notes_text_frame.text)
```

Group shapes hold children in `sh.shapes`; recurse when `sh.shape_type == MSO_SHAPE_TYPE.GROUP`.

## Replace text without losing formatting

Setting `text_frame.text` resets formatting. Edit runs instead:

```python
def replace_text(prs, old, new):
    def fix_frame(tf):
        for p in tf.paragraphs:
            for r in p.runs:
                if old in r.text:
                    r.text = r.text.replace(old, new)

    def walk(shapes):
        for sh in shapes:
            if sh.shape_type == 6:                       # group
                walk(sh.shapes)
            elif sh.has_text_frame:
                fix_frame(sh.text_frame)
            elif getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells:
                        fix_frame(cell.text_frame)

    for s in prs.slides:
        walk(s.shapes)
```

A placeholder like `{{name}}` can be split across several runs by PowerPoint. If a replacement does not take effect, join the paragraph's runs: put the full text in the first run and clear the others.

## Delete a slide

python-pptx has no public delete; remove it from the slide id list and drop the relationship:

```python
def delete_slide(prs, index):
    sldIdLst = prs.slides._sldIdLst
    sldId = sldIdLst[index]
    prs.part.drop_rel(sldId.rId)
    sldIdLst.remove(sldId)
```

## Reorder slides

```python
def move_slide(prs, old, new):
    lst = prs.slides._sldIdLst
    el = list(lst)[old]
    lst.remove(el)
    lst.insert(new, el)
```

## Duplicate a slide

There is no public API. Create a slide from the same layout and deep-copy the shapes; pictures and charts need their relationships recreated, so for slides with media prefer rebuilding the slide from code or from the same layout instead of copying.

```python
import copy
def duplicate_simple_slide(prs, index):
    src = prs.slides[index]
    dst = prs.slides.add_slide(src.slide_layout)
    for sh in list(dst.shapes):
        sh._element.getparent().remove(sh._element)
    for sh in src.shapes:
        dst.shapes._spTree.insert_element_before(copy.deepcopy(sh._element), "p:extLst")
    return dst
```

This is safe only for text and plain shapes.

## Merge two decks

Same limitation: copying slides across presentations loses media and chart relationships. Safer options: rebuild the slides in the target deck from their extracted content, or ask the user to use PowerPoint's "Reuse slides" for media-heavy decks. Tell the user which route you took.

## Gotchas

- Units are EMU internally (914400 per inch); use `Inches()`/`Pt()`.
- `Presentation()` is 4:3 by default; set the size before adding slides.
- Autofit is not computed by python-pptx; text can overflow. Size boxes conservatively and run the check script.
- Setting `.text` on a frame discards run formatting.
- Fonts are referenced by name; the viewer's machine must have them.
- Saving to the same path as an open file in PowerPoint fails on Windows; write to a new filename.
