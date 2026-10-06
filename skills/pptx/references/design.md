# Deck design guide

Decide these four things before writing any code, and apply them to every slide.

## 1. Palette

Choose colors for this topic and audience, not a generic default. Use one dominant color (about 60-70% of the visual weight), one or two supporting tones, and a single sharp accent.

Define them once as constants:

```python
from pptx.dml.color import RGBColor
INK     = RGBColor(0x1B, 0x26, 0x3B)   # dominant dark
PAPER   = RGBColor(0xF7, 0xF5, 0xF0)   # light background
MUTED   = RGBColor(0x5C, 0x67, 0x7D)   # secondary text
ACCENT  = RGBColor(0xE0, 0x6C, 0x3C)   # one accent, used sparingly
```

Starting ideas (adapt, do not copy blindly):

| Mood | Dominant | Support | Accent |
|------|----------|---------|--------|
| Finance, serious | deep navy | pale gray-blue | amber |
| Sustainability | forest green | sage | warm sand |
| Product launch | near-black | off-white | vivid coral |
| Healthcare | teal | soft mint | deep blue |
| Editorial | warm charcoal | cream | brick red |

Contrast: text against its background must be clearly legible. Avoid light-gray on white and dark-on-dark.

## 2. Type

Pick a header font with some character and a clean body font. Stick to fonts available on most machines (Georgia, Cambria, Calibri, Arial, Trebuchet MS, Palatino, Consolas).

| Element | Size |
|---------|------|
| Slide title | 32-44pt, bold |
| Section label | 18-24pt |
| Body | 14-18pt |
| Caption, source | 10-12pt, muted |

## 3. Structure

- Dark background for title and closing slides, light for content ("sandwich"), or commit to dark throughout.
- Same margins on every slide (0.5in minimum) and a consistent gap between blocks (pick 0.3in or 0.5in).
- Left-align body text. Center only titles or single big numbers.
- Avoid decorative lines under titles; whitespace does the job.

## 4. One repeated motif

Pick one distinctive element and use it on every slide: icons in colored circles, a thick left border on cards, rounded picture frames, or numbered badges. Consistency reads as design; mixing several reads as noise.

## Layout menu

Vary layouts across the deck; do not repeat the same one more than twice in a row.

- **Title:** large title, subtitle, one visual element.
- **Two column:** text on one side, image or diagram on the other.
- **Icon rows:** icon, bold heading, one-line description, repeated 3-4 times.
- **Grid:** 2x2 or 2x3 cards.
- **Big number:** one stat at 60-72pt with a small label.
- **Comparison:** before/after or option A/B in matched columns.
- **Process:** numbered steps with connectors.
- **Chart:** one chart, one takeaway title, source line.
- **Closing:** the ask or next steps.

## Charts and tables

- One message per chart; the slide title states it.
- Label directly where possible; drop gridlines and legends that add nothing.
- Use the accent color for the series that matters, muted tones for the rest.
- Tables: no more than about 6 rows by 5 columns on a slide; larger data belongs in an appendix.

## Common mistakes

- Default blue with no reason.
- Text-only slides with 6+ bullets.
- Styling the first few slides and leaving the rest plain.
- Small body text to squeeze in more content; split the slide instead.
- Mixing spacing values at random.
- Text boxes with default internal padding when aligning to shapes: set margins explicitly.
