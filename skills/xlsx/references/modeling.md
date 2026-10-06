# Financial and scenario models

Conventions for models other people will audit, flex, and reuse. They are widely used industry conventions, not rules of the platform. **If the user supplies a template or house style, follow that instead** and match it exactly.

## Structure

- **Assumptions** on their own sheet or a clearly marked block: growth rates, margins, prices, multiples, dates. Every formula refers to them; none contains the number itself.
- **Calculations** on separate sheets or blocks: one row per line item, one column per period, the same formula across the whole row.
- **Outputs/summary** that reads from the calculations.
- A short **Notes** sheet or header block: purpose, units, version, author, last updated, and how to use it.

## Formulas

```python
ws["C5"] = "=B5*(1+Assumptions!$B$2)"        # growth from an input cell
ws["C6"] = "=IF(C4=0,0,C5/C4)"               # guarded ratio
```

- No hardcoded numbers inside formulas. `=B5*1.05` should be `=B5*(1+$B$6)`.
- Anchor input references with `$` so they survive copying across.
- Keep the formula consistent across a row; a different formula in one column is a bug until proven otherwise.
- Avoid circular references unless the model needs them and iteration is switched on and documented.
- Avoid whole-column references (`A:A`) in heavy lookups; use bounded ranges.
- Prefer `INDEX`/`MATCH` or `XLOOKUP` (check the user's Excel version) over positional lookups.

## Colour conventions (when no template says otherwise)

| Meaning | Font colour |
|---------|-------------|
| Hardcoded input or assumption a user may change | blue `0000FF` |
| Formula or calculation on the same sheet | black `000000` |
| Link pulling from another sheet in the workbook | green `008000` |
| Link to another file | red `FF0000` |

Highlight key assumptions that need attention or review with a light yellow fill. Explain the colours in the Notes sheet.

## Number formats

- Currency and amounts: `#,##0` or `$#,##0`; state units in the header (`Revenue ($000)`).
- Negatives in brackets and zeros as a dash: `#,##0;(#,##0);"-"`.
- Percentages: `0.0%`.
- Multiples: `0.0x`.
- Years as labels: store as text so they display `2025`, not `2,025`.
- Same decimals throughout a block.

## Sourcing hardcoded values

Every hardcoded input needs a source, in a comment or an adjacent cell, in a consistent format:

```
Source: [document or system], [date], [specific reference], [URL if public]
```

Examples: `Source: Annual report FY2024, p.45, revenue note` or `Source: internal ledger export, 2025-01-31, sheet "GL"`. Do not invent a source. If the origin is unknown, write `Source: not provided` and tell the user.

## Checks inside the model

Add a small block of check cells that should always equal zero or TRUE: balance checks, sum-of-parts equals total, no negative cash, input within range. Surface one overall check at the top of the summary.

## Before delivering

1. Recalculate (see `SKILL.md`) and run `scripts/check_xlsx.py`: zero errors.
2. Change a key assumption and confirm outputs move in the expected direction; set it back.
3. Test edge values: zero, negative, very large, blank.
4. Confirm units, signs, and period labels are consistent across sheets.
5. Tell the user what is an assumption and what is sourced data.
