# Calendar Table Specification

This document describes the structure and styling of the Word calendar tables based on analysis of reference/example1.docx, example2.docx, and example3.docx.

## Overall Structure

The calendar is a single continuous table spanning multiple months. Each month occupies 5-6 rows depending on how its days fall.

### Table Properties

- **Total width**: 8064 dxa (twips) ≈ 5.6 inches
- **Layout**: Fixed column widths
- **Columns**: 7 columns (Sun-Sat), each ~1152 dxa (~0.8 inches)
- **Row height**: 576 dxa (~0.4 inches) per row
- **Cell margins**: Left/right set to 0

### Grid Layout

```
| Sun | Mon | Tue | Wed | Thu | Fri | Sat |
```

Each week is one table row. A month flows continuously from one row to the next. Rows start on Sun and end on Sat; this is not configurable.

## Month Label

The 3-letter month abbreviation (e.g., "Mar", "Apr", "Jul") placement depends on which day of the week day 1 falls:

**Case 1: Day 1 is NOT Sunday**
- Label appears in the cell immediately preceding day 1 (same row, one column to the left)
- Rotated 90° counter-clockwise (`<w:textDirection w:val="btLr"/>`)
- Vertically aligned to bottom (`<w:vAlign w:val="bottom"/>`)
- Horizontally centered within the rotated text area

**Case 2: Day 1 is Sunday** (see example3)
- A dedicated "label row" is inserted before the first week of the month
- Label appears in the Sunday (first) cell of this row, not rotated
- Vertically aligned to bottom (`<w:vAlign w:val="bottom"/>`)
- All 7 cells in the label row are empty (except the label in the first cell), with all borders hidden and no background shading
- Day 1 appears in the following row

Common styling:
- Gray color (#A6A6A6)
- Font size 16 half-points (8pt)

Month names are not configurable. They are hard-coded to English three-letter abbreviations.

## Cell Content Structure

Each day cell contains a paragraph with two lines:

**Line 1:** Day number (1-31) followed by a space
- The day number is in a text run with gray color (#A6A6A6)
- The trailing space is in a separate text run (or part of default paragraph styling)

**Line 2:** Empty (for user to add events)

The paragraph as a whole has default black text color. This allows users to position their cursor on line 2 (or after the space on line 1) and type in black.

Font size: 16 half-points (8pt) throughout.

## Cell Background Colors

| Color | Hex | Usage |
|-------|-----|-------|
| Light gray | #F2F2F2 | Weekend days (Saturday, Sunday) |
| None | - | Weekdays (Monday-Friday) |

## Cell Borders

- Internal borders: Single line, default weight
- The month label cell (Case 1, rotated) has left border hidden, right border visible

## Empty Cells (No Content, No Borders, No Shading)

"Empty cells" appear in several situations and share common properties:
- No content
- All borders hidden (top/left/right/bottom set to `nil`)
- No background shading

Empty cells occur in:

1. **Leading cells in a month's first row** (when day 1 is not Sunday): cells before the month label cell
2. **Trailing cells in a month's last row** (when the month doesn't end on Saturday): cells after the last day
3. **Label row cells** (when day 1 is Sunday): all 7 cells in the label row (except the first cell contains the month label text)

## Row Properties

- `<w:cantSplit/>` - Prevents row from breaking across pages
- Fixed height of 576 dxa

## Typography

All text uses:
- Font size: 16 half-points (8pt equivalent)
- Both `<w:sz>` and `<w:szCs>` set to 16

## Color Reference

| Element | Color Code |
|---------|------------|
| Day number text run | #A6A6A6 (gray) |
| Month label | #A6A6A6 (gray) |
| Default paragraph text | #000000 (black) |
| Weekend background | #F2F2F2 |
