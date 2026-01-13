import {
  Document,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  HeightRule,
  BorderStyle,
  TextDirection,
  VerticalAlign,
  TableLayoutType,
} from "docx";

export { Packer } from "docx";

const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const COLUMN_WIDTH = 1152; // dxa (~0.8 inches)
const TABLE_WIDTH = 8064; // dxa (7 columns * 1152)
const ROW_HEIGHT = 576; // dxa (~0.4 inches)
const GRAY_COLOR = "A6A6A6";
const WEEKEND_SHADING = "F2F2F2";
const FONT_SIZE = 16; // half-points (8pt)

const NIL_BORDER = { style: BorderStyle.NIL, size: 0, color: "FFFFFF" };
const SINGLE_BORDER = { style: BorderStyle.SINGLE, size: 4, color: "000000" };

function getDaysInMonth(year: number, month: number): number {
  return new Date(year, month, 0).getDate();
}

/** Returns day of week for the 1st of the month: 0 = Sunday, 6 = Saturday. */
function getFirstDayOfWeek(year: number, month: number): number {
  return new Date(year, month - 1, 1).getDay();
}

function isWeekend(dayOfWeek: number): boolean {
  return dayOfWeek === 0 || dayOfWeek === 6;
}

/** Cell with no content, all borders hidden, no shading. Used for padding. */
function createEmptyCell(): TableCell {
  return new TableCell({
    children: [new Paragraph({})],
    width: { size: COLUMN_WIDTH, type: WidthType.DXA },
    borders: {
      top: NIL_BORDER,
      bottom: NIL_BORDER,
      left: NIL_BORDER,
      right: NIL_BORDER,
    },
  });
}

/**
 * Month label cell (e.g. "Mar"). Gray text, bottom-aligned.
 * @param rotated - true: 90° CCW rotation, right border visible (inline before day 1).
 *                  false: no rotation, all borders hidden (in dedicated label row).
 */
function createLabelCell(monthName: string, rotated: boolean): TableCell {
  const paragraph = new Paragraph({
    children: [
      new TextRun({
        text: monthName,
        color: GRAY_COLOR,
        size: FONT_SIZE,
      }),
    ],
    alignment: rotated ? "center" : undefined,
  });

  return new TableCell({
    children: [paragraph],
    width: { size: COLUMN_WIDTH, type: WidthType.DXA },
    verticalAlign: VerticalAlign.BOTTOM,
    textDirection: rotated ? TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT : undefined,
    borders: rotated
      ? {
          top: NIL_BORDER,
          bottom: NIL_BORDER,
          left: NIL_BORDER,
          right: SINGLE_BORDER,
        }
      : {
          top: NIL_BORDER,
          bottom: NIL_BORDER,
          left: NIL_BORDER,
          right: NIL_BORDER,
        },
  });
}

/**
 * Day cell with two lines: gray day number + empty line for user notes.
 * The trailing space after the number is unstyled, so user-typed text appears black.
 * Weekend cells (Sun/Sat) have light gray background shading.
 */
function createDayCell(day: number, dayOfWeek: number): TableCell {
  const weekend = isWeekend(dayOfWeek);
  return new TableCell({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: String(day),
            color: GRAY_COLOR,
            size: FONT_SIZE,
          }),
          new TextRun({
            text: " ",
            size: FONT_SIZE,
          }),
        ],
      }),
      new Paragraph({
        children: [],
      }),
    ],
    width: { size: COLUMN_WIDTH, type: WidthType.DXA },
    shading: weekend ? { fill: WEEKEND_SHADING } : undefined,
    borders: {
      top: SINGLE_BORDER,
      bottom: SINGLE_BORDER,
      left: SINGLE_BORDER,
      right: SINGLE_BORDER,
    },
  });
}

/**
 * Generates rows for a continuous calendar table spanning multiple months.
 *
 * Each month occupies 5-6 rows. Month label placement depends on which day
 * of the week day 1 falls:
 * - Day 1 is Sunday: insert a dedicated label row above, label not rotated
 * - Day 1 is Mon-Sat: label goes in the cell immediately before day 1, rotated 90° CCW
 *
 * Empty cells (no borders, no shading) fill gaps: before the label in the first
 * week, and after the last day of the month.
 */
function generateCalendarRows(startYear: number, startMonth: number, months: number): TableRow[] {
  const rows: TableRow[] = [];

  for (let m = 0; m < months; m++) {
    const year = startYear + Math.floor((startMonth - 1 + m) / 12);
    const month = ((startMonth - 1 + m) % 12) + 1;
    const monthName = MONTH_NAMES[month - 1];
    const daysInMonth = getDaysInMonth(year, month);
    const firstDayOfWeek = getFirstDayOfWeek(year, month);

    // Case 2: Day 1 is Sunday - need a separate label row
    if (firstDayOfWeek === 0) {
      const labelCells: TableCell[] = [createLabelCell(monthName, false)];
      for (let i = 1; i < 7; i++) {
        labelCells.push(createEmptyCell());
      }
      rows.push(
        new TableRow({
          children: labelCells,
          height: { value: ROW_HEIGHT, rule: HeightRule.EXACT },
          cantSplit: true,
        })
      );
    }

    // Generate week rows
    let currentDay = 1;
    let isFirstWeek = true;

    while (currentDay <= daysInMonth) {
      const cells: TableCell[] = [];

      for (let dayOfWeek = 0; dayOfWeek < 7; dayOfWeek++) {
        if (isFirstWeek && dayOfWeek < firstDayOfWeek) {
          // Before the first day of the month
          if (firstDayOfWeek !== 0 && dayOfWeek === firstDayOfWeek - 1) {
            // This is the label cell (Case 1: day 1 is not Sunday)
            cells.push(createLabelCell(monthName, true));
          } else {
            // Empty cell before the label
            cells.push(createEmptyCell());
          }
        } else if (currentDay <= daysInMonth) {
          // Regular day cell
          cells.push(createDayCell(currentDay, dayOfWeek));
          currentDay++;
        } else {
          // After the last day of the month
          cells.push(createEmptyCell());
        }
      }

      rows.push(
        new TableRow({
          children: cells,
          height: { value: ROW_HEIGHT, rule: HeightRule.EXACT },
          cantSplit: true,
        })
      );

      isFirstWeek = false;
    }
  }

  return rows;
}

/** Creates a Word document containing a single calendar table. */
export function generateCalendarDocx(startYear: number, startMonth: number, months: number): Document {
  const rows = generateCalendarRows(startYear, startMonth, months);

  const table = new Table({
    rows,
    width: { size: TABLE_WIDTH, type: WidthType.DXA },
    columnWidths: Array(7).fill(COLUMN_WIDTH),
    layout: TableLayoutType.FIXED,
  });

  return new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Calibri",
            size: FONT_SIZE,
            noProof: true,
          },
          paragraph: {},
        },
      },
    },
    sections: [
      {
        children: [table],
      },
    ],
  });
}
