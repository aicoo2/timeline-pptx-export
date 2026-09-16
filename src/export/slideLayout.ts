import { labelIndent, MAX_LABEL_INDENT_STEPS } from '../utils/barNesting';
import type { TaskStatus } from '../types/timeline';
import { estimateWrappedLines, measureLetterSpacingWidthIn, measureTextWidthIn } from './textMetrics';

// Shared page geometry (inches) for both the PPTX and PDF exporters, so a
// "slide" looks identical in either format. Every value the export handoff
// fixes is here, written in its own pixels; see docs/export-handoff-map.md
// for what it fixes and where each number comes from.

// Colour is deliberately *not* here. This module is the slide's geometry and
// type scale; the palette has one home per concern and duplicating any of it
// here would create a second one:
//
//   status colours   src/utils/statusColors.ts — TASK_STATUS_SCALE
//   phase + neutrals src/export/theme.ts    — COLORS, PHASE_PALETTE
//
// Both exporters already read one computed model (timelineExportModel.ts),
// which resolves every fill, so pptxExporter and pdfExporter cannot drift
// apart: neither picks a colour of its own.

/** The handoff draws a 1920x1080 slide; PowerPoint's own 16:9 page is
 * 13.333 x 7.5in, which is the same shape at 144 pixels to the inch. Every
 * measurement below is written as the handoff's own pixel value and converted
 * once, here — so a value can be read straight off docs/export-handoff-map.md
 * and found in this file unchanged. */
export const PX_PER_IN = 144;
const px = (value: number) => value / PX_PER_IN;
/** The handoff's px, as points: 144px = 1in = 72pt. */
const pxPt = (value: number) => value / 2;

export const PAGE_WIDTH_IN = px(1920);
export const PAGE_HEIGHT_IN = px(1080);

// --- The frame every slide shares -------------------------------------------
// `padding: 56px 72px 48px` (gantt-export.html:13), on a white slide with
// --foreground text. There is no header band any more: the title is text on
// the slide, and the deck's other slides take the same frame so a reader does
// not meet two different chromes in one file.

export const FRAME_X_IN = px(72);
export const FRAME_TOP_IN = px(56);
export const FRAME_BOTTOM_IN = px(48);

/** Title: 44px / 600 / -0.02em, one line, never wrapped (`:16`). */
export const TITLE_FONT_SIZE_PT = pxPt(44);
export const TITLE_TRACKING_EM = -0.02;
const TITLE_LINE_RATIO = 1.15;
export const TITLE_LINE_HEIGHT_IN = px(44 * TITLE_LINE_RATIO);
/** The line to the right of the title — the window this slide covers (`:17`). */
export const META_FONT_SIZE_PT = pxPt(26);

// The legend: three status icons and their words, with the zoom caption pushed
// to the right (`:20-24`). 24px text, 24px icons, 30px between items, 10px
// between an icon and its word.
export const LEGEND_MARGIN_TOP_IN = px(26);
export const LEGEND_FONT_SIZE_PT = pxPt(24);
export const LEGEND_ROW_HEIGHT_IN = px(28);
export const LEGEND_ITEM_GAP_IN = px(30);
export const LEGEND_ICON_GAP_IN = px(10);
/** Status icons are 24x24 wherever they appear — legend and task cell alike. */
export const STATUS_ICON_SIZE_IN = px(24);

/** The status glyphs, on the handoff's own 24-unit grid.
 *
 * The handoff ships three inline SVGs and says to match them exactly, so the
 * three are transcribed here from its paths rather than substituted with the
 * lucide icon they were drawn from: a ringed check, a ringed play triangle,
 * and a bare ring.
 *
 * Drawn from primitives rather than from an image: PowerPoint would take an
 * SVG, jsPDF would not, and the handoff's own first export trap is icons that
 * are not vectors.
 *
 * One geometry, two engines: each scales the grid to STATUS_ICON_SIZE_IN. */
export const STATUS_ICON_GRID = 24;
export const STATUS_ICON_STROKE_UNITS = 2;
export const STATUS_ICON_RING_RADIUS_UNITS = 10;

export interface StatusIconGeometry {
  /** Points of a stroked polyline, in grid units. */
  polyline?: readonly (readonly [number, number])[];
  /** A filled triangle, as its three corners in grid units, in path order. */
  triangle?: readonly [
    readonly [number, number],
    readonly [number, number],
    readonly [number, number],
  ];
}

export const STATUS_ICON_GEOMETRY: Record<TaskStatus, StatusIconGeometry> = {
  // The handoff's check: "M7.5 12.3 L10.6 15.4 L16.5 9.2"
  done: { polyline: [[7.5, 12.3], [10.6, 15.4], [16.5, 9.2]] },
  // The handoff's play triangle, filled: "M9.3 7.2 L17 12 L9.3 16.8 Z"
  in_progress: { triangle: [[9.3, 7.2], [17, 12], [9.3, 16.8]] },
  // The ring alone
  todo: {},
};

/** The legend's three entries, in the order a task moves through them — the
 * handoff's own list, and its own words (Title Case, unlike the lowercase
 * TASK_STATUS_LABELS the appendix uses). */
export const LEGEND_ITEMS: { status: TaskStatus; label: string }[] = [
  { status: 'done', label: 'Done' },
  { status: 'in_progress', label: 'In progress' },
  { status: 'todo', label: 'To do' },
];

// The chart card (`:27`): a white surface with a hairline border and a 14px
// radius, holding the column header and every row.
export const CARD_MARGIN_TOP_IN = px(16);
export const CARD_RADIUS_IN = px(14);
export const CARD_BORDER_WIDTH_PT = 0.75;

/** The footer line — the coverage note ("+N tasks not shown"), and nothing
 * else. Not in the handoff, which has no footer at all; kept because the note
 * is the deck's own promise that nothing was dropped silently, and sized at
 * the handoff's 24px floor for that reason.
 *
 * The provenance mark used to share this line. It moved to the right of the
 * title (EXPORT_MARK_TEXT), which is where the overview's window caption and
 * zoom caption used to be — so the line is reserved whether or not a slide
 * has a note to put on it, and the geometry below is unchanged by the move. */
export const FOOTER_HEIGHT_IN = px(28);
export const FOOTER_FONT_SIZE_PT = pxPt(24);

export const CONTENT_X_IN = FRAME_X_IN;
export const CONTENT_WIDTH_IN = PAGE_WIDTH_IN - FRAME_X_IN * 2;
/** Where a slide's content starts under its title. The overview puts its
 * legend here; every other slide starts its own content at the same line, so
 * the deck keeps one rhythm. */
export const CONTENT_TOP_IN = FRAME_TOP_IN + TITLE_LINE_HEIGHT_IN + LEGEND_MARGIN_TOP_IN;
export const CONTENT_BOTTOM_IN = PAGE_HEIGHT_IN - FRAME_BOTTOM_IN - FOOTER_HEIGHT_IN;
export const CONTENT_HEIGHT_IN = CONTENT_BOTTOM_IN - CONTENT_TOP_IN;

export const CARD_X_IN = CONTENT_X_IN;
export const CARD_WIDTH_IN = CONTENT_WIDTH_IN;
export const CARD_TOP_IN = CONTENT_TOP_IN + LEGEND_ROW_HEIGHT_IN + CARD_MARGIN_TOP_IN;
export const CARD_BOTTOM_IN = CONTENT_BOTTOM_IN;
export const CARD_HEIGHT_IN = CARD_BOTTOM_IN - CARD_TOP_IN;

// --- The card's own grid ----------------------------------------------------
// `grid-template-columns: 480px repeat(N, 1fr)` (`:29`): a fixed task column
// and N equal time columns. Equal, deliberately — the handoff positions bars
// as a percentage of the window and never in column indices, so a month column
// is as wide as the month beside it whatever their day counts (README, step 1).

export const TASK_COL_WIDTH_IN = px(480);
export const TIMELINE_X_IN = CARD_X_IN + TASK_COL_WIDTH_IN;
export const TIMELINE_WIDTH_IN = CARD_WIDTH_IN - TASK_COL_WIDTH_IN;

// The column header row (`:30-31`): "TASK" set small and tracked out over the
// task column, then two centred lines per time column.
const COLUMN_HEADER_PAD_TOP_IN = px(12);
const COLUMN_HEADER_PAD_BOTTOM_IN = px(14);
export const COLUMN_HEADER_LINE_GAP_IN = px(2);
export const COLUMN_HEADER_FONT_SIZE_PT = pxPt(24);
export const COLUMN_HEADER_LINE_IN = px(28);
export const COLUMN_HEADER_HEIGHT_IN =
  COLUMN_HEADER_PAD_TOP_IN + COLUMN_HEADER_LINE_IN * 2 + COLUMN_HEADER_LINE_GAP_IN + COLUMN_HEADER_PAD_BOTTOM_IN;
export const COLUMN_HEADER_TOP_LINE_Y_IN = COLUMN_HEADER_PAD_TOP_IN;
export const COLUMN_HEADER_SUB_LINE_Y_IN =
  COLUMN_HEADER_PAD_TOP_IN + COLUMN_HEADER_LINE_IN + COLUMN_HEADER_LINE_GAP_IN;
/** "TASK" — 0.08em of tracking, uppercase, muted (`:30`). */
export const COLUMN_HEADER_TRACKING_EM = 0.08;

// The rows area: everything under the column header, inside the card.
export const ROWS_AREA_TOP_IN = CARD_TOP_IN + COLUMN_HEADER_HEIGHT_IN;
export const ROWS_AREA_HEIGHT_IN = CARD_BOTTOM_IN - ROWS_AREA_TOP_IN;

// One row height, on every slide of the deck.
//
// **This is a deliberate divergence from the handoff**, asked for by the
// customer. The handoff makes a row `flex:1`, so the rows share the card's
// height and their pitch is a result of how many a slide carries — which
// means a slide of four rows draws them three times taller than a slide of
// fifteen, and two slides of the same deck cannot be read against each other.
// A fixed pitch costs the opposite: a slide with few rows is left part-empty,
// its card ruled down to a blank lower half. That was put to the customer with
// their own example (slide 34, half filled) and accepted. See Phase 5 in
// docs/export-handoff-map.md.
//
// The number is the one the floor already used: 48px, the densest row the
// prototype draws (16 rows in ~761px). So a row is exactly as tall as the
// tightest row the handoff itself draws, and never tighter.
export const OVERVIEW_ROW_HEIGHT_IN = px(48);

/** How many rows a slide holds — the rows area divided by the pitch above,
 * which is also what the Compact/Full overflow rules cut to. Derived rather
 * than written down, so the capacity can never disagree with the height the
 * rows are actually drawn at. */
export const MAX_OVERVIEW_BARS_PER_SLIDE = Math.floor(ROWS_AREA_HEIGHT_IN / OVERVIEW_ROW_HEIGHT_IN);

/** The hairline under a row (`:67`) — `--border` at 0.6 alpha, against the
 * full-strength border the card and the column rules use. */
export const ROW_RULE_ALPHA = 0.6;
export const ROW_RULE_WIDTH_PT = 0.75;

// The task cell (`:100`, `:140`): a phase starts 24px in, a nested task 52px,
// and each further level adds the same 28px step — the handoff draws two
// levels and gives one step between them, so deeper trees continue it, capped
// where the screen's own ladder caps (MAX_LABEL_INDENT_STEPS).
export const TASK_CELL_PAD_IN = px(24);
export const TASK_CELL_INDENT_STEP_IN = px(28);
export const TASK_ICON_GAP_IN = px(14);
export const TASK_NAME_FONT_SIZE_PT = pxPt(26);

/** A task cell's left inset at `depth`, in inches. */
export function taskCellIndent(depth: number): number {
  const steps = Math.min(Math.max(depth, 0), MAX_LABEL_INDENT_STEPS);
  return TASK_CELL_PAD_IN + TASK_CELL_INDENT_STEP_IN * steps;
}

// Bars (`:104`, `:146`): a phase is 26px of solid colour, anything nested is
// 20px of the same hue at the palette's tint. 8px radius on both, and a floor
// of 8px so a same-day task is still a mark rather than nothing.
export const OVERVIEW_BAR_HEIGHT_IN = px(26);
export const OVERVIEW_NESTED_BAR_HEIGHT_IN = px(20);
export const BAR_RADIUS_IN = px(8);
export const MIN_BAR_WIDTH_IN = px(8);

/** A bar carries no text of its own — the task's name sits in the column to its
 * left. It is still drawn as a *text frame* rather than a plain shape, so that
 * someone labelling a bar in PowerPoint after the export types at the deck's
 * size and flush to the bar's edge, instead of at PowerPoint's default 18pt
 * inset by a default margin. An empty frame carries nothing else.
 *
 * `BAR_OUTLINE_WIDTH_PT` only restates the width `addShape` applied silently on
 * its own: `addText` has no such default, and the bar may not change width. */
export const BAR_OUTLINE_WIDTH_PT = 1;
export const BAR_TYPING_FONT_SIZE_PT = 10;

/** The "today" rule (`:66`): 2px of --destructive at 80%, the full height of
 * the rows area. */
export const TODAY_LINE_WIDTH_PT = 1.5;
export const TODAY_LINE_ALPHA = 0.8;

// A chevron marks a bar the export timeframe clipped. Drawn as a triangle
// rather than as a character: the handoff's first export trap is glyph icons,
// which macOS substitutes with colour emoji.
export const CHEVRON_WIDTH_IN = px(10);
export const CHEVRON_HEIGHT_IN = px(14);

// --- The appendix, dashboard and summary slides ------------------------------
// The handoff describes the chart slide and nothing else, so what follows is
// unchanged except for the frame it now sits in.

export const ROW_LABEL_HEIGHT_IN = 0.22;

/** The unit the appendix's rows are measured in, and the base of its indent
 * ladder (subtaskRowIndent below). It is not the overview's bar height any
 * more — the handoff fixes that at 26px — but it keeps its name because the
 * indent-parity check and its documentation are written against it. */
export const BAR_HEIGHT_IN = 0.28;

// "Back to overview" on the appendix slides. It sits on the title's own line,
// at the right — the slot the overview gives its window caption, which every
// other slide leaves empty. Below the title it crowded the first section's
// heading, and beside it there is a whole slide's width doing nothing.
//
// No arrow glyph in front of it: an appendix slide is not the place to find
// out that a font substituted one (the handoff's first export trap), and the
// words say it without help.
export const BACK_LINK_TEXT = 'Back to overview';
export const BACK_LINK_FONT_SIZE_PT = META_FONT_SIZE_PT;
export const BACK_LINK_Y_IN = FRAME_TOP_IN;
export const BACK_LINK_HEIGHT_IN = TITLE_LINE_HEIGHT_IN;
export const BACK_LINK_WIDTH_IN = 2.2;

// Left padding inside a column, so no text sits on a column edge.
// Appendix type scale. A detail slide's subtask row packs a label and its
// dates on the left and a status on the right of one line.
export const SUBTASK_TEXT_FONT_SIZE_PT = 11;
export const SUBTASK_STATUS_FONT_SIZE_PT = 10;
export const SUBTASK_DATE_FONT_SIZE_PT = 10;
export const STATUS_LETTER_SPACING_EM = 0.06;
export const DATE_LETTER_SPACING_EM = 0.02;
export const SUMMARY_LEGEND_STATUS_FONT_SIZE_PT = 9;
/** The clear space between two appendix rows. */
export const ROW_GAP_IN = 0.04;

/** Width of a status word as it is actually drawn: its glyphs plus its
 * tracking, which neither engine folds into its own metrics. Everything that
 * has to fit a status into a box measures it through here. */
export function statusTextWidthIn(text: string, fontSizePt: number): number {
  return (
    measureTextWidthIn(text, fontSizePt) +
    measureLetterSpacingWidthIn(text, letterSpacingPt(fontSizePt, STATUS_LETTER_SPACING_EM))
  );
}

/** Tracking in points for `fontSizePt` — the unit pptxgenjs's charSpacing
 * wants, and 1/72 of what jsPDF's setCharSpace wants in inches. */
export function letterSpacingPt(fontSizePt: number, em: number): number {
  return fontSizePt * em;
}

// Clear space kept between a right-aligned status text and the right edge of
// the content area, on the overview bars and on a detail slide's subtask
// rows alike. Without it the status is anchored *on* that edge: measured in
// the generated PDF, "In progress" ended at 9.502in against a content edge
// of 9.500, i.e. touching it and a hair past. Wide enough to read as a
// margin rather than as a rounding accident.
export const STATUS_RIGHT_PADDING_IN = 0.12;

// Top of the dashboard tables' slides (Delayed). A slide's content
// normally starts at CONTENT_TOP_IN, which is right for text — but these
// slides open with a filled table header, and a solid band wants a little
// more air under the title than a line of text does. Kept as its own constant
// rather than raising CONTENT_TOP_IN, which the overview's card is measured
// from.
export const DASHBOARD_TABLE_TOP_IN = CONTENT_TOP_IN + 0.18;
// How wide a dashboard table is drawn. The whole content width now: the slide
// used to keep a 2.4in column on the right for a QR code into the same list on
// screen, and that screen is gone. Lives here rather than privately in each
// exporter because the model has to know how wide a table's columns are to work
// out how many of its rows fit (see fitTableRows in dashboardSlides).
export const DASHBOARD_TABLE_WIDTH_IN = CONTENT_WIDTH_IN;
// All the height a dashboard table has. Neither engine paginates a table — the
// PDF is explicitly stopped from trying (withoutPageBreaks) — so rows past this
// are drawn off the slide unless the model cuts them first.
export const DASHBOARD_TABLE_MAX_HEIGHT_IN = CONTENT_BOTTOM_IN - DASHBOARD_TABLE_TOP_IN;

// Minimum clear gap kept between a detail row's label and the status text
// sharing its line, after the label has been truncated against the status's
// own measured width — a buffer against the font-metric approximation in
// textMetrics.ts, so a truncated label never visually touches the status.
export const SUBTASK_META_STATUS_GAP_IN = 0.05;
// Gap between the three pieces of a subtask row's left side — task name,
// then its dates, then its progress. These used to be one string joined by
// "  —  " separators; now that each piece is drawn in its own face and size
// (name / monospace date / progress), whitespace alone separates them and
// the dashes would just be noise.
export const SUBTASK_META_GAP_IN = 0.09;

// The whole subtask block's inset from a detail slide's content edge: what
// separates the rows from the section title above them, before any nesting is
// taken into account. Every row gets it whatever its depth — the levels are
// told apart by subtaskRowIndent() below.
export const DETAIL_ROW_INDENT_IN = 0.2;

/** A subtask row's left indent, for a `depth` measured *within its section*
 * (0 = a direct child of the section's parent).
 *
 * The block's own inset, plus the shared depth ladder from barNesting rebased
 * so that first level sits exactly at that inset. Rebasing — rather than
 * starting the ladder at rung 0 — is what makes the *step* between two levels
 * identical to the on-screen label column's at every depth, the cap included:
 * on screen an indent stops growing after MAX_LABEL_INDENT_STEPS, and a ladder
 * that started a rung late would still be stepping once the screen had stopped,
 * so five-level trees would read differently on the two surfaces.
 *
 * Lives here rather than in either exporter because the model needs it to place
 * the row and to size the text that truncates against what's left, and the
 * coverage check needs it to verify neither drifted. */
export function subtaskRowIndent(depth: number): number {
  return DETAIL_ROW_INDENT_IN + labelIndent(BAR_HEIGHT_IN, depth + 1) - labelIndent(BAR_HEIGHT_IN, 1);
}
// Breathing room either side of a slide's own date window, as a fraction of
// the *widest* window in the export rather than of each slide's own.
//
// A fraction of each slide's own span would pad a three-month slide by days
// and a three-year one by months, so the same clear space would read as a
// different amount of time on every slide. Taken from the widest window, one
// ratio produces one identical margin in inches everywhere — which is what
// "the same axis, drawn at the same density" has to mean visually.
export const OVERVIEW_WINDOW_PAD_RATIO = 0.02;

export const LIST_ROW_HEIGHT_IN = 0.32;
export const SECTION_GAP_IN = 0.2;
// Gap between one parent's whole subtasks/comments block and the next
// parent's block when several are packed onto the same appendix slide.
export const PARENT_SECTION_GAP_IN = 0.3;

// Body text of a comment's paragraphs and list items. Lives here rather
// than in each exporter because the model measures wrapped line counts
// against it (estimateBlockHeight) at exactly the size both engines then
// draw it at — three private copies of the number is how the three quietly
// drift apart.
export const COMMENT_BODY_FONT_SIZE_PT = 12;

// Every single-line row on a detail slide (subtask rows, the section
// headings) is drawn vertically centered in its own row box rather than hung
// from the box's top: the pieces sharing one line are drawn at three
// different sizes, and a shared top edge puts three different baselines on
// what should read as one line.
//
// In pptxgenjs that's `valign: 'middle'` on a box of the row's height; in
// jsPDF, `baseline: 'middle'` at the row's center Y, which this helper
// resolves so the two engines can't drift apart.
export function rowCenterY(rowY: number, rowHeightIn: number): number {
  return rowY + rowHeightIn / 2;
}

// Comment bodies are parsed markdown (see src/utils/renderMarkdown.ts) and
// rendered as real headings/paragraphs/lists/tables. Heights below reuse the
// same row-pitch scale as the rest of a detail section (ROW_LABEL_HEIGHT_IN
// for heading-weight rows, LIST_ROW_HEIGHT_IN for body-weight rows) so a
// comment's blocks sit at the same visual rhythm as subtask rows.
export const COMMENT_META_ROW_HEIGHT_IN = 0.18;
export const COMMENT_HEADING_ROW_HEIGHT_IN = ROW_LABEL_HEIGHT_IN;
export const COMMENT_LINE_HEIGHT_IN = LIST_ROW_HEIGHT_IN;
// A markdown table inside a comment. These are *derived* rather than picked,
// because jspdf-autotable will draw whatever height its own metrics say and the
// model has to predict it exactly: a table estimated shorter than it draws runs
// past the bottom of the content area, and autoTable answers that by inserting
// a page of its own — a physical PDF page with no header, no footer and no
// slide model behind it, which desyncs every slide after it (see
// docs/export-coverage.md, and the note above drawTableBlock in pdfExporter).
//
// One text line plus padding above and below it, where the line height is the
// font size times jsPDF's own 1.15 line-height factor. The padding is passed
// *into* autoTable rather than read off it, so the renderer follows this number
// instead of this number chasing the renderer's defaults.
export const COMMENT_TABLE_FONT_SIZE_PT = 9;
export const COMMENT_TABLE_CELL_PADDING_IN = 0.07;
const PDF_LINE_HEIGHT_FACTOR = 1.15;
export const COMMENT_TABLE_LINE_HEIGHT_IN = (COMMENT_TABLE_FONT_SIZE_PT * PDF_LINE_HEIGHT_FACTOR) / 72;

/** Text width of one column of a `tableWidthIn`-wide table with `columnCount`
 * equal columns. Equal columns are what pptxgenjs is told to draw and what
 * jspdf-autotable converges on at these widths, so one number serves both. */
export function tableColumnTextWidthIn(tableWidthIn: number, columnCount: number): number {
  return tableWidthIn / Math.max(columnCount, 1) - COMMENT_TABLE_CELL_PADDING_IN * 2;
}

/** Height of one table row whose cells wrap at `columnTextWidthIn`. Every cell
 * is measured and the tallest decides the row: a cell whose text wraps to two
 * lines makes the whole row two lines tall in either renderer, and a row
 * assumed to be one line is exactly how a table ends up taller than the space
 * reserved for it. Shared by a comment's markdown tables and the dashboard's
 * own, so "how tall is this table" has a single answer. */
export function tableRowHeightIn(cells: string[], columnTextWidthIn: number): number {
  const lines = Math.max(
    ...cells.map((cell) => estimateWrappedLines(cell, COMMENT_TABLE_FONT_SIZE_PT, columnTextWidthIn)),
  );
  return lines * COMMENT_TABLE_LINE_HEIGHT_IN + COMMENT_TABLE_CELL_PADDING_IN * 2;
}
// Small vertical breathing room between adjacent blocks within one comment,
// and between one comment's blocks and the next comment's meta line.
export const COMMENT_BLOCK_GAP_IN = 0.06;
export const COMMENT_GAP_IN = 0.14;
