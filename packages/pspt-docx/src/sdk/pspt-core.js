'use strict';
/**
 * =============================================================================
 *  PIC MATRIX / DOCUMENTATION SOP — SHARED STYLE CONSTANTS
 * =============================================================================
 * Every value below was extracted directly from the source template's XML
 * (word/document.xml, word/styles.xml, word/header1.xml, word/footer1.xml) —
 * not eyeballed from a render. This is the single source of truth for fonts,
 * sizes, colors, page geometry, and table proportions. Do not hand-tune
 * values in the SDK file itself; change them here so every section stays in
 * sync.
 * =============================================================================
 */

// -----------------------------------------------------------------------
// FONT
// -----------------------------------------------------------------------
// The template body font is Calibri throughout (docDefaults in styles.xml).
const DOCX_FONT = 'Calibri';

// -----------------------------------------------------------------------
// SIZES (docx "size" is in half-points: 20 = 10pt, 22 = 11pt, etc.)
// -----------------------------------------------------------------------
const DOCX_SIZE = {
  title: 56, // Cover page product/document title ("Title" style) = 28pt
  subtitle: 24, // Cover page italic tagline under the title = 12pt
  partTitle: 36, // "Part" divider page heading & Heading1 in the master doc = 18pt
  h1: 30, // Numbered section heading ("1. Team") = 15pt; also Heading2 in master doc
  h2: 24, // Sub-heading inside a section = 12pt; also Heading3 in master doc
  h3: 20, // Minor heading = 10pt
  h4: 21, // Heading4 in the master doc (subsections one level deeper) = 10.5pt
  body: 20, // Standard paragraph / table cell text = 10pt
  tableHeader: 20, // Navy table header row text = 10pt
  labelCell: 20, // Bold label cells (cover info table, identity tables) = 10pt
  headerFooter: 16, // Page header/footer text = 8pt
  caption: 16, // Caption/footnote text = 8pt
};

// -----------------------------------------------------------------------
// COLORS (hex, no leading '#', matching docx w:color / w:fill values)
// -----------------------------------------------------------------------
const DOCX_COLOR = {
  black: '000000',
  white: 'ffffff',

  // Headings
  h1: '1f3b57', // Deep navy — main document title & Heading1 base color
  h1Rendered: '2e74b5', // Actual on-screen H1 run color override in the template
  h2: '1f3b57',
  h3: '2f6690', // Slate blue — Heading3 / subtitle italic color
  headingRule: '3e92cc', // Bottom-border rule color under H1-level headings
  headingRuleThin: '3e92cc',

  // Body / muted text
  bodyText: '000000',
  mutedText: '6b7280', // Footer text, secondary notes
  purposeText: '2f6690', // Italic "Purpose:" explanatory line color
  purposeLabel: '1f3b57',

  // Tables
  tableHeaderBg: '1f3b57', // Navy fill for table header rows
  tableHeaderText: 'ffffff',
  tableBorder: 'd0d5dd', // Light gray-blue cell border (PIC matrix / cover tables)
  tableBorderAlt: 'cbd5e0', // Light gray-blue cell border variant used in Style-Guide-style tables
  tableOuterBorder: '000000', // Outer table border (tblBorders) — solid black
  zebraFill: 'edf2f7', // Light blue-gray fill for odd body rows / info tables
  zebraFillAlt: 'ffffff', // White fill for even body rows

  // Cover / identity table
  coverInfoFill: 'edf2f7', // Uniform light fill across the whole cover info table
  coverLabelText: '1f3b57',

  // Links
  hyperlink: '0000ee',

  // Header/footer chrome
  headerFooterText: '1f3b57',
  headerFooterMuted: '6b7280',
  confidential: '990000',
  headerRuleColor: 'd0d5dd',
  footerRuleColor: 'd0d5dd',

  // Placeholder / callout text
  placeholderText: '2f6690',
  highlightYellow: 'FFFF00', // Used for dropdown "selected value" style callouts if needed
};

// -----------------------------------------------------------------------
// PAGE GEOMETRY (DXA; 1440 = 1 inch) — US Letter, exactly as in sectPr
// -----------------------------------------------------------------------
const DOCX_PAGE = {
  width: 12240, // 8.5"
  height: 15840, // 11"
  margin: {
    top: 1440,
    bottom: 1440,
    left: 1440,
    right: 1440,
    header: 720,
    footer: 720,
  },
};

// Usable content width = page width - left margin - right margin
const DOCX_CONTENT_WIDTH = DOCX_PAGE.width - DOCX_PAGE.margin.left - DOCX_PAGE.margin.right; // 9360

// -----------------------------------------------------------------------
// PLACEHOLDER TEXT for empty table rows (rendered in italics)
// -----------------------------------------------------------------------
const PLACEHOLDER_ROW_TEXT = '[To be filled]';

// -----------------------------------------------------------------------
// HELPERS
// -----------------------------------------------------------------------

/**
 * Distributes column widths (DXA) across the full content width according
 * to each column's relative `weight` (default 1). Guarantees the widths
 * sum exactly to `totalWidth` (last column absorbs any rounding remainder).
 * @param {Array<{weight?:number}>} columns
 * @param {number} totalWidth
 * @returns {number[]}
 */
function distributeColumnWidths(columns, totalWidth) {
  const weights = columns.map((c) => (c && c.weight ? c.weight : 1));
  const totalWeight = weights.reduce((a, b) => a + b, 0);
  const widths = weights.map((w) => Math.floor((w / totalWeight) * totalWidth));
  const sum = widths.reduce((a, b) => a + b, 0);
  widths[widths.length - 1] += totalWidth - sum; // fix rounding drift
  return widths;
}

/**
 * Decides whether a table should render real data rows or the italic
 * placeholder row(s), and returns the row set to actually render.
 * @param {Array<Object>|undefined} rows
 * @param {Array<Object>} placeholderRows
 * @returns {{hasData:boolean, rows:Array<Object>}}
 */
function resolveRows(rows, placeholderRows) {
  const hasData = Array.isArray(rows) && rows.length > 0;
  return { hasData, rows: hasData ? rows : placeholderRows };
}

// -----------------------------------------------------------------------
// CODE-BLOCK TOKENS (shaded monospace blocks: repo trees, shell steps, JSON)
// -----------------------------------------------------------------------
const DOCX_CODE = {
  font: 'Consolas',
  fill: 'f5f5f5',
  size: 19, // 9.5pt, matches source
};

// =======================================================================
//  SHARED DOCX PRIMITIVES
// =======================================================================
// Every Part module used to carry its own byte-identical copy of these.
// They live here now so a styling change lands in one place. Anything a
// single Part genuinely renders differently (the PIC Matrix's and Style
// Guide's bespoke table cells) deliberately stays local to that file.
// =======================================================================

const {
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  ShadingType,
  BorderStyle,
  AlignmentType,
} = require('docx');

/** A standard body run: Calibri 10pt. Pass opts to override (bold, italics, color...). */
function run(text, opts = {}) {
  return new TextRun(
    Object.assign({ text: String(text), font: DOCX_FONT, size: DOCX_SIZE.body }, opts),
  );
}

/** A heading run — no explicit size/color, so the paragraph style supplies them. */
function headingRun(text) {
  return new TextRun({ text, font: DOCX_FONT });
}

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [headingRun(text)] });
}

function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [headingRun(text)] });
}

function h3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [headingRun(text)] });
}

/** Italic muted-gray guidance/instruction note that precedes many sections. */
function guidanceNote(text) {
  return new Paragraph({
    spacing: { after: 120, line: 276, lineRule: 'auto' },
    children: [run(text, { italics: true, color: DOCX_COLOR.mutedText })],
  });
}

/** Standard left-aligned body paragraph. */
function bodyPara(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 160, line: 276, lineRule: 'auto' },
    children: [run(text, opts)],
  });
}

/**
 * Justified body paragraph (Project Brief / Style Guide prose).
 * Pass `{ justify: false }` to fall back to the document default alignment.
 */
function bodyParaJustified(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 160, line: 276, lineRule: 'auto' },
    alignment: opts.justify === false ? undefined : AlignmentType.JUSTIFIED,
    children: [run(text, opts)],
  });
}

/** Standard table cell padding (DXA). */
function cellMargins() {
  return { top: 60, bottom: 60, left: 100, right: 100 };
}

/** Solid black border on every edge, inside and out. */
function outerBorderSet() {
  const edge = { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLOR.tableOuterBorder };
  return {
    top: edge,
    bottom: edge,
    left: edge,
    right: edge,
    insideHorizontal: edge,
    insideVertical: edge,
  };
}

/**
 * The three-level bullet definition, identical in every Part. Each Part
 * passes its own `reference` so two merged Parts can't collide.
 * @param {string} reference - the numbering reference id for this Part
 */
function bulletNumberingConfig(reference) {
  return {
    reference,
    levels: [
      {
        level: 0,
        format: 'bullet',
        text: '●',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      },
      {
        level: 1,
        format: 'bullet',
        text: '○',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1440, hanging: 360 } } },
      },
      {
        level: 2,
        format: 'bullet',
        text: '■',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 2160, hanging: 360 } } },
      },
    ],
  };
}

/**
 * Builds a `bullet(text, opts)` bound to one Part's numbering reference, so
 * call sites stay unchanged: `const bullet = makeBullet(BULLET_NUMBERING_REF);`
 * @param {string} reference
 */
function makeBullet(reference) {
  return function bullet(text, opts = {}) {
    return new Paragraph({
      numbering: { reference, level: 0 },
      spacing: { after: 60, line: 276, lineRule: 'auto' },
      children: [run(text, opts)],
    });
  };
}

/** A checkbox-prefixed checklist line (Security Checklist, Post-Deployment Verification). */
function checklistItem(text) {
  return new Paragraph({
    spacing: { after: 80, line: 276, lineRule: 'auto' },
    children: [run('☐ ', {}), run(text, {})],
  });
}

/**
 * A shaded monospace code block. Takes an array of lines (or a single
 * string); every line becomes one paragraph so the shading runs continuous.
 * Empty-string entries render as a blank shaded line.
 * @param {string[]|string} lines
 */
function codeBlock(lines) {
  const arr = Array.isArray(lines) ? lines : [lines];
  return arr.map(
    (line, i) =>
      new Paragraph({
        shading: { fill: DOCX_CODE.fill, type: ShadingType.CLEAR, color: 'auto' },
        spacing: {
          before: i === 0 ? 60 : 0,
          after: i === arr.length - 1 ? 120 : 0,
          line: 240,
          lineRule: 'auto',
        },
        children: [new TextRun({ text: line || ' ', font: DOCX_CODE.font, size: DOCX_CODE.size })],
      }),
  );
}

/**
 * The standard navy-header, zebra-body table used by 11 of the 14 Parts.
 * Odd-numbered body rows take the zebra fill; even rows are left unshaded.
 *
 * @param {Array<{key:string, header:string, weight?:number}>} columns
 * @param {Array<Object>|undefined} rows - caller's data; empty/absent falls back to `placeholderRows`
 * @param {Array<Object>|Object|undefined} placeholderRows - the template's own default row(s)
 * @param {Object} [opts]
 * @param {string[]} [opts.placeholderKeys] - restrict the "[To be filled]" marker to these
 *   column keys (the Change Request Log only marks its `crId` column). Default: every column.
 */
function buildZebraTable(columns, rows, placeholderRows, opts = {}) {
  const widths = distributeColumnWidths(columns, DOCX_CONTENT_WIDTH);
  const hasData = Array.isArray(rows) && rows.length > 0;
  let useRows;
  if (hasData) useRows = rows;
  else if (Array.isArray(placeholderRows)) useRows = placeholderRows;
  else if (placeholderRows == null) useRows = [];
  else useRows = [placeholderRows];
  const isPlaceholder = !hasData;
  const placeholderKeys = opts.placeholderKeys || null;

  const headerRow = new TableRow({
    tableHeader: true,
    children: columns.map(
      (col, idx) =>
        new TableCell({
          width: { size: widths[idx], type: WidthType.DXA },
          shading: { fill: DOCX_COLOR.tableHeaderBg, type: ShadingType.CLEAR, color: 'auto' },
          margins: cellMargins(),
          children: [
            new Paragraph({
              children: [run(col.header, { bold: true, color: DOCX_COLOR.tableHeaderText })],
            }),
          ],
        }),
    ),
  });

  const bodyRows = useRows.map((r, rowIdx) => {
    const shadeFill = rowIdx % 2 === 0 ? null : DOCX_COLOR.zebraFill;
    return new TableRow({
      children: columns.map((col, idx) => {
        const value = r ? r[col.key] : undefined;
        const marks = isPlaceholder && (!placeholderKeys || placeholderKeys.includes(col.key));
        const text =
          value == null || value === '' ? (marks ? PLACEHOLDER_ROW_TEXT : '') : String(value);
        const cellProps = {
          width: { size: widths[idx], type: WidthType.DXA },
          margins: cellMargins(),
          children: [new Paragraph({ children: [run(text)] })],
        };
        if (shadeFill)
          cellProps.shading = { fill: shadeFill, type: ShadingType.CLEAR, color: 'auto' };
        return new TableCell(cellProps);
      }),
    });
  });

  return new Table({
    width: { size: DOCX_CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: widths,
    borders: outerBorderSet(),
    rows: [headerRow, ...bodyRows],
  });
}

module.exports = {
  DOCX_FONT,
  DOCX_SIZE,
  DOCX_COLOR,
  DOCX_PAGE,
  DOCX_CONTENT_WIDTH,
  PLACEHOLDER_ROW_TEXT,
  DOCX_CODE,
  distributeColumnWidths,
  resolveRows,

  // shared docx primitives
  run,
  headingRun,
  h1,
  h2,
  h3,
  guidanceNote,
  bodyPara,
  bodyParaJustified,
  cellMargins,
  outerBorderSet,
  bulletNumberingConfig,
  makeBullet,
  checklistItem,
  codeBlock,
  buildZebraTable,
};
