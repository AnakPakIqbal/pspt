'use strict';
/**
 * =============================================================================
 *  DOCUMENTATION STYLE GUIDE (Part 1 of the Product Documentation Master)
 * =============================================================================
 *
 * Regenerates "Part 1. Documentation Style Guide" exactly as it appears in
 * the master template: a Part-title-only page (no header/footer) followed
 * by the content pages (with the "Style Guide | Confidential" running
 * header and "Product Team / Internal Only / Page X of Y" footer).
 *
 * This module also establishes the reusable "Part" pattern used by every
 * other chapter of the master document:
 *   - Section A: a lone big Heading-1-styled title on its own page, blank
 *     header/footer (`titlePg` trick: the section's "first" header/footer
 *     carries no content, and there is exactly one page in the section).
 *   - Section B: the actual content, repeating the title as a normal
 *     Heading1, with the chapter's running header label and footer.
 * Later chapters (Project Brief, BRD, PRD, ...) reuse `buildPartTitlePage`
 * and `buildChapterHeader`/`buildChapterFooter` from this file.
 *
 * WHO THIS IS FOR
 * ----------------
 * Call `set...()` for whichever tables/paragraphs you want to override.
 * Everything not overridden falls back to the master template's own
 * verbatim default content (this chapter is largely a fixed style
 * reference, so most callers will just call `generate()` as-is or override
 * only `setDocumentSuiteMap` to reflect a real product's part list).
 *
 * HOW TO USE
 * -----------
 *   const StyleGuideSDK = require('./style-guide-docx');
 *   const doc = new StyleGuideSDK();
 *   doc.setHeaderFooterLabels({ productNameLabel: 'Acme Widget' });
 *   await doc.generate('/mnt/user-data/outputs/01-style-guide.docx');
 * =============================================================================
 */

const {
  Document,
  Packer,
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
  Header,
  Footer,
  TabStopType,
  TabStopPosition,
  PageNumber,
} = require('docx');
const fs = require('fs');

const {
  DOCX_COLOR: COLOR,
  DOCX_FONT: FONT,
  DOCX_PAGE: PAGE,
  DOCX_SIZE: SIZE,
  bodyParaJustified: bodyPara,
  distributeColumnWidths,
  h2,
  h3,
  outerBorderSet,
  run,
} = require('./pspt-core');

// -----------------------------------------------------------------------
// This chapter's own page geometry: 0.94" margins, 0.49" header/footer,
// exactly as declared in "3.1 Page Setup" and measured from the XML.
// -----------------------------------------------------------------------
const CHAPTER_MARGIN = 1354; // 0.94"
const CHAPTER_HEADER_FOOTER = 708; // 0.49"
const CHAPTER_CONTENT_WIDTH = PAGE.width - CHAPTER_MARGIN * 2; // 9532

function italicNote(text) {
  return new Paragraph({
    spacing: { before: 120, after: 200, line: 276, lineRule: 'auto' },
    children: [run(text, { italics: true, color: COLOR.mutedText })],
  });
}

function cellBorderSet(color) {
  const edge = { style: BorderStyle.SINGLE, size: 4, color: color || COLOR.tableBorderAlt };
  return { top: edge, bottom: edge, left: edge, right: edge };
}

function headerCell(text, width) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: cellBorderSet(),
    shading: { fill: COLOR.tableHeaderBg, type: ShadingType.CLEAR, color: 'auto' },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [
      new Paragraph({ children: [run(text, { bold: true, color: COLOR.tableHeaderText })] }),
    ],
  });
}

/** Body cell. `shadeFill` null/undefined = no shading (white, row 1 style); pass a hex to shade. */
function bodyCell(content, width, opts = {}) {
  const cellChildren = Array.isArray(content)
    ? content
    : [
        new Paragraph({
          children: Array.isArray(opts.runs)
            ? opts.runs
            : [run(content == null ? '' : String(content), opts.runOpts || {})],
        }),
      ];
  const cellProps = {
    width: { size: width, type: WidthType.DXA },
    borders: cellBorderSet(),
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: cellChildren,
  };
  if (opts.shadeFill) {
    cellProps.shading = { fill: opts.shadeFill, type: ShadingType.CLEAR, color: 'auto' };
  }
  if (opts.valign) cellProps.verticalAlign = opts.valign;
  return new TableCell(cellProps);
}

/**
 * Zebra table matching this chapter's exact style: navy header, cbd5e0
 * borders, row 1 unshaded (white), row 2+ alternating edf2f7/white.
 * @param {Array<{key:string, header:string, weight?:number}>} columns
 * @param {Array<Object>} rows
 * @param {Object} [opts] - {width, cellRenderers: {key: (value,row)=>Paragraph[]}}
 */
function buildStyleTable(columns, rows, opts = {}) {
  const tableWidth = opts.width || CHAPTER_CONTENT_WIDTH;
  const widths = distributeColumnWidths(columns, tableWidth);

  const headerRow = new TableRow({
    tableHeader: true,
    children: columns.map((col, idx) => headerCell(col.header, widths[idx])),
  });

  const bodyRows = rows.map((r, rowIdx) => {
    // Row 1 (rowIdx 0) unshaded/white; subsequent rows alternate edf2f7/white.
    const shadeFill = rowIdx === 0 ? null : rowIdx % 2 === 0 ? null : COLOR.zebraFill;
    return new TableRow({
      children: columns.map((col, idx) => {
        const renderer = opts.cellRenderers && opts.cellRenderers[col.key];
        const value = r[col.key];
        if (renderer) {
          return bodyCell(renderer(value, r), widths[idx], { shadeFill });
        }
        return bodyCell(value, widths[idx], { shadeFill, runOpts: col.runOpts });
      }),
    });
  });

  return new Table({
    width: { size: tableWidth, type: WidthType.DXA },
    columnWidths: widths,
    borders: outerBorderSet(),
    rows: [headerRow, ...bodyRows],
  });
}

// =============================================================================
// DEFAULT CONTENT — verbatim from the master template.
// =============================================================================

const DEFAULT_FONT_FAMILIES = [
  { role: 'Title', font: 'Calibri (Bold)' },
  { role: 'Headings', font: 'Calibri (Bold)' },
  { role: 'Body text', font: 'Calibri (Regular)' },
  { role: 'Code / technical values', font: 'Consolas', mono: true },
];

const DEFAULT_TYPE_SCALE = [
  {
    element: 'Document Title',
    size: '28 pt',
    weight: 'Bold',
    color: 'Navy #1F3B57',
    wordStyle: 'Title',
  },
  {
    element: 'H1 (Part title)',
    size: '18 pt',
    weight: 'Bold',
    color: 'Navy #1F3B57',
    wordStyle: 'Heading 1',
  },
  {
    element: 'H2 (Section)',
    size: '15 pt',
    weight: 'Bold',
    color: 'Navy #1F3B57',
    wordStyle: 'Heading 2',
  },
  {
    element: 'H3 (Subsection)',
    size: '12 pt',
    weight: 'Bold',
    color: 'Blue #2F6690',
    wordStyle: 'Heading 3',
  },
  {
    element: 'Body text',
    size: '10 pt',
    weight: 'Regular',
    color: 'Gray #44494D',
    wordStyle: 'Normal',
  },
  {
    element: 'Table header',
    size: '10 pt',
    weight: 'Bold',
    color: 'White on Navy',
    wordStyle: 'Table Header',
  },
  {
    element: 'Caption/footnote',
    size: '8 pt',
    weight: 'Regular / Italic',
    color: 'Muted #6B7280',
    wordStyle: 'Caption',
  },
];

const DEFAULT_FORMATTING_RULES = [
  {
    category: 'Bold',
    rule: 'Use for UI labels, key terms on first use, and table headers — never whole sentences.',
  },
  {
    category: 'Italics',
    rule: 'Use for placeholders, captions, and citations — not for emphasis in body text.',
  },
  { category: 'Underlining', rule: 'Avoid except for actual hyperlinks.' },
  {
    category: 'Alignment',
    rule: 'Justify body text; left-align table content (except pricing); centre-align document titles and cover page elements.',
  },
];

const DEFAULT_COLOR_PALETTE = [
  { name: 'Navy (Primary)', hex: '#1F3B57', rgb: '31, 59, 87', use: 'Titles, H1/H2, header text' },
  { name: 'Blue (Secondary)', hex: '#2F6690', rgb: '47, 102, 144', use: 'H3, subheads' },
  { name: 'Accent', hex: '#3E92CC', rgb: '62, 146, 204', use: 'Links, highlights, rules' },
  {
    name: 'Light Background',
    hex: '#EDF2F7',
    rgb: '237, 242, 247',
    use: 'Alternating rows, callouts',
  },
  { name: 'Body Gray', hex: '#44494D', rgb: '68, 73, 77', use: 'Body text' },
  {
    name: 'Muted Gray',
    hex: '#6B7280',
    rgb: '107, 114, 128',
    use: 'Captions, footers, placeholders',
  },
];

const DEFAULT_STATUS_COLORS = [
  { status: 'Success / Done', hex: '#2E7D32', use: 'Completed milestones, passed QA' },
  { status: 'Warning / At Risk', hex: '#B8860B', use: 'Slipping timelines need attention' },
  { status: 'Error / Blocked', hex: '#C0392B', use: 'Blockers, failed requirements' },
];

const DEFAULT_HEADER_FOOTER_GUIDELINES = [
  { element: 'Header, left', guideline: 'Application Name/Logo, bold, 8 pt, Navy.' },
  {
    element: 'Header, right',
    guideline: 'Document title and Confidentiality label, 8 pt, Muted Grey/red.',
  },
  {
    element: 'Header rule',
    guideline:
      'A thin 0.5 pt Border-Grey rule sits under the header. No page numbers in the header.',
  },
  { element: 'Footer, left', guideline: 'Product Team, 8 pt, Muted Grey.' },
  { element: 'Footer, right', guideline: 'Classification (Public, Internal), 8 pt, Muted Grey.' },
  { element: 'Footer, middle bottom', guideline: "'Page X of Y', 8 pt, Muted Grey." },
  {
    element: 'Footer rule',
    guideline: 'A thin 0.5 pt Border-Grey rule sits above the footer, mirroring the header rule.',
  },
  {
    element: 'Suppression',
    guideline:
      "Headers/footers are suppressed on each Part's opening page, so the Part title page reads cleanly.",
  },
];

const DEFAULT_PAGE_SETUP = [
  { setting: 'Page size', value: 'US Letter (8.5" × 11")' },
  { setting: 'Margins', value: '0.94" top/bottom, 0.94" left/right' },
  { setting: 'Line spacing', value: '1.15–1.3x for body paragraphs' },
];

const DEFAULT_SUITE_MAP = [
  { part: '2', document: 'Project Brief', audience: 'Sponsors, stakeholders' },
  {
    part: '3',
    document: 'Business Requirements Document (BRD) — incl. Business Process & Use Cases',
    audience: 'Business stakeholders, Eng, QA',
  },
  { part: '4', document: 'Product Requirements Document (PRD)', audience: 'Product, Eng, Design' },
  { part: '5', document: 'Software Requirements Specification (SRS)', audience: 'Engineering, QA' },
  {
    part: '6',
    document: 'Technical Documentation — System Architecture, Data Model, API Spec, Security',
    audience: 'Engineering, Security/Compliance',
  },
  { part: '7', document: 'UI/UX Documentation', audience: 'Design, Engineering' },
  { part: '8', document: 'Test Plan / QA Criteria', audience: 'QA, Engineering' },
  { part: '9', document: 'Deployment Guide', audience: 'DevOps, Engineering' },
  { part: '10', document: 'User Manual', audience: 'End users, Support' },
  { part: '11', document: 'Changelog', audience: 'Everyone' },
  { part: '12', document: 'Change Request Log', audience: 'Everyone' },
  { part: '13', document: 'Glossary', audience: 'Everyone' },
  { part: '14', document: 'Appendices', audience: 'Everyone' },
];

// =============================================================================
// SECTION GUIDE
// =============================================================================

const SECTION_GUIDE = [
  {
    method: 'setHeaderFooterLabels',
    purpose: 'Overrides the running header/footer product name label.',
    example: { productNameLabel: 'Acme Widget' },
  },
  {
    method: 'setFontFamilies',
    purpose: 'Overrides the "1.1 Font Families" table.',
    example: [{ role: 'Title', font: 'Inter (Bold)' }],
  },
  {
    method: 'setTypeScale',
    purpose: 'Overrides the "1.2 Type Scale" table.',
    example: [
      {
        element: 'Document Title',
        size: '28 pt',
        weight: 'Bold',
        color: 'Navy #1F3B57',
        wordStyle: 'Title',
      },
    ],
  },
  {
    method: 'setFormattingRules',
    purpose: 'Overrides the "1.3 Text Formatting Rules" table.',
    example: [{ category: 'Bold', rule: 'Use for key terms.' }],
  },
  {
    method: 'setColorPalette',
    purpose: 'Overrides the "2. Colour Palette" table.',
    example: [{ name: 'Navy', hex: '#1F3B57', rgb: '31, 59, 87', use: 'Titles' }],
  },
  {
    method: 'setStatusColors',
    purpose: 'Overrides the "2.1 Semantic Status Colours" table.',
    example: [{ status: 'Success', hex: '#2E7D32', use: 'Completed milestones' }],
  },
  {
    method: 'setHeaderFooterGuidelines',
    purpose: 'Overrides the "3. Header & Footer Guidelines" table.',
    example: [{ element: 'Header, left', guideline: 'Logo, bold, 8pt.' }],
  },
  {
    method: 'setPageSetup',
    purpose: 'Overrides the "3.1 Page Setup" table.',
    example: [{ setting: 'Page size', value: 'US Letter' }],
  },
  {
    method: 'setDocumentSuiteMap',
    purpose:
      'Overrides the "4. Document Suite Map" table — the real list of parts for this product.',
    example: [{ part: '2', document: 'Project Brief', audience: 'Sponsors' }],
  },
];

// =============================================================================
// MAIN SDK CLASS
// =============================================================================

class StyleGuideSDK {
  constructor() {
    this.data = {};
  }

  static sectionGuide() {
    return SECTION_GUIDE;
  }

  /** @param {Object} p @param {string} [p.productNameLabel] - defaults to "Product Name/Logo". */
  setHeaderFooterLabels(p = {}) {
    this.data.headerFooterLabels = p;
    return this;
  }

  /** @param {Array<{role:string, font:string}>} rows */
  setFontFamilies(rows) {
    this.data.fontFamilies = rows;
    return this;
  }

  /** @param {Array<{element:string, size:string, weight:string, color:string, wordStyle:string}>} rows */
  setTypeScale(rows) {
    this.data.typeScale = rows;
    return this;
  }

  /** @param {Array<{category:string, rule:string}>} rows */
  setFormattingRules(rows) {
    this.data.formattingRules = rows;
    return this;
  }

  /** @param {Array<{name:string, hex:string, rgb:string, use:string}>} rows */
  setColorPalette(rows) {
    this.data.colorPalette = rows;
    return this;
  }

  /** @param {Array<{status:string, hex:string, use:string}>} rows */
  setStatusColors(rows) {
    this.data.statusColors = rows;
    return this;
  }

  /** @param {Array<{element:string, guideline:string}>} rows */
  setHeaderFooterGuidelines(rows) {
    this.data.headerFooterGuidelines = rows;
    return this;
  }

  /** @param {Array<{setting:string, value:string}>} rows */
  setPageSetup(rows) {
    this.data.pageSetup = rows;
    return this;
  }

  /** @param {Array<{part:string, document:string, audience:string}>} rows */
  setDocumentSuiteMap(rows) {
    this.data.suiteMap = rows;
    return this;
  }

  // =========================================================================
  // BUILD
  // =========================================================================

  _buildPartTitlePage() {
    // A single Heading1 paragraph is the entire content of this section.
    // Note: no explicit run size here — leaving it unset lets the run
    // inherit the Heading1 paragraph style's size (18pt) instead of being
    // clobbered by run()'s default body size.
    return [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: 'Documentation Style Guide', font: FONT })],
      }),
    ];
  }

  _buildContent() {
    const children = [];
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: 'Documentation Style Guide', font: FONT })],
      }),
    );

    // 1. Typography
    children.push(h2('1. Typography'));
    children.push(
      bodyPara(
        'Use a single heading font and a single body font across all product documentation. Do not mix additional fonts, including in tables or callouts.',
      ),
    );

    children.push(h3('1.1 Font Families'));
    const fontRows =
      this.data.fontFamilies && this.data.fontFamilies.length
        ? this.data.fontFamilies
        : DEFAULT_FONT_FAMILIES;
    children.push(
      buildStyleTable(
        [
          { key: 'role', header: 'Role', weight: 1 },
          { key: 'font', header: 'Font', weight: 2.9 },
        ],
        fontRows,
        {
          cellRenderers: {
            font: (value, row) => [
              new Paragraph({ children: [run(value, row.mono ? { font: 'Consolas' } : {})] }),
            ],
          },
        },
      ),
    );

    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('1.2 Type Scale (sizes & weight)'));
    children.push(
      bodyPara(
        'Sizes below are point size, matched to the Word style used. Keep line spacing at 1.15–1.3x for body copy.',
      ),
    );
    const typeScaleRows =
      this.data.typeScale && this.data.typeScale.length ? this.data.typeScale : DEFAULT_TYPE_SCALE;
    children.push(
      buildStyleTable(
        [
          { key: 'element', header: 'Element', weight: 1.4 },
          { key: 'size', header: 'Size', weight: 0.7 },
          { key: 'weight', header: 'Weight', weight: 0.9 },
          { key: 'color', header: 'Color', weight: 1.2 },
          { key: 'wordStyle', header: 'Word Style', weight: 1.1 },
        ],
        typeScaleRows,
      ),
    );

    children.push(
      italicNote(
        "This master document nests Part titles as Heading 1, so each part's own top-level sections shift down to Heading 2, and their subsections to Heading 3 — one level deeper than in standalone single-topic templates.",
      ),
    );

    children.push(h3('1.3 Text Formatting Rules'));
    const formattingRows =
      this.data.formattingRules && this.data.formattingRules.length
        ? this.data.formattingRules
        : DEFAULT_FORMATTING_RULES;
    children.push(
      buildStyleTable(
        [
          { key: 'category', header: 'Category', weight: 1 },
          { key: 'rule', header: 'Rule', weight: 2.9 },
        ],
        formattingRows,
      ),
    );

    // 2. Colour Palette
    children.push(h2('2. Colour Palette'));
    children.push(
      bodyPara(
        'These colours are used throughout this document, so switching between sections feels consistent.',
      ),
    );
    const paletteRows =
      this.data.colorPalette && this.data.colorPalette.length
        ? this.data.colorPalette
        : DEFAULT_COLOR_PALETTE;
    children.push(
      buildStyleTable(
        [
          { key: 'name', header: 'Name', weight: 1.1 },
          { key: 'hex', header: 'Hex', weight: 0.9 },
          { key: 'rgb', header: 'RGB', weight: 1 },
          { key: 'use', header: 'Use', weight: 1.5 },
        ],
        paletteRows,
        {
          cellRenderers: {
            hex: (value) => [
              new Paragraph({
                shading: undefined,
                children: [run(value)],
              }),
            ],
          },
        },
      ),
    );

    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('2.1 Semantic Status Colours'));
    children.push(bodyPara('Use only for status labels or callouts — never as decoration.'));
    const statusRows =
      this.data.statusColors && this.data.statusColors.length
        ? this.data.statusColors
        : DEFAULT_STATUS_COLORS;
    children.push(
      buildStyleTable(
        [
          { key: 'status', header: 'Status', weight: 1.1 },
          { key: 'hex', header: 'Hex', weight: 0.9 },
          { key: 'use', header: 'Example use', weight: 2.5 },
        ],
        statusRows,
      ),
    );

    // 3. Header & Footer Guidelines
    children.push(h2('3. Header & Footer Guidelines'));
    const hfRows =
      this.data.headerFooterGuidelines && this.data.headerFooterGuidelines.length
        ? this.data.headerFooterGuidelines
        : DEFAULT_HEADER_FOOTER_GUIDELINES;
    children.push(
      buildStyleTable(
        [
          { key: 'element', header: 'Element', weight: 1.2 },
          { key: 'guideline', header: 'Guidelines', weight: 2.8 },
        ],
        hfRows,
      ),
    );

    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('3.1 Page Setup'));
    const pageSetupRows =
      this.data.pageSetup && this.data.pageSetup.length ? this.data.pageSetup : DEFAULT_PAGE_SETUP;
    children.push(
      buildStyleTable(
        [
          { key: 'setting', header: 'Setting', weight: 1.2 },
          { key: 'value', header: 'Value', weight: 2.8 },
        ],
        pageSetupRows,
      ),
    );

    // 4. Document Suite Map
    children.push(h2('4. Document Suite Map'));
    const suiteRows =
      this.data.suiteMap && this.data.suiteMap.length ? this.data.suiteMap : DEFAULT_SUITE_MAP;
    children.push(
      buildStyleTable(
        [
          { key: 'part', header: 'Part', weight: 0.6 },
          { key: 'document', header: 'Document', weight: 2 },
          { key: 'audience', header: 'Primary Audience', weight: 1.8 },
        ],
        suiteRows,
        {
          cellRenderers: {
            part: (value) => [
              new Paragraph({ alignment: AlignmentType.CENTER, children: [run(value)] }),
            ],
          },
        },
      ),
    );

    return children;
  }

  // -- Header/footer builders (reusable pattern for future Parts) ---------

  _blankHeader() {
    return new Header({ children: [new Paragraph({ children: [] })] });
  }

  _blankFooter() {
    return new Footer({ children: [new Paragraph({ children: [] })] });
  }

  /** The chapter's running header: "{ProductName/Logo}  ...  {ChapterLabel}|Confidential" */
  _chapterHeader(chapterLabel) {
    const labels = this.data.headerFooterLabels || {};
    return new Header({
      children: [
        new Paragraph({
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR.tableBorder, space: 4 },
          },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            run(`${labels.productNameLabel || 'Product Name/Logo'}\t`, {
              bold: true,
              size: SIZE.headerFooter,
              color: COLOR.headerFooterText,
            }),
            run(chapterLabel, {
              bold: true,
              size: SIZE.headerFooter,
              color: COLOR.headerFooterText,
            }),
            run('|', { size: SIZE.headerFooter, color: COLOR.mutedText }),
            run('Confidential', { bold: true, size: SIZE.headerFooter, color: COLOR.confidential }),
          ],
        }),
      ],
    });
  }

  /** The chapter's running footer: "Product Team ... Internal Only" + centered "Page X of Y". */
  _chapterFooter() {
    return new Footer({
      children: [
        new Paragraph({
          border: {
            top: { style: BorderStyle.SINGLE, size: 4, color: COLOR.tableBorder, space: 4 },
          },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            run('Product Team\t', { size: SIZE.headerFooter, color: COLOR.mutedText }),
            run('Internal Only', { size: SIZE.headerFooter, color: COLOR.mutedText }),
          ],
        }),
        new Paragraph({
          border: {
            top: { style: BorderStyle.SINGLE, size: 4, color: COLOR.tableBorder, space: 4 },
          },
          alignment: AlignmentType.CENTER,
          children: [
            run('Page ', { size: SIZE.headerFooter, color: COLOR.mutedText }),
            new TextRun({
              children: [PageNumber.CURRENT],
              size: SIZE.headerFooter,
              color: COLOR.mutedText,
              font: FONT,
            }),
            run(' of ', { size: SIZE.headerFooter, color: COLOR.mutedText }),
            new TextRun({
              children: [PageNumber.TOTAL_PAGES],
              size: SIZE.headerFooter,
              color: COLOR.mutedText,
              font: FONT,
            }),
          ],
        }),
      ],
    });
  }

  /** Builds the underlying `docx` Document object (two sections: title page + content). */
  /**
   * The raw `docx` Document options for this Part — styles, numbering, and
   * its two sections (title page + content). `generate-master.js` reads this
   * so every Part can be concatenated into one Document instead of merged
   * as separate .docx files.
   */
  documentOptions() {
    return {
      styles: {
        default: { document: { run: { font: FONT, size: SIZE.body } } },
        paragraphStyles: [
          {
            id: 'Heading1',
            name: 'Heading 1',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            run: { font: FONT, size: SIZE.partTitle, bold: true, color: COLOR.h1 },
            paragraph: { spacing: { before: 0, after: 200 } },
          },
          {
            id: 'Heading2',
            name: 'Heading 2',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            run: { font: FONT, size: SIZE.h1, bold: true, color: COLOR.h1 },
            paragraph: { spacing: { before: 320, after: 160 } },
          },
          {
            id: 'Heading3',
            name: 'Heading 3',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            run: { font: FONT, size: SIZE.h2, bold: true, color: COLOR.h3 },
            paragraph: { spacing: { before: 240, after: 120 } },
          },
          {
            id: 'Heading4',
            name: 'Heading 4',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            run: { font: FONT, size: SIZE.h4, bold: true, color: COLOR.h1Rendered },
            paragraph: { spacing: { before: 200, after: 100 } },
          },
        ],
      },
      sections: [
        {
          properties: {
            titlePage: true,
            page: {
              size: { width: PAGE.width, height: PAGE.height },
              margin: {
                top: CHAPTER_MARGIN,
                bottom: CHAPTER_MARGIN,
                left: CHAPTER_MARGIN,
                right: CHAPTER_MARGIN,
                header: PAGE.margin.header,
                footer: PAGE.margin.footer,
              },
            },
          },
          headers: { default: this._blankHeader(), first: this._blankHeader() },
          footers: { default: this._blankFooter(), first: this._blankFooter() },
          children: this._buildPartTitlePage(),
        },
        {
          properties: {
            page: {
              size: { width: PAGE.width, height: PAGE.height },
              margin: {
                top: PAGE.margin.top,
                bottom: PAGE.margin.bottom,
                left: PAGE.margin.left,
                right: PAGE.margin.right,
                header: PAGE.margin.header,
                footer: PAGE.margin.footer,
              },
            },
          },
          headers: { default: this._chapterHeader('Style Guide') },
          footers: { default: this._chapterFooter() },
          children: this._buildContent(),
        },
      ],
    };
  }

  toDocument() {
    return new Document(this.documentOptions());
  }

  async toBuffer() {
    return Packer.toBuffer(this.toDocument());
  }

  async generate(outputPath) {
    const buf = await this.toBuffer();
    fs.writeFileSync(outputPath, buf);
    return outputPath;
  }
}

module.exports = StyleGuideSDK;
module.exports.SECTION_GUIDE = SECTION_GUIDE;
module.exports.buildStyleTable = buildStyleTable;
module.exports.CHAPTER_MARGIN = CHAPTER_MARGIN;
module.exports.CHAPTER_HEADER_FOOTER = CHAPTER_HEADER_FOOTER;
