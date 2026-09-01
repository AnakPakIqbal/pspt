'use strict';
/**
 * =============================================================================
 *  DOCUMENTATION FLOW & PIC MATRIX — DOCX SDK
 * =============================================================================
 *
 * Regenerates the "Name of Product Documentation Flow & PIC Matrix" Word
 * template with byte-for-byte matched STYLE: Calibri throughout, the same
 * heading colors/sizes, the same navy-header/white-text + zebra-striped
 * table style, the same US-Letter page size & margins, and the same
 * header/footer layout as the source template.
 *
 * WHO THIS IS FOR
 * ----------------
 * This file is meant to be driven by another AI (or a human) that has the
 * actual content. You do NOT need to touch styling, fonts, colors, margins,
 * or table borders — that is all baked in and matches the template exactly.
 * You only call the `set...()` method for whichever section(s) you have
 * content for.
 *
 * HOW TO USE (quick start)
 * -------------------------
 *   const PicMatrixSDK = require('./pic-matrix-docx');
 *   const doc = new PicMatrixSDK();
 *
 *   doc.setCoverInfo({
 *     productName: 'Acme Widget',
 *     productType: 'SaaS',
 *     applicationEngineer: 'Jane Doe',
 *     productLead: 'John Smith',
 *     status: 'Live',
 *     writer: 'Jane Doe',
 *     checker: { name: 'Atya Salma', email: 'atya.salma@point-star.com' },
 *     approver: { name: 'Muhammad Iqbal', email: 'iqbal@point-star.com' },
 *     lastUpdate: '2026-08-20',
 *     latestHistory: 'Updated document format Aug 20, 2026',
 *   });
 *   doc.setTeam([ ... ]);
 *   doc.setPicMatrix([ ... ]);
 *   doc.setStorageGuide([ ... ]);
 *
 *   await doc.generate('/mnt/user-data/outputs/pic-matrix.docx');
 *
 * IMPORTANT FOR AI CALLERS
 * -------------------------
 * - Every `set...()` method has a JSDoc comment directly above it describing
 *   (a) what real-world information belongs there and (b) the exact shape
 *   of the object/array it expects. READ IT before calling.
 * - You do not have to call every method. Any section you skip renders with
 *   the same "[To be filled]" placeholder text the original template
 *   implies, in italics, so a human reviewer can see what's still missing.
 *   Calling `generate()` with zero setter calls reproduces the blank
 *   template structure exactly.
 * - Call `PicMatrixSDK.sectionGuide()` (static, no instance needed) for a
 *   machine-readable JSON array of every section: method name, purpose,
 *   and an example payload.
 * - Tables: pass an array of plain objects. Keys are given in each method's
 *   JSDoc. Omit a method entirely (don't call it) to get the placeholder
 *   row(s).
 * - The PIC Matrix and Storage Guide tables come with the template's
 *   default row content already pre-populated (Style Guide, Project Brief,
 *   BRD, PRD, etc. / Shared Drive, Folder Path, etc.) — call the setter
 *   only if you want to override or extend those rows; otherwise the
 *   original template rows are used verbatim.
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
  VerticalAlign,
  Header,
  Footer,
  TabStopType,
  TabStopPosition,
  ExternalHyperlink,
  PageBreak,
} = require('docx');
const fs = require('fs');

const {
  DOCX_COLOR: COLOR,
  DOCX_CONTENT_WIDTH: CONTENT_WIDTH,
  DOCX_FONT: FONT,
  DOCX_PAGE: PAGE,
  DOCX_SIZE: SIZE,
  PLACEHOLDER_ROW_TEXT,
  bodyPara,
  distributeColumnWidths,
  outerBorderSet,
  resolveRows,
  run,
} = require('./pspt-core');

// =============================================================================
// LOW-LEVEL BUILDERS (not exported — internal rendering helpers)
// =============================================================================

/** A bold, colored, underlined section heading with the template's bottom rule. */
function sectionHeading(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: COLOR.headingRule, space: 4 },
    },
    spacing: { before: 400, after: 200 },
    children: [run(text, { size: SIZE.h1, bold: false, color: COLOR.h1Rendered })],
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

/** Thin blue rule used as a plain divider (no heading text), matching the
 *  spacer rules seen between sections in the source template. */
function dividerRule() {
  return new Paragraph({
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: COLOR.headingRule, space: 4 },
    },
    spacing: { before: 200, after: 200 },
    children: [],
  });
}

function cellBorderSet() {
  const edge = { style: BorderStyle.SINGLE, size: 4, color: COLOR.tableBorder };
  return { top: edge, bottom: edge, left: edge, right: edge };
}

/** Navy header cell with centered, bold, white text — used for every table's header row. */
function headerCell(text, width, opts = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: cellBorderSet(),
    shading: { fill: COLOR.tableHeaderBg, type: ShadingType.CLEAR, color: 'auto' },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [
      new Paragraph({
        alignment: opts.align || AlignmentType.CENTER,
        children: [run(text, { bold: true, color: COLOR.tableHeaderText, size: SIZE.tableHeader })],
      }),
    ],
  });
}

/** A standard zebra-striped body cell. `shadeFill` controls the row's zebra color. */
function bodyCell(text, width, opts = {}) {
  const italics = Boolean(opts.placeholder);
  const isEmpty = text == null || text === '';
  let value;
  if (isEmpty) {
    value = opts.placeholder ? PLACEHOLDER_ROW_TEXT : '';
  } else {
    value = String(text);
  }
  const lines = value.split(/\r?\n+/);
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: cellBorderSet(),
    shading: { fill: opts.shadeFill || COLOR.zebraFillAlt, type: ShadingType.CLEAR, color: 'auto' },
    margins: { top: 90, bottom: 90, left: 100, right: 100 },
    verticalAlign: VerticalAlign.CENTER,
    children: lines.map(
      (line) =>
        new Paragraph({
          children: [
            run(line, {
              italics,
              bold: Boolean(opts.bold),
              color: italics ? COLOR.placeholderText : opts.color || COLOR.bodyText,
            }),
          ],
        }),
    ),
  });
}

/**
 * Builds a zebra-striped table matching the template exactly: navy header
 * row with centered white bold text, light gray-blue (d0d5dd) borders,
 * alternating edf2f7 / white row fills, columns proportioned across the
 * full page content width.
 * @param {Array<{key:string, header:string, weight?:number, bold?:boolean}>} columns
 * @param {Array<Object>|undefined} rows
 * @param {Array<Object>} placeholderRows
 * @param {Object} [opts]
 * @param {number} [opts.width] - table width in DXA (defaults to full content width)
 * @param {number} [opts.firstColBoldKey] - key of a column whose cells should render bold (row labels)
 */
function buildZebraTable(columns, rows, placeholderRows, opts = {}) {
  const { hasData, rows: useRows } = resolveRows(rows, placeholderRows);
  const tableWidth = opts.width || CONTENT_WIDTH;
  const widths = distributeColumnWidths(columns, tableWidth);
  // Only render italic "empty" styling when we truly have nothing (no rows
  // AND no default content supplied via placeholderRows-as-real-data, e.g.
  // the template's own default PIC matrix/storage rows are real content,
  // not placeholders, so `opts.defaultsAreRealContent` suppresses italics).
  const treatAsPlaceholder = !hasData && !opts.defaultsAreRealContent;

  const headerRow = new TableRow({
    tableHeader: true,
    children: columns.map((col, idx) => headerCell(col.header, widths[idx])),
  });

  const bodyRows = (useRows || []).map((r, rowIdx) => {
    const shadeFill = rowIdx % 2 === 0 ? COLOR.zebraFill : COLOR.zebraFillAlt;
    return new TableRow({
      children: columns.map((col, idx) =>
        bodyCell(r ? r[col.key] : undefined, widths[idx], {
          placeholder: treatAsPlaceholder,
          shadeFill,
          bold: Boolean(col.boldValues),
        }),
      ),
    });
  });

  return new Table({
    width: { size: tableWidth, type: WidthType.DXA },
    columnWidths: widths,
    borders: outerBorderSet(),
    rows: [headerRow, ...bodyRows],
  });
}

/**
 * Builds the label/value "identity" style table used on the cover page:
 * bold navy labels in column 1, values in column 2, uniform light-blue fill
 * across every cell (no zebra), matching the cover-info table exactly.
 * @param {Array<{label:string, value:string|{name:string,email:string}|null, isLink?:boolean}>} rows
 */
function buildInfoTable(rows) {
  const labelWidth = 2200;
  const valueWidth = 4800;
  const tableWidth = labelWidth + valueWidth;

  const trRows = rows.map((r) => {
    let valueChildren;
    if (r.isLink && r.value && typeof r.value === 'object') {
      const isPlaceholder = !r.value.name;
      if (isPlaceholder) {
        valueChildren = [
          run(PLACEHOLDER_ROW_TEXT, {
            italics: true,
            color: COLOR.placeholderText,
            size: SIZE.body,
          }),
        ];
      } else {
        valueChildren = [
          new ExternalHyperlink({
            link: r.value.email ? `mailto:${r.value.email}` : '#',
            children: [
              new TextRun({
                text: r.value.name,
                font: FONT,
                size: SIZE.body,
                color: COLOR.hyperlink,
                underline: {},
              }),
            ],
          }),
        ];
      }
    } else {
      const isEmpty = r.value == null || String(r.value).trim() === '';
      valueChildren = [
        run(isEmpty ? '' : String(r.value), {
          size: SIZE.body,
          italics: false,
        }),
      ];
    }
    return new TableRow({
      children: [
        new TableCell({
          width: { size: labelWidth, type: WidthType.DXA },
          borders: cellBorderSet(),
          shading: { fill: COLOR.coverInfoFill, type: ShadingType.CLEAR, color: 'auto' },
          margins: { top: 90, bottom: 90, left: 120, right: 120 },
          children: [
            new Paragraph({
              children: [
                run(r.label, { bold: true, color: COLOR.coverLabelText, size: SIZE.labelCell }),
              ],
            }),
          ],
        }),
        new TableCell({
          width: { size: valueWidth, type: WidthType.DXA },
          borders: cellBorderSet(),
          shading: { fill: COLOR.coverInfoFill, type: ShadingType.CLEAR, color: 'auto' },
          margins: { top: 90, bottom: 90, left: 120, right: 120 },
          children: [new Paragraph({ children: valueChildren })],
        }),
      ],
    });
  });

  return new Table({
    width: { size: tableWidth, type: WidthType.DXA },
    columnWidths: [labelWidth, valueWidth],
    borders: outerBorderSet(),
    alignment: AlignmentType.CENTER,
    rows: trRows,
  });
}

/**
 * A numbered "1. / a. / i." style block, used by the Team responsibilities
 * column. `numberingRef` must be a unique numbering instance reference per
 * table cell/row so the "1." restarts fresh for each team member instead of
 * continuing to count across the whole table.
 */
function numberedBlock(items, numberingRef) {
  const ref = numberingRef || 'pic-matrix-numbering';
  const paras = [];
  (items || []).forEach((item) => {
    const title = typeof item === 'string' ? item : item.title;
    const highlight = typeof item === 'object' ? item.highlight : null;

    if (highlight && title.includes(highlight)) {
      const idx = title.indexOf(highlight);
      const before = title.slice(0, idx);
      const after = title.slice(idx + highlight.length);
      paras.push(
        new Paragraph({
          numbering: { reference: ref, level: 0 },
          spacing: { after: 160, line: 276, lineRule: 'auto' },
          alignment: AlignmentType.JUSTIFIED,
          children: [
            run(before, { size: SIZE.body }),
            run(highlight, { bold: true, color: '0000ff', size: SIZE.body }),
            run(after, { size: SIZE.body }),
          ],
        }),
      );
    } else {
      paras.push(
        new Paragraph({
          numbering: { reference: ref, level: 0 },
          spacing: { after: 160, line: 276, lineRule: 'auto' },
          alignment: AlignmentType.JUSTIFIED,
          children: [run(title, { size: SIZE.body })],
        }),
      );
    }

    if (item && Array.isArray(item.children)) {
      item.children.forEach((sub) => {
        paras.push(
          new Paragraph({
            numbering: { reference: ref, level: 1 },
            spacing: { after: 160, line: 276, lineRule: 'auto' },
            alignment: AlignmentType.JUSTIFIED,
            children: [run(sub, { size: SIZE.body })],
          }),
        );
      });
    }
  });
  return paras;
}

// =============================================================================
// DEFAULT TEMPLATE ROWS — used whenever the caller does not override a
// table via its setter, so the doc reproduces the original template's
// standard PIC matrix / storage-guide content verbatim.
// =============================================================================

const DEFAULT_PIC_MATRIX_ROWS = [
  {
    docType: 'Style Guide',
    maker: 'PM & Dev Lead',
    checker: 'PM & Dev Manager',
    approver: 'PM & Dev Manager',
  },
  {
    docType: 'Project Brief',
    maker: 'PM & Dev Lead',
    checker: 'PM & Dev Manager',
    approver: 'PM & Dev Manager',
  },
  {
    docType: 'BRD (incl. Business Process & Use Cases)',
    maker: 'PM & Dev Lead',
    checker: 'Application Engineer',
    approver: 'PM & Dev Manager',
  },
  {
    docType: 'PRD',
    maker: 'Application Engineer',
    checker: 'PM & Dev Lead',
    approver: 'PM & Dev Manager',
  },
  {
    docType: 'Software Requirements Specification (SRS)',
    maker: 'Application Engineer',
    checker: 'PM & Dev Lead',
    approver: 'PM & Dev Manager',
  },
  {
    docType: 'Technical Documentation',
    maker: 'Application Engineer',
    checker: 'PM & Dev Lead',
    approver: 'PM & Dev Manager',
  },
  {
    docType: 'UI/UX Documentation',
    maker: 'Application Engineer',
    checker: 'PM & Dev Lead',
    approver: 'PM & Dev Manager',
  },
  {
    docType: 'UAT',
    maker: 'PM & Dev Lead',
    checker: 'Application Engineer',
    approver: 'PM & Dev Manager',
  },
  {
    docType: 'Deployment Guide',
    maker: 'Application Engineer',
    checker: 'PM & Dev Lead',
    approver: 'PM & Dev Manager',
  },
  {
    docType: 'User Manual',
    maker: 'Application Engineer',
    checker: 'PM & Dev Lead',
    approver: 'PM & Dev Manager',
  },
  {
    docType: 'Changelog',
    maker: 'PM & Dev Lead & Application Engineer',
    checker: 'PM & Dev Lead',
    approver: 'PM & Dev Manager',
  },
  {
    docType: 'Glossary',
    maker: 'PM & Dev Lead & Application Engineer',
    checker: 'PM & Dev Lead',
    approver: 'PM & Dev Manager',
  },
  {
    docType: 'Appendices',
    maker: 'PM & Dev Lead & Application Engineer',
    checker: 'PM & Dev Lead',
    approver: 'PM & Dev Manager',
  },
];

const DEFAULT_STORAGE_ROWS = [
  {
    category: 'Shared Drive',
    process: 'Store all documentation in the Product Management and Development Team shared drive.',
  },
  {
    category: 'Folder Path',
    process: 'Product Management and Development Team → Product List → [App Name] → Documentation',
  },
  {
    category: 'Storage Location',
    process:
      'Save all project-related documents inside the Documentation folder of the corresponding application.',
  },
  {
    category: 'Document Naming Convention',
    process: 'Name every document using the following format: (Product Name) - (Document Name)',
  },
  {
    category: 'Example File Names',
    process:
      'Inventory System - Product Requirements Document\nInventory System - API Documentation\nInventory System - User Guide\nInventory System - Release Notes',
  },
  {
    category: 'Purpose',
    process:
      'Following this structure ensures that documents are easy to locate, consistently organised, and searchable across the board.',
  },
];

const DEFAULT_TEAM_ROWS = [
  {
    role: 'PM & Dev Manager',
    name: null,
    roleType: 'Approver',
    responsibilities: [
      {
        title:
          'The individual is responsible for final authorisation or approval for product-related tasks or documents.',
        highlight: 'final authorisation or approval',
        children: [
          'Conduct a final review to ensure compliance with all relevant company policies and organisational standards.',
          'Authorised to approve or reject the task or document based on the review.',
        ],
      },
    ],
  },
  {
    role: 'PM & Dev Lead',
    name: null,
    roleType: 'Checker',
    responsibilities: [
      {
        title:
          'Responsible for reviewing and verifying the accuracy and completeness of product-related tasks or documents.',
        highlight: 'reviewing and verifying',
        children: [
          'Review the task or document for correctness, completeness, and adherence to the company policies and procedures.',
          'Verify that all required documentation and information are included.',
        ],
      },
    ],
  },
  {
    role: 'Application Engineer',
    name: null,
    roleType: 'Maker',
    responsibilities: [
      {
        title:
          'The individual is responsible for initiating or creating product-related tasks or documents.',
        highlight: 'initiating or creating',
        children: [
          'Prepare and initiate administrative tasks or documents in accordance with the company policies.',
          'Ensure all information is complete and accurate before submission.',
          'The Maker is responsible for completing the document with all the approvers.',
        ],
      },
    ],
  },
];

// =============================================================================
// SECTION GUIDE — machine-readable manifest of every section.
// =============================================================================

const SECTION_GUIDE = [
  {
    method: 'setCoverInfo',
    purpose:
      'Cover-page identity table: application name, product type, application engineer, product lead/manager, status, writer, checker, approver, last-update date, latest history note.',
    example: {
      productName: 'Acme Widget',
      productType: 'SaaS',
      applicationEngineer: 'Jane Doe',
      productLead: 'John Smith',
      status: 'Live',
      writer: 'Jane Doe',
      checker: { name: 'Atya Salma', email: 'atya.salma@point-star.com' },
      approver: { name: 'Muhammad Iqbal', email: 'iqbal@point-star.com' },
      lastUpdate: '2026-08-20',
      latestHistory: 'Updated document format Aug 20, 2026',
    },
  },
  {
    method: 'setDocumentTitle',
    purpose: 'Overrides the big cover title and italic tagline underneath it.',
    example: {
      title: 'Name of Product Documentation Flow & PIC Matrix',
      tagline: 'Creates, checks, and approves each part of the Product Documentation master file',
    },
  },
  {
    method: 'setTeam',
    purpose:
      '"1. Team" table: who holds each role, their name(s) (optionally linked to email), the PIC role (Maker/Checker/Approver), and a numbered list of what they do in this process.',
    example: [
      {
        role: 'PM & Dev Manager',
        people: [{ name: 'John Smith', email: 'john@company.com' }],
        roleType: 'Approver',
        responsibilities: [
          {
            title: 'Responsible for final authorisation of product-related documents.',
            highlight: 'final authorisation',
            children: ['Reviews for compliance.', 'Approves or rejects based on the review.'],
          },
        ],
      },
    ],
  },
  {
    method: 'setPicMatrix',
    purpose:
      '"2. PIC Matrix" table: for each document type, who is the Maker/Checker/Approver. If not called, the template\'s default 13-row matrix (Style Guide through Appendices) is used verbatim.',
    example: [
      {
        docType: 'Style Guide',
        maker: 'PM & Dev Lead',
        checker: 'PM & Dev Manager',
        approver: 'PM & Dev Manager',
      },
    ],
  },
  {
    method: 'setStorageGuide',
    purpose:
      '"3. Where and How to Store the Documentation" table: category + standard/process pairs. If not called, the template\'s default 6-row guide is used verbatim.',
    example: [
      { category: 'Shared Drive', process: 'Store all documentation in the shared drive.' },
    ],
  },
  {
    method: 'setHeaderFooterLabels',
    purpose:
      'Overrides the running header ("Product Name" / "Documentation SOP | Confidential") and footer ("PointStar / Product Team").',
    example: {
      productNameLabel: 'Acme Widget',
      docTypeLabel: 'Documentation SOP',
      companyLabel: 'Acme Inc.',
      teamLabel: 'Product Team',
    },
  },
];

// =============================================================================
// MAIN SDK CLASS
// =============================================================================

class PicMatrixSDK {
  constructor() {
    this.data = {};
  }

  /** Returns the machine-readable section manifest described above. */
  static sectionGuide() {
    return SECTION_GUIDE;
  }

  // ---------------------------------------------------------------------
  // COVER PAGE
  // ---------------------------------------------------------------------
  /**
   * Cover-page identity table.
   * @param {Object} p
   * @param {string} p.productName - Application/product name.
   * @param {string} p.productType - e.g. "SaaS", "Mobile App", "Internal Tool".
   * @param {string} p.applicationEngineer - Name of the application engineer.
   * @param {string} p.productLead - Name of the product lead/manager.
   * @param {string} p.status - e.g. "Live", "In Development", "Deprecated".
   * @param {string} p.writer - Name of whoever wrote the document.
   * @param {{name:string, email?:string}} p.checker - Checker's name (+ optional email, rendered as a mailto link).
   * @param {{name:string, email?:string}} p.approver - Approver's name (+ optional email, rendered as a mailto link).
   * @param {string} p.lastUpdate - Date this document was last updated, e.g. "2026-08-20".
   * @param {string} p.latestHistory - One-line description of the most recent change.
   */
  setCoverInfo(p = {}) {
    this.data.coverInfo = p;
    return this;
  }

  /**
   * Overrides the cover title and italic tagline.
   * @param {Object} p
   * @param {string} p.title - Big title text.
   * @param {string} p.tagline - Italic one-line description underneath.
   */
  setDocumentTitle(p = {}) {
    this.data.documentTitle = p;
    return this;
  }

  // ---------------------------------------------------------------------
  // 1. TEAM
  // ---------------------------------------------------------------------
  /**
   * "1. Team" table rows.
   * @param {Array<{
   *   role:string,
   *   people?:Array<{name:string,email?:string}>,
   *   roleType:string,
   *   responsibilities:Array<{title:string, highlight?:string, children?:string[]}>
   * }>} rows
   *   - `role`: the team role, e.g. "PM & Dev Manager".
   *   - `people`: one or more named individuals filling that role (rendered as mailto links if email given).
   *   - `roleType`: their PIC function — "Maker" | "Checker" | "Approver".
   *   - `responsibilities`: numbered list describing what they do; `highlight` bolds+colors that
   *     substring within `title`; `children` are the lettered sub-points.
   */
  setTeam(rows) {
    this.data.team = rows;
    return this;
  }

  // ---------------------------------------------------------------------
  // 2. PIC MATRIX
  // ---------------------------------------------------------------------
  /**
   * "2. PIC Matrix" table rows — who does what, per document type.
   * @param {Array<{docType:string, maker:string, checker:string, approver:string}>} rows
   */
  setPicMatrix(rows) {
    this.data.picMatrix = rows;
    return this;
  }

  // ---------------------------------------------------------------------
  // 3. STORAGE GUIDE
  // ---------------------------------------------------------------------
  /**
   * "3. Where and How to Store the Documentation" table rows.
   * @param {Array<{category:string, process:string}>} rows - `process` may contain
   *   newline-separated lines (each rendered as its own paragraph within the cell).
   */
  setStorageGuide(rows) {
    this.data.storageGuide = rows;
    return this;
  }

  // ---------------------------------------------------------------------
  // HEADER / FOOTER
  // ---------------------------------------------------------------------
  /**
   * Overrides running header/footer labels.
   * @param {Object} p
   * @param {string} [p.productNameLabel] - Left side of header, defaults to "Product Name".
   * @param {string} [p.docTypeLabel] - Middle of header, defaults to "Documentation SOP".
   * @param {string} [p.companyLabel] - Left side of footer, defaults to "PointStar".
   * @param {string} [p.teamLabel] - Right side of footer, defaults to "Product Team".
   */
  setHeaderFooterLabels(p = {}) {
    this.data.headerFooterLabels = p;
    return this;
  }

  // =========================================================================
  // BUILD
  // =========================================================================

  _buildCover() {
    const children = [];
    const title = this.data.documentTitle || {};
    const cover = this.data.coverInfo || {};

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 200 },
        children: [
          run(title.title || 'Name of Product Documentation Flow & PIC Matrix', {
            bold: true,
            color: COLOR.h1,
            size: SIZE.title,
          }),
        ],
      }),
    );
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: {
          bottom: { style: BorderStyle.SINGLE, size: 10, color: COLOR.headingRule, space: 10 },
        },
        spacing: { after: 500 },
        children: [
          run(
            title.tagline ||
              'Creates, checks, and approves each part of the Product Documentation master file',
            { italics: true, color: COLOR.h3, size: SIZE.h2 },
          ),
        ],
      }),
    );

    children.push(
      buildInfoTable([
        { label: 'Name of Application', value: cover.productName || null },
        { label: 'Product Type', value: cover.productType || null },
        { label: 'Application Engineer', value: cover.applicationEngineer || null },
        { label: 'Product Lead/Manager', value: cover.productLead || null },
        { label: 'Status', value: cover.status || null },
        { label: 'Writer', value: cover.writer || null },
        { label: 'Checker', value: cover.checker || null, isLink: true },
        { label: 'Approver', value: cover.approver || null, isLink: true },
        { label: 'Last Update', value: cover.lastUpdate || null },
        { label: 'Latest History', value: cover.latestHistory || null },
      ]),
    );

    children.push(new Paragraph({ spacing: { before: 300 }, children: [] }));
    children.push(dividerRule());
    children.push(new Paragraph({ children: [] }));

    return children;
  }

  _buildTeam() {
    const children = [];
    children.push(sectionHeading('1. Team'));

    const rows = this.data.team && this.data.team.length ? this.data.team : DEFAULT_TEAM_ROWS;
    const widths = distributeColumnWidths(
      [{ weight: 1.15 }, { weight: 1.9 }, { weight: 1.15 }, { weight: 4.65 }],
      9525,
    );

    const headerRow = new TableRow({
      tableHeader: true,
      children: [
        headerCell('Role', widths[0]),
        headerCell('Name', widths[1]),
        headerCell('Role', widths[2]),
        headerCell('What they do in this process', widths[3]),
      ],
    });

    const bodyRows = rows.map((r, rowIdx) => {
      const roleCell = new TableCell({
        width: { size: widths[0], type: WidthType.DXA },
        borders: cellBorderSet(),
        shading: { fill: COLOR.zebraFill, type: ShadingType.CLEAR, color: 'auto' },
        margins: { top: 90, bottom: 90, left: 100, right: 100 },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({ children: [run(r.role || PLACEHOLDER_ROW_TEXT, { bold: true })] }),
        ],
      });

      const people =
        r.people && r.people.length ? r.people : r.name ? [{ name: r.name, email: r.email }] : [];
      const nameParas =
        people.length > 0
          ? people.map(
              (person) =>
                new Paragraph({
                  children: person.email
                    ? [
                        new ExternalHyperlink({
                          link: `mailto:${person.email}`,
                          children: [
                            new TextRun({
                              text: person.name,
                              font: FONT,
                              size: SIZE.body,
                              color: COLOR.hyperlink,
                              underline: {},
                            }),
                          ],
                        }),
                      ]
                    : [run(person.name)],
                }),
            )
          : [
              new Paragraph({
                children: [
                  run(PLACEHOLDER_ROW_TEXT, { italics: true, color: COLOR.placeholderText }),
                ],
              }),
            ];

      const nameCell = new TableCell({
        width: { size: widths[1], type: WidthType.DXA },
        borders: cellBorderSet(),
        shading: { fill: COLOR.zebraFill, type: ShadingType.CLEAR, color: 'auto' },
        margins: { top: 90, bottom: 90, left: 100, right: 100 },
        verticalAlign: VerticalAlign.CENTER,
        children: nameParas,
      });

      const roleTypeCell = new TableCell({
        width: { size: widths[2], type: WidthType.DXA },
        borders: cellBorderSet(),
        shading: { fill: COLOR.zebraFill, type: ShadingType.CLEAR, color: 'auto' },
        margins: { top: 90, bottom: 90, left: 100, right: 100 },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [run(r.roleType || PLACEHOLDER_ROW_TEXT)],
          }),
        ],
      });

      const respCell = new TableCell({
        width: { size: widths[3], type: WidthType.DXA },
        borders: cellBorderSet(),
        shading: { fill: COLOR.zebraFill, type: ShadingType.CLEAR, color: 'auto' },
        margins: { top: 90, bottom: 90, left: 100, right: 100 },
        verticalAlign: VerticalAlign.CENTER,
        children:
          r.responsibilities && r.responsibilities.length
            ? numberedBlock(r.responsibilities, `pic-matrix-numbering-row-${rowIdx}`)
            : [
                new Paragraph({
                  children: [
                    run(PLACEHOLDER_ROW_TEXT, { italics: true, color: COLOR.placeholderText }),
                  ],
                }),
              ],
      });

      return new TableRow({ children: [roleCell, nameCell, roleTypeCell, respCell] });
    });

    children.push(
      new Table({
        width: { size: 9525, type: WidthType.DXA },
        columnWidths: widths,
        borders: outerBorderSet(),
        rows: [headerRow, ...bodyRows],
      }),
    );

    children.push(new Paragraph({ spacing: { before: 300 }, children: [] }));
    children.push(dividerRule());
    children.push(new Paragraph({ children: [] }));

    return children;
  }

  _buildPicMatrix() {
    const children = [];
    children.push(sectionHeading('2. PIC Matrix — Who Does What, Per Document'));

    children.push(
      buildZebraTable(
        [
          { key: 'docType', header: 'Type of Document', weight: 1.8, boldValues: true },
          { key: 'maker', header: 'Maker', weight: 1.3 },
          { key: 'checker', header: 'Checker', weight: 1.3 },
          { key: 'approver', header: 'Approver', weight: 1.3 },
        ],
        this.data.picMatrix,
        DEFAULT_PIC_MATRIX_ROWS,
        { width: 9360, defaultsAreRealContent: true },
      ),
    );

    children.push(pageBreak());
    return children;
  }

  _buildStorageGuide() {
    const children = [];
    children.push(sectionHeading('3. Where and How to Store the Documentation'));

    const rows =
      this.data.storageGuide && this.data.storageGuide.length
        ? this.data.storageGuide
        : DEFAULT_STORAGE_ROWS;

    // This table has bold navy labels in column 1 (own row, not header) rather
    // than a header row — matches the source template's Category/Standard
    // layout where column headers ARE the navy header row.
    const widths = distributeColumnWidths([{ weight: 1 }, { weight: 3.1 }], 9360);
    const headerRow = new TableRow({
      tableHeader: true,
      children: [headerCell('Category', widths[0]), headerCell('Standard / Process', widths[1])],
    });

    const bodyRows = rows.map((r, idx) => {
      const shadeFill = idx % 2 === 0 ? COLOR.zebraFill : COLOR.zebraFillAlt;
      const labelCell = new TableCell({
        width: { size: widths[0], type: WidthType.DXA },
        borders: cellBorderSet(),
        shading: { fill: shadeFill, type: ShadingType.CLEAR, color: 'auto' },
        margins: { top: 90, bottom: 90, left: 100, right: 100 },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            children: [run(r.category || PLACEHOLDER_ROW_TEXT, { bold: true, color: COLOR.h1 })],
          }),
        ],
      });
      const lines = (r.process || PLACEHOLDER_ROW_TEXT).split(/\r?\n+/);
      const valueCell = new TableCell({
        width: { size: widths[1], type: WidthType.DXA },
        borders: cellBorderSet(),
        shading: { fill: shadeFill, type: ShadingType.CLEAR, color: 'auto' },
        margins: { top: 90, bottom: 90, left: 100, right: 100 },
        verticalAlign: VerticalAlign.CENTER,
        children: lines.map(
          (line) =>
            new Paragraph({
              children: [
                run(line, {
                  italics: !r.process,
                  color: r.process ? COLOR.bodyText : COLOR.placeholderText,
                }),
              ],
            }),
        ),
      });
      return new TableRow({ children: [labelCell, valueCell] });
    });

    children.push(
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: widths,
        borders: outerBorderSet(),
        rows: [headerRow, ...bodyRows],
      }),
    );

    return children;
  }

  /** Assembles the full document body (array of docx elements). */
  _buildChildren() {
    return [
      ...this._buildCover(),
      ...this._buildTeam(),
      ...this._buildPicMatrix(),
      ...this._buildStorageGuide(),
    ];
  }

  _buildHeader() {
    const labels = this.data.headerFooterLabels || {};
    return new Header({
      children: [
        new Paragraph({
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR.headerRuleColor, space: 4 },
          },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            run(`${labels.productNameLabel || 'Product Name'}\t`, {
              bold: true,
              size: SIZE.headerFooter,
              color: COLOR.headerFooterText,
            }),
            run(`${labels.docTypeLabel || 'Documentation SOP'}|`, {
              bold: true,
              size: SIZE.headerFooter,
              color: COLOR.headerFooterText,
            }),
            run('Confidential', { bold: true, size: SIZE.headerFooter, color: COLOR.confidential }),
          ],
        }),
      ],
    });
  }

  _buildFooter() {
    const labels = this.data.headerFooterLabels || {};
    return new Footer({
      children: [
        new Paragraph({
          border: {
            top: { style: BorderStyle.SINGLE, size: 4, color: COLOR.footerRuleColor, space: 4 },
          },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            run(`${labels.companyLabel || 'PointStar'} \t `, {
              size: SIZE.headerFooter,
              color: COLOR.headerFooterMuted,
            }),
            run(labels.teamLabel || 'Product Team', {
              size: SIZE.headerFooter,
              color: COLOR.headerFooterMuted,
            }),
          ],
        }),
        new Paragraph({
          border: {
            top: { style: BorderStyle.SINGLE, size: 4, color: COLOR.footerRuleColor, space: 4 },
          },
          alignment: AlignmentType.CENTER,
          children: [
            run('Page 1 of 1', { size: SIZE.headerFooter, color: COLOR.headerFooterMuted }),
          ],
        }),
      ],
    });
  }

  /**
   * Builds one abstract-numbering config entry per team row so each row's
   * "1. / a. / i." list restarts fresh instead of continuing to count
   * across the whole "1. Team" table. `count` is sized generously above
   * however many team rows are actually supplied.
   */
  _buildNumberingConfig(count) {
    const levels = () => [
      {
        level: 0,
        format: 'decimal',
        text: '%1.',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      },
      {
        level: 1,
        format: 'lowerLetter',
        text: '%2.',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1440, hanging: 360 } } },
      },
      {
        level: 2,
        format: 'lowerRoman',
        text: '%3.',
        alignment: AlignmentType.RIGHT,
        style: { paragraph: { indent: { left: 2160, hanging: 360 } } },
      },
    ];
    const configs = [{ reference: 'pic-matrix-numbering', levels: levels() }];
    for (let i = 0; i < count; i++) {
      configs.push({ reference: `pic-matrix-numbering-row-${i}`, levels: levels() });
    }
    return configs;
  }

  /** Builds the underlying `docx` Document object. */
  /**
   * The raw `docx` Document options for this Part — styles, numbering, and
   * its two sections (title page + content). `generate-master.js` reads this
   * so every Part can be concatenated into one Document instead of merged
   * as separate .docx files.
   */
  documentOptions() {
    const teamRowCount = (
      this.data.team && this.data.team.length ? this.data.team : DEFAULT_TEAM_ROWS
    ).length;
    return {
      numbering: {
        config: this._buildNumberingConfig(teamRowCount),
      },
      styles: {
        default: { document: { run: { font: FONT, size: SIZE.body } } },
        paragraphStyles: [
          {
            id: 'Title',
            name: 'Title',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            run: { font: FONT, size: SIZE.title, bold: true, color: COLOR.h1 },
            paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 200 } },
          },
          {
            id: 'Heading1',
            name: 'Heading 1',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            run: { font: FONT, size: SIZE.h1, bold: true, color: COLOR.h1 },
            paragraph: { spacing: { before: 400, after: 200 } },
          },
          {
            id: 'Heading2',
            name: 'Heading 2',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            run: { font: FONT, size: SIZE.h2, bold: true, color: COLOR.h2 },
            paragraph: { spacing: { before: 320, after: 160 } },
          },
          {
            id: 'Heading3',
            name: 'Heading 3',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            run: { font: FONT, size: SIZE.h3, bold: true, color: COLOR.h3 },
            paragraph: { spacing: { before: 240, after: 120 } },
          },
        ],
      },
      sections: [
        {
          properties: {
            page: {
              size: { width: PAGE.width, height: PAGE.height },
              margin: PAGE.margin,
            },
          },
          headers: { default: this._buildHeader() },
          footers: { default: this._buildFooter() },
          children: this._buildChildren(),
        },
      ],
    };
  }

  toDocument() {
    return new Document(this.documentOptions());
  }

  /** Returns a Buffer containing the .docx file (does not write to disk). */
  async toBuffer() {
    return Packer.toBuffer(this.toDocument());
  }

  /**
   * Renders and writes the .docx file to disk.
   * @param {string} outputPath
   * @returns {Promise<string>} the output path
   */
  async generate(outputPath) {
    const buf = await this.toBuffer();
    fs.writeFileSync(outputPath, buf);
    return outputPath;
  }
}

module.exports = PicMatrixSDK;
module.exports.SECTION_GUIDE = SECTION_GUIDE;
module.exports.DEFAULT_PIC_MATRIX_ROWS = DEFAULT_PIC_MATRIX_ROWS;
module.exports.DEFAULT_STORAGE_ROWS = DEFAULT_STORAGE_ROWS;
module.exports.DEFAULT_TEAM_ROWS = DEFAULT_TEAM_ROWS;
module.exports.numberedBlock = numberedBlock;
module.exports.buildZebraTable = buildZebraTable;
