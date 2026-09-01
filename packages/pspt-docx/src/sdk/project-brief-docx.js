'use strict';
/**
 * =============================================================================
 *  PROJECT BRIEF (Part 2 of the Product Documentation Master)
 * =============================================================================
 * Regenerates "Part 2. Project Brief" exactly: title-only page, then content
 * with a vertical navy-label metadata table (Writer/Status/Version/Last
 * Update), guidance italic notes, bullet lists, and simple 2-column tables.
 *
 * HOW TO USE
 * -----------
 *   const ProjectBriefSDK = require('./project-brief-docx');
 *   const doc = new ProjectBriefSDK();
 *   doc.setMetadata({ writer: 'Jane Doe', status: 'Draft', version: 'V1 (Phase 1)', lastUpdate: 'Aug 31, 2026' });
 *   doc.setOverview('Acme Widget lets teams track expenses in real time...');
 *   doc.setBackgroundPains(['Manual expense entry takes 20 min/week', 'No approval audit trail']);
 *   doc.setObjectives(['Cut entry time to under 2 minutes', 'Achieve 90% adoption by Q4']);
 *   doc.setKeyModules([{ module: 'Receipts', features: 'OCR capture + auto-categorisation' }]);
 *   doc.setTimeline([{ phase: 'Phase 1 — Discovery & Design', duration: '2 weeks' }]);
 *   doc.setDeliverables(['Mobile app (iOS/Android)', 'Technical documentation (this document)']);
 *   doc.setPreliminaryRisks([{ risk: 'Bank API rate limits', mitigation: 'Cache + backoff strategy' }]);
 *   await doc.generate('/mnt/user-data/outputs/02-project-brief.docx');
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
  PageNumber,
} = require('docx');
const fs = require('fs');

const {
  DOCX_COLOR: COLOR,
  DOCX_CONTENT_WIDTH: CONTENT_WIDTH,
  DOCX_FONT: FONT,
  DOCX_PAGE: PAGE,
  DOCX_SIZE: SIZE,
  PLACEHOLDER_ROW_TEXT,
  bodyParaJustified: bodyPara,
  bulletNumberingConfig,
  distributeColumnWidths,
  guidanceNote,
  h1,
  h2,
  headingRun,
  makeBullet,
  outerBorderSet,
  run,
} = require('./pspt-core');

// -----------------------------------------------------------------------
// Bullet-list numbering reference shared by every "Part" content chapter
// that uses simple (non-restarting) bullet lists: ● / ○ / ■.
// -----------------------------------------------------------------------
const BULLET_NUMBERING_REF = 'part-bullet-list';

/** Single-level bullet paragraph bound to this Part's numbering reference. */
const bullet = makeBullet(BULLET_NUMBERING_REF);

function plainCellMargins() {
  return { top: 80, bottom: 80, left: 120, right: 120 };
}

/**
 * Vertical navy-label metadata table: label column (navy fill, bold white
 * text) on the left, value column (plain white, no shading) on the right.
 * @param {Array<{label:string, value:string}>} rows
 */
function buildMetadataTable(rows) {
  const labelWidth = 2400;
  const valueWidth = CONTENT_WIDTH - labelWidth;
  const trRows = rows.map(
    (r) =>
      new TableRow({
        children: [
          new TableCell({
            width: { size: labelWidth, type: WidthType.DXA },
            shading: { fill: COLOR.tableHeaderBg, type: ShadingType.CLEAR, color: 'auto' },
            margins: plainCellMargins(),
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                children: [run(r.label, { bold: true, color: COLOR.tableHeaderText })],
              }),
            ],
          }),
          new TableCell({
            width: { size: valueWidth, type: WidthType.DXA },
            margins: plainCellMargins(),
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                children: [run(r.value == null || r.value === '' ? '' : String(r.value), {})],
              }),
            ],
          }),
        ],
      }),
  );
  return new Table({
    width: { size: labelWidth + valueWidth, type: WidthType.DXA },
    columnWidths: [labelWidth, valueWidth],
    borders: outerBorderSet(),
    rows: trRows,
  });
}

/** Simple 2-column navy-header table (no zebra shading — matches the
 * Project Brief's plain "Module/Core Features" and "Risk/Mitigation" style). */
function buildSimpleTable(columns, rows, placeholderRow) {
  const widths = distributeColumnWidths(columns, CONTENT_WIDTH);
  const useRows = rows && rows.length ? rows : [placeholderRow];
  const isPlaceholder = !rows || !rows.length;

  const headerRow = new TableRow({
    tableHeader: true,
    children: columns.map(
      (col, idx) =>
        new TableCell({
          width: { size: widths[idx], type: WidthType.DXA },
          shading: { fill: COLOR.tableHeaderBg, type: ShadingType.CLEAR, color: 'auto' },
          margins: plainCellMargins(),
          children: [
            new Paragraph({
              children: [run(col.header, { bold: true, color: COLOR.tableHeaderText })],
            }),
          ],
        }),
    ),
  });

  const bodyRows = useRows.map(
    (r) =>
      new TableRow({
        children: columns.map((col, idx) => {
          const value = r ? r[col.key] : undefined;
          const text =
            value == null || value === ''
              ? isPlaceholder
                ? PLACEHOLDER_ROW_TEXT
                : ''
              : String(value);
          return new TableCell({
            width: { size: widths[idx], type: WidthType.DXA },
            margins: plainCellMargins(),
            children: [new Paragraph({ children: [run(text)] })],
          });
        }),
      }),
  );

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: widths,
    borders: outerBorderSet(),
    rows: [headerRow, ...bodyRows],
  });
}

// =============================================================================
// DEFAULTS — verbatim from the master template.
// =============================================================================

const DEFAULT_METADATA = {
  writer: '',
  status: 'Draft',
  version: 'V1 (Phase )',
  lastUpdate: 'Aug 20, 2026',
};
const DEFAULT_OVERVIEW =
  '[Summarise what is being built, who it is for, and what core problem or opportunity it addresses.]';
const DEFAULT_BACKGROUND_PAINS = ['[Pain point 1]', '[Pain point 2]', '[…]'];
const DEFAULT_OBJECTIVES = ['[Objective 1]', '[Objective 2]', '[…]'];
const DEFAULT_KEY_MODULES = [{ module: '[Module]', features: '[One-line feature summary]' }];
const DEFAULT_TIMELINE = [
  { phase: 'Phase 1 — Discovery & Design', duration: '[duration]' },
  { phase: 'Phase 2 — MVP', duration: '[duration]' },
  { phase: 'Phase 3 — Remaining scope', duration: '[duration]' },
  { phase: 'Phase 4 — Testing & UAT', duration: '[duration]' },
  { phase: 'Phase 5 — Deployment & Training', duration: '[duration]' },
];
const DEFAULT_DELIVERABLES = [
  '[Software deliverable]',
  'Technical documentation (this document)',
  'User manual (Part 13)',
  'Training/onboarding session',
];
const DEFAULT_RISKS = [{ risk: '[Risk]', mitigation: '[Mitigation]' }];

const SECTION_GUIDE = [
  {
    method: 'setHeaderFooterLabels',
    purpose: 'Overrides the running header product name label.',
    example: { productNameLabel: 'Acme Widget' },
  },
  {
    method: 'setMetadata',
    purpose: 'Writer/Status/Version/Last Update metadata table.',
    example: { writer: 'Jane Doe', status: 'Draft', version: 'V1', lastUpdate: 'Aug 31, 2026' },
  },
  {
    method: 'setOverview',
    purpose: '2-4 sentence elevator pitch (Section 1).',
    example: 'Acme Widget is a...',
  },
  {
    method: 'setBackgroundPains',
    purpose: 'Bulleted pain points (Section 2).',
    example: ['Manual entry is slow', 'No audit trail'],
  },
  {
    method: 'setObjectives',
    purpose: 'Bulleted measurable objectives (Section 3).',
    example: ['Cut entry time to 2 minutes'],
  },
  {
    method: 'setKeyModules',
    purpose: 'Module/Core Features table (Section 4).',
    example: [{ module: 'Receipts', features: 'OCR capture' }],
  },
  {
    method: 'setTimeline',
    purpose: 'Phase/duration bullet list (Section 5).',
    example: [{ phase: 'Phase 1 — Discovery', duration: '2 weeks' }],
  },
  {
    method: 'setDeliverables',
    purpose: 'Bulleted deliverables list (Section 6).',
    example: ['Mobile app', 'User manual'],
  },
  {
    method: 'setPreliminaryRisks',
    purpose: 'Risk/Mitigation table (Section 7).',
    example: [{ risk: 'API rate limits', mitigation: 'Backoff + caching' }],
  },
];

class ProjectBriefSDK {
  constructor() {
    this.data = {};
  }

  static sectionGuide() {
    return SECTION_GUIDE;
  }

  setHeaderFooterLabels(p = {}) {
    this.data.headerFooterLabels = p;
    return this;
  }

  /** @param {Object} p @param {string} p.writer @param {string} p.status @param {string} p.version @param {string} p.lastUpdate */
  setMetadata(p = {}) {
    this.data.metadata = p;
    return this;
  }

  /** @param {string} text - 2-4 sentence elevator pitch. */
  setOverview(text) {
    this.data.overview = text;
    return this;
  }

  /** @param {string[]} items */
  setBackgroundPains(items) {
    this.data.backgroundPains = items;
    return this;
  }

  /** @param {string[]} items */
  setObjectives(items) {
    this.data.objectives = items;
    return this;
  }

  /** @param {Array<{module:string, features:string}>} rows */
  setKeyModules(rows) {
    this.data.keyModules = rows;
    return this;
  }

  /** @param {Array<{phase:string, duration:string}>} rows */
  setTimeline(rows) {
    this.data.timeline = rows;
    return this;
  }

  /** @param {string[]} items */
  setDeliverables(items) {
    this.data.deliverables = items;
    return this;
  }

  /** @param {Array<{risk:string, mitigation:string}>} rows */
  setPreliminaryRisks(rows) {
    this.data.preliminaryRisks = rows;
    return this;
  }

  // =========================================================================
  // BUILD
  // =========================================================================

  _buildPartTitlePage() {
    return [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [headingRun('Project Brief')] }),
    ];
  }

  _buildContent() {
    const children = [];
    children.push(h1('Project Brief'));

    const meta = Object.assign({}, DEFAULT_METADATA, this.data.metadata || {});
    children.push(
      buildMetadataTable([
        { label: 'Writer', value: meta.writer },
        { label: 'Status', value: meta.status },
        { label: 'Version', value: meta.version },
        { label: 'Last Update', value: meta.lastUpdate },
      ]),
    );
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));

    children.push(h2('1. Project Overview'));
    children.push(
      guidanceNote(
        "2–4 sentences. A reader who only reads this paragraph should understand the project's elevator pitch without needing to read any other section.",
      ),
    );
    children.push(bodyPara(this.data.overview || DEFAULT_OVERVIEW));

    children.push(h2('2. Background & Problem Statement'));
    children.push(
      guidanceNote(
        'Describe the current (as-is) situation and the specific pains that justify this project. Focus on facts and symptoms — the solution comes later.',
      ),
    );
    const pains =
      this.data.backgroundPains && this.data.backgroundPains.length
        ? this.data.backgroundPains
        : DEFAULT_BACKGROUND_PAINS;
    pains.forEach((p) => children.push(bullet(p)));

    children.push(h2('3. Project Objectives'));
    children.push(
      guidanceNote(
        'Measurable outcomes the project should achieve. Each objective should map back to at least one of the problems above.',
      ),
    );
    const objectives =
      this.data.objectives && this.data.objectives.length
        ? this.data.objectives
        : DEFAULT_OBJECTIVES;
    objectives.forEach((o) => children.push(bullet(o)));

    children.push(h2('4. Key Modules & Features'));
    children.push(
      guidanceNote(
        'Keep module names consistent with the BRD (Part 3) and PRD (Part 4) — this table seeds both.',
      ),
    );
    children.push(
      buildSimpleTable(
        [
          { key: 'module', header: 'Module', weight: 1 },
          { key: 'features', header: 'Core Features', weight: 2.5 },
        ],
        this.data.keyModules,
        DEFAULT_KEY_MODULES[0],
      ),
    );

    children.push(h2('5. Timeline Estimation'));
    children.push(guidanceNote('Set estimation — not a detailed schedule.'));
    const timeline =
      this.data.timeline && this.data.timeline.length ? this.data.timeline : DEFAULT_TIMELINE;
    timeline.forEach((t) => children.push(bullet(`${t.phase}: ${t.duration}`)));

    children.push(h2('6. Deliverables'));
    const deliverables =
      this.data.deliverables && this.data.deliverables.length
        ? this.data.deliverables
        : DEFAULT_DELIVERABLES;
    deliverables.forEach((d) => children.push(bullet(d)));

    children.push(h2('7. Preliminary Risks'));
    children.push(
      guidanceNote(
        "Early risks are visible before requirements exist. Once the PRD's Risk Register (Part 4, Section 8) is opened, log new risks there — this table isn't repeated or re-maintained in parallel.",
      ),
    );
    children.push(
      buildSimpleTable(
        [
          { key: 'risk', header: 'Risk', weight: 1.2 },
          { key: 'mitigation', header: 'Mitigation', weight: 1.2 },
        ],
        this.data.preliminaryRisks,
        DEFAULT_RISKS[0],
      ),
    );

    return children;
  }

  _blankHeader() {
    return new Header({ children: [new Paragraph({ children: [] })] });
  }

  _blankFooter() {
    return new Footer({ children: [new Paragraph({ children: [] })] });
  }

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
            run(`${labels.productNameLabel || 'Product name/logo'}\t`, {
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

  /**
   * The raw `docx` Document options for this Part — styles, numbering, and
   * its two sections (title page + content). `generate-master.js` reads this
   * so every Part can be concatenated into one Document instead of merged
   * as separate .docx files.
   */
  documentOptions() {
    return {
      numbering: {
        config: [bulletNumberingConfig(BULLET_NUMBERING_REF)],
      },
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
        ],
      },
      sections: [
        {
          properties: {
            titlePage: true,
            page: {
              size: { width: PAGE.width, height: PAGE.height },
              margin: PAGE.margin,
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
              margin: PAGE.margin,
            },
          },
          headers: { default: this._chapterHeader('Project Brief') },
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

module.exports = ProjectBriefSDK;
module.exports.SECTION_GUIDE = SECTION_GUIDE;
module.exports.buildMetadataTable = buildMetadataTable;
module.exports.buildSimpleTable = buildSimpleTable;
module.exports.BULLET_NUMBERING_REF = BULLET_NUMBERING_REF;
