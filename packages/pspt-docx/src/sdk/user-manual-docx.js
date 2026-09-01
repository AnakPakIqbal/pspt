'use strict';
/**
 * =============================================================================
 *  USER MANUAL (Part 10 of the Product Documentation Master)
 * =============================================================================
 * Regenerates "Part 10. User Manual" exactly: title page, metadata table,
 * 7 numbered sections (Introduction, Getting Started, Navigating the
 * Interface, Feature Walkthrough with repeatable "How to [Task]" H3 blocks,
 * FAQ, Troubleshooting table, Support Contact). Italic bracketed screenshot
 * placeholders ("[Insert screenshot of the login screen]") appear inline
 * after several sections.
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
  bodyPara,
  buildZebraTable,
  bulletNumberingConfig,
  guidanceNote,
  h1,
  h2,
  h3,
  headingRun,
  makeBullet,
  outerBorderSet,
  run,
} = require('./pspt-core');

const BULLET_NUMBERING_REF = 'usermanual-bullet-list';

/** Single-level bullet paragraph bound to this Part's numbering reference. */
const bullet = makeBullet(BULLET_NUMBERING_REF);

function screenshotNote(text) {
  return new Paragraph({
    spacing: { after: 160, line: 276, lineRule: 'auto' },
    children: [run(text, { italics: true })],
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
const DEFAULT_INTRODUCTION =
  '[What the application is and who this manual is for, in plain language.]';
const DEFAULT_GETTING_STARTED =
  '[How to access the system and log in for the first time, including prerequisites.]';
const DEFAULT_GETTING_STARTED_SCREENSHOT = '[Insert screenshot of the login screen]';
const DEFAULT_NAVIGATING =
  '[Orient the user to the main layout — sidebar, header, common buttons — before task walkthroughs.]';
const DEFAULT_NAVIGATING_SCREENSHOT = '[Insert annotated screenshot of the main layout]';
const DEFAULT_FEATURE_WALKTHROUGHS = [
  {
    taskTitle: 'How to [Do a Core Task]',
    steps: [
      'In the sidebar, click [Module].',
      'Click the [+ New] button.',
      'Fill in the required fields.',
      'Click [Save].',
    ],
    screenshot: '[Insert screenshot of the form]',
  },
];
const DEFAULT_FAQ = 'Q: [Common question]? A: [Answer].';
const DEFAULT_TROUBLESHOOTING = [
  {
    problem: '[e.g. "Invalid credentials" on login]',
    solution: '[e.g. Confirm Caps Lock is off; use Forgot Password]',
  },
  { problem: "[e.g. Download won't start]", solution: '[e.g. Check pop-up blocker settings]' },
];
const DEFAULT_SUPPORT_CONTACT =
  'For further help, contact [support channel] or via [internal help-desk portal].';

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
    method: 'setIntroduction',
    purpose: 'Section 1 introduction paragraph.',
    example: 'Acme Widget is an expense tracker for small teams.',
  },
  {
    method: 'setGettingStarted',
    purpose: 'Section 2 access/login instructions.',
    example: 'Visit app.acme.com and sign in with your company email.',
  },
  {
    method: 'setGettingStartedScreenshot',
    purpose: 'Section 2 screenshot caption.',
    example: '[Insert screenshot of the login screen]',
  },
  {
    method: 'setNavigating',
    purpose: 'Section 3 layout orientation paragraph.',
    example: 'The sidebar on the left lists Expenses, Reports, and Settings.',
  },
  {
    method: 'setNavigatingScreenshot',
    purpose: 'Section 3 screenshot caption.',
    example: '[Insert annotated screenshot of the main layout]',
  },
  {
    method: 'setFeatureWalkthroughs',
    purpose:
      'Section 4 repeatable "How to [Task]" blocks with numbered steps and an optional screenshot line.',
    example: [
      {
        taskTitle: 'How to submit an expense',
        steps: ['Click + New Expense.', 'Fill in amount and category.', 'Click Save.'],
        screenshot: '[Insert screenshot of the expense form]',
      },
    ],
  },
  {
    method: 'setFaq',
    purpose: 'Section 5 FAQ paragraph (Q:/A: format).',
    example: 'Q: Why was my expense rejected? A: Check the rejection reason in the activity log.',
  },
  {
    method: 'setTroubleshooting',
    purpose: 'Section 6 Problem/Solution table.',
    example: [{ problem: 'Upload fails', solution: 'Check file is under 10MB' }],
  },
  {
    method: 'setSupportContact',
    purpose: 'Section 7 support contact paragraph.',
    example: 'Contact support@acme.com or use the in-app chat widget.',
  },
];

class UserManualSDK {
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
  setMetadata(p = {}) {
    this.data.metadata = p;
    return this;
  }
  setIntroduction(t) {
    this.data.introduction = t;
    return this;
  }
  setGettingStarted(t) {
    this.data.gettingStarted = t;
    return this;
  }
  setGettingStartedScreenshot(t) {
    this.data.gettingStartedScreenshot = t;
    return this;
  }
  setNavigating(t) {
    this.data.navigating = t;
    return this;
  }
  setNavigatingScreenshot(t) {
    this.data.navigatingScreenshot = t;
    return this;
  }
  setFeatureWalkthroughs(items) {
    this.data.featureWalkthroughs = items;
    return this;
  }
  setFaq(t) {
    this.data.faq = t;
    return this;
  }
  setTroubleshooting(rows) {
    this.data.troubleshooting = rows;
    return this;
  }
  setSupportContact(t) {
    this.data.supportContact = t;
    return this;
  }

  _buildPartTitlePage() {
    return [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [headingRun('User Manual')] }),
    ];
  }

  _buildContent() {
    const children = [];
    children.push(h1('User Manual'));
    const meta = Object.assign({}, DEFAULT_METADATA, this.data.metadata || {});
    children.push(this._metadataTable(meta));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));

    children.push(h2('1. Introduction'));
    children.push(bodyPara(this.data.introduction || DEFAULT_INTRODUCTION));

    children.push(h2('2. Getting Started'));
    children.push(bodyPara(this.data.gettingStarted || DEFAULT_GETTING_STARTED));
    children.push(
      screenshotNote(this.data.gettingStartedScreenshot || DEFAULT_GETTING_STARTED_SCREENSHOT),
    );

    children.push(h2('3. Navigating the Interface'));
    children.push(bodyPara(this.data.navigating || DEFAULT_NAVIGATING));
    children.push(screenshotNote(this.data.navigatingScreenshot || DEFAULT_NAVIGATING_SCREENSHOT));

    children.push(h2('4. Feature Walkthrough'));
    children.push(
      guidanceNote(
        'Repeat this block per feature/module — step-by-step, task-oriented instructions a first-time user could follow without assistance.',
      ),
    );
    const walkthroughs =
      this.data.featureWalkthroughs && this.data.featureWalkthroughs.length
        ? this.data.featureWalkthroughs
        : DEFAULT_FEATURE_WALKTHROUGHS;
    walkthroughs.forEach((w) => {
      children.push(h3(w.taskTitle));
      (w.steps || []).forEach((s) => children.push(bullet(s)));
      if (w.screenshot) {
        children.push(new Paragraph({ spacing: { before: 100 }, children: [] }));
        children.push(screenshotNote(w.screenshot));
      }
    });

    children.push(h2('5. Frequently Asked Questions (FAQ)'));
    children.push(bodyPara(this.data.faq || DEFAULT_FAQ));

    children.push(h2('6. Troubleshooting'));
    children.push(
      buildZebraTable(
        [
          { key: 'problem', header: 'Problem', weight: 1.5 },
          { key: 'solution', header: 'Solution', weight: 1.5 },
        ],
        this.data.troubleshooting,
        DEFAULT_TROUBLESHOOTING,
      ),
    );

    children.push(h2('7. Support Contact'));
    children.push(bodyPara(this.data.supportContact || DEFAULT_SUPPORT_CONTACT));

    return children;
  }

  _metadataTable(meta) {
    const labelWidth = 2400;
    const valueWidth = CONTENT_WIDTH - labelWidth;
    const rows = [
      { label: 'Writer', value: meta.writer },
      { label: 'Status', value: meta.status },
      { label: 'Version', value: meta.version },
      { label: 'Last Update', value: meta.lastUpdate },
    ];
    const trRows = rows.map(
      (r) =>
        new TableRow({
          children: [
            new TableCell({
              width: { size: labelWidth, type: WidthType.DXA },
              shading: { fill: COLOR.tableHeaderBg, type: ShadingType.CLEAR, color: 'auto' },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              verticalAlign: VerticalAlign.CENTER,
              children: [
                new Paragraph({
                  children: [run(r.label, { bold: true, color: COLOR.tableHeaderText })],
                }),
              ],
            }),
            new TableCell({
              width: { size: valueWidth, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              verticalAlign: VerticalAlign.CENTER,
              children: [
                new Paragraph({ children: [run(r.value == null ? '' : String(r.value))] }),
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

  _blankHeader() {
    return new Header({ children: [new Paragraph({ children: [] })] });
  }
  _blankFooter() {
    return new Footer({ children: [new Paragraph({ children: [] })] });
  }
  _chapterHeader() {
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
            run('User Manual', {
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
          {
            id: 'Heading3',
            name: 'Heading 3',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            run: { font: FONT, size: SIZE.h2, bold: true, color: COLOR.h3 },
            paragraph: { spacing: { before: 240, after: 120 } },
          },
        ],
      },
      sections: [
        {
          properties: {
            titlePage: true,
            page: { size: { width: PAGE.width, height: PAGE.height }, margin: PAGE.margin },
          },
          headers: { default: this._blankHeader(), first: this._blankHeader() },
          footers: { default: this._blankFooter(), first: this._blankFooter() },
          children: this._buildPartTitlePage(),
        },
        {
          properties: {
            page: { size: { width: PAGE.width, height: PAGE.height }, margin: PAGE.margin },
          },
          headers: { default: this._chapterHeader() },
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

module.exports = UserManualSDK;
module.exports.SECTION_GUIDE = SECTION_GUIDE;
module.exports.buildZebraTable = buildZebraTable;
module.exports.BULLET_NUMBERING_REF = BULLET_NUMBERING_REF;
