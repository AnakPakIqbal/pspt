'use strict';
/**
 * =============================================================================
 *  USER ACCEPTANCE TESTING / UAT (Part 8 of the Product Documentation Master)
 * =============================================================================
 * Regenerates "Part 8. User Acceptance Testing (UAT)" exactly: title page
 * (reads "User Acceptance Testing(UAT)", no space before the parenthesis —
 * reproduced verbatim from the source), metadata table, 9 numbered sections
 * (Objectives & Scope, Test Strategy, Test Environment, Test Schedule,
 * Roles & Responsibilities, Test Cases, Non-Functional Testing, Bug Severity
 * Definitions, Sign-off). Running header label is the short form "UAT".
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
  h1,
  h2,
  h3,
  headingRun,
  makeBullet,
  outerBorderSet,
  run,
} = require('./pspt-core');

const BULLET_NUMBERING_REF = 'uat-bullet-list';

/** Single-level bullet paragraph bound to this Part's numbering reference. */
const bullet = makeBullet(BULLET_NUMBERING_REF);

// =============================================================================
// DEFAULTS — verbatim from the master template.
// =============================================================================

const DEFAULT_METADATA = {
  writer: '',
  status: 'Draft',
  version: 'V1 (Phase )',
  lastUpdate: 'Aug 20, 2026',
};
const DEFAULT_OBJECTIVES_SCOPE =
  "[What this test plan verifies and why — a quality goal, not just 'find bugs'.]";
const DEFAULT_FEATURES_TO_TEST = ['[Feature/flow being tested]'];
const DEFAULT_TEST_STRATEGY = [
  { level: 'Unit', description: 'Individual functions/components', owner: 'Developers' },
  {
    level: 'Integration',
    description: 'API endpoints against a test database',
    owner: 'Developers',
  },
  { level: 'System', description: 'Full workflows end-to-end', owner: 'QA' },
  {
    level: 'UAT',
    description: 'Business stakeholders verify real scenarios',
    owner: 'Business reps',
  },
];
const DEFAULT_TEST_ENVIRONMENTS = [
  {
    environment: 'Staging',
    url: '[staging.example.com]',
    database: 'Seeded test data, reset weekly',
    notes: 'Used for System & UAT testing',
  },
];
const DEFAULT_TEST_ACCOUNTS = [
  'Test accounts: [credentials location — never paste passwords here]',
  'Devices/browsers: [e.g. Chrome, Safari, iOS 17+, Android 13+]',
];
const DEFAULT_TEST_SCHEDULE = [
  { activity: 'System Testing', start: '[date]', end: '[date]' },
  { activity: 'UAT', start: '[date]', end: '[date]' },
];
const DEFAULT_ROLES_RESPONSIBILITIES = [
  { role: 'QA Engineer', responsibility: 'Write and execute test cases, log defects' },
  { role: 'Dev Lead', responsibility: 'Triage and fix defects' },
  { role: 'Business/UAT Rep', responsibility: 'UAT sign-off' },
];
const DEFAULT_TEST_CASES = [
  {
    id: 'TC-01',
    relatedUseCase: 'UC-01',
    steps: '1. [Step] 2. [Step]',
    expectedResult: '[Expected result]',
    status: 'Not Run',
  },
  {
    id: 'TC-02',
    relatedUseCase: 'UC-03',
    steps: '1. [Step] 2. [Step]',
    expectedResult: '[Expected result]',
    status: 'Not Run',
  },
];
const DEFAULT_NON_FUNCTIONAL_TESTING = [
  { type: 'Performance', criteria: 'Page load < 2s at p95', required: 'Yes', status: 'Not Run' },
  { type: 'Accessibility', criteria: 'WCAG 2.1 AA', required: 'Yes', status: 'Not Run' },
  { type: 'Security', criteria: 'See Part 6 §4.9 checklist', required: 'Yes', status: 'Not Run' },
];
const DEFAULT_BUG_SEVERITY = [
  { severity: 'S0 — Blocker', definition: 'Crashes, data loss, or block release entirely' },
  { severity: 'S1 — Critical', definition: 'Core flow broken, no workaround' },
  { severity: 'S2 — Major', definition: 'Significant issue, workaround exists' },
  { severity: 'S3 — Minor', definition: 'Cosmetic or low-impact issue' },
];
const DEFAULT_SIGNOFF = [
  { role: 'Product Lead/Manager', name: '[name]', approved: '[Yes/No]', date: '[date]' },
  { role: 'Engineer', name: '[name]', approved: '[Yes/No]', date: '[date]' },
  { role: 'User/UAT Rep', name: '[name]', approved: '[Yes/No]', date: '[date]' },
];

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
    method: 'setObjectivesScope',
    purpose: 'Section 1 objectives/scope paragraph.',
    example: 'Verifies the expense workflow meets business needs before launch.',
  },
  {
    method: 'setFeaturesToTest',
    purpose: 'Section 1.1 bulleted features.',
    example: ['Expense submission', 'Approval workflow'],
  },
  {
    method: 'setTestStrategy',
    purpose: 'Section 2 Level/Description/Owner table.',
    example: [{ level: 'Unit', description: 'Component tests', owner: 'Developers' }],
  },
  {
    method: 'setTestEnvironments',
    purpose: 'Section 3 Environment/URL/Database/Notes table.',
    example: [
      {
        environment: 'Staging',
        url: 'staging.acme.com',
        database: 'Seeded weekly',
        notes: 'Primary test env',
      },
    ],
  },
  {
    method: 'setTestAccounts',
    purpose: 'Section 3.1 bulleted accounts/devices notes.',
    example: ['Test accounts in 1Password vault', 'Chrome, Safari, iOS 17+'],
  },
  {
    method: 'setTestSchedule',
    purpose: 'Section 4 Activity/Start/End table.',
    example: [{ activity: 'System Testing', start: 'Sept 1', end: 'Sept 10' }],
  },
  {
    method: 'setRolesResponsibilities',
    purpose: 'Section 5 Role/Responsibility table.',
    example: [{ role: 'QA Engineer', responsibility: 'Execute test cases' }],
  },
  {
    method: 'setTestCases',
    purpose: 'Section 6 ID/Related Use Case/Steps/Expected Result/Status table.',
    example: [
      {
        id: 'TC-01',
        relatedUseCase: 'UC-01',
        steps: '1. Login 2. Submit expense',
        expectedResult: 'Expense saved',
        status: 'Not Run',
      },
    ],
  },
  {
    method: 'setNonFunctionalTesting',
    purpose: 'Section 7 Type/Criteria/Required/Status table.',
    example: [
      { type: 'Performance', criteria: 'Page load < 2s', required: 'Yes', status: 'Not Run' },
    ],
  },
  {
    method: 'setBugSeverityDefinitions',
    purpose: 'Section 8 Severity/Definition table.',
    example: [{ severity: 'S0 — Blocker', definition: 'Crashes or data loss' }],
  },
  {
    method: 'setSignoff',
    purpose: 'Section 9 Role/Name/Approved/Date table.',
    example: [{ role: 'Product Lead', name: 'Jane Doe', approved: 'Yes', date: 'Aug 31, 2026' }],
  },
];

class UatSDK {
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
  setObjectivesScope(t) {
    this.data.objectivesScope = t;
    return this;
  }
  setFeaturesToTest(items) {
    this.data.featuresToTest = items;
    return this;
  }
  setTestStrategy(rows) {
    this.data.testStrategy = rows;
    return this;
  }
  setTestEnvironments(rows) {
    this.data.testEnvironments = rows;
    return this;
  }
  setTestAccounts(items) {
    this.data.testAccounts = items;
    return this;
  }
  setTestSchedule(rows) {
    this.data.testSchedule = rows;
    return this;
  }
  setRolesResponsibilities(rows) {
    this.data.rolesResponsibilities = rows;
    return this;
  }
  setTestCases(rows) {
    this.data.testCases = rows;
    return this;
  }
  setNonFunctionalTesting(rows) {
    this.data.nonFunctionalTesting = rows;
    return this;
  }
  setBugSeverityDefinitions(rows) {
    this.data.bugSeverityDefinitions = rows;
    return this;
  }
  setSignoff(rows) {
    this.data.signoff = rows;
    return this;
  }

  _buildPartTitlePage() {
    // Reproduced verbatim: no space before the parenthesis on the title page.
    return [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [headingRun('User Acceptance Testing(UAT)')],
      }),
    ];
  }

  _buildContent() {
    const children = [];
    children.push(h1('User Acceptance Testing'));
    const meta = Object.assign({}, DEFAULT_METADATA, this.data.metadata || {});
    children.push(this._metadataTable(meta));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));

    children.push(h2('1. Objectives & Scope'));
    children.push(bodyPara(this.data.objectivesScope || DEFAULT_OBJECTIVES_SCOPE));
    children.push(h3('1.1 Features to Test'));
    (this.data.featuresToTest && this.data.featuresToTest.length
      ? this.data.featuresToTest
      : DEFAULT_FEATURES_TO_TEST
    ).forEach((f) => children.push(bullet(f)));

    children.push(h2('2. Test Strategy'));
    children.push(
      buildZebraTable(
        [
          { key: 'level', header: 'Level', weight: 1 },
          { key: 'description', header: 'Description', weight: 2 },
          { key: 'owner', header: 'Owner', weight: 1.2 },
        ],
        this.data.testStrategy,
        DEFAULT_TEST_STRATEGY,
      ),
    );

    children.push(h2('3. Test Environment'));
    children.push(
      buildZebraTable(
        [
          { key: 'environment', header: 'Environment', weight: 1 },
          { key: 'url', header: 'URL', weight: 1.3 },
          { key: 'database', header: 'Database', weight: 1.3 },
          { key: 'notes', header: 'Notes', weight: 1.3 },
        ],
        this.data.testEnvironments,
        DEFAULT_TEST_ENVIRONMENTS,
      ),
    );
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('3.1 Test Accounts & Devices'));
    (this.data.testAccounts && this.data.testAccounts.length
      ? this.data.testAccounts
      : DEFAULT_TEST_ACCOUNTS
    ).forEach((t) => children.push(bullet(t)));

    children.push(h2('4. Test Schedule'));
    children.push(
      buildZebraTable(
        [
          { key: 'activity', header: 'Activity', weight: 1.2 },
          { key: 'start', header: 'Start', weight: 1 },
          { key: 'end', header: 'End', weight: 1 },
        ],
        this.data.testSchedule,
        DEFAULT_TEST_SCHEDULE,
      ),
    );

    children.push(h2('5. Roles & Responsibilities'));
    children.push(
      buildZebraTable(
        [
          { key: 'role', header: 'Role', weight: 1 },
          { key: 'responsibility', header: 'Responsibility', weight: 2.5 },
        ],
        this.data.rolesResponsibilities,
        DEFAULT_ROLES_RESPONSIBILITIES,
      ),
    );

    children.push(h2('6. Test Cases'));
    children.push(
      buildZebraTable(
        [
          { key: 'id', header: 'ID', weight: 0.7 },
          { key: 'relatedUseCase', header: 'Related Use Case', weight: 1 },
          { key: 'steps', header: 'Steps', weight: 1.6 },
          { key: 'expectedResult', header: 'Expected Result', weight: 1.3 },
          { key: 'status', header: 'Status', weight: 0.9 },
        ],
        this.data.testCases,
        DEFAULT_TEST_CASES,
      ),
    );

    children.push(h2('7. Non-Functional Testing'));
    children.push(
      buildZebraTable(
        [
          { key: 'type', header: 'Type', weight: 1 },
          { key: 'criteria', header: 'Criteria', weight: 1.8 },
          { key: 'required', header: 'Required?', weight: 0.8 },
          { key: 'status', header: 'Status', weight: 0.9 },
        ],
        this.data.nonFunctionalTesting,
        DEFAULT_NON_FUNCTIONAL_TESTING,
      ),
    );

    children.push(h2('8. Bug Severity Definitions'));
    children.push(
      buildZebraTable(
        [
          { key: 'severity', header: 'Severity', weight: 1 },
          { key: 'definition', header: 'Definition', weight: 2.8 },
        ],
        this.data.bugSeverityDefinitions,
        DEFAULT_BUG_SEVERITY,
      ),
    );

    children.push(h2('9. Sign-off'));
    children.push(
      buildZebraTable(
        [
          { key: 'role', header: 'Role', weight: 1.2 },
          { key: 'name', header: 'Name', weight: 1 },
          { key: 'approved', header: 'Approved', weight: 1 },
          { key: 'date', header: 'Date', weight: 1 },
        ],
        this.data.signoff,
        DEFAULT_SIGNOFF,
      ),
    );

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
            run('UAT', { bold: true, size: SIZE.headerFooter, color: COLOR.headerFooterText }),
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

module.exports = UatSDK;
module.exports.SECTION_GUIDE = SECTION_GUIDE;
module.exports.buildZebraTable = buildZebraTable;
module.exports.BULLET_NUMBERING_REF = BULLET_NUMBERING_REF;
