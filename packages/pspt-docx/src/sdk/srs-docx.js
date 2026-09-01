'use strict';
/**
 * =============================================================================
 *  SOFTWARE REQUIREMENTS SPECIFICATION / SRS (Part 5 of the Product Documentation Master)
 * =============================================================================
 * Regenerates "Part 5. Software Requirements Specification (SRS)" exactly:
 * title page, metadata table, 9 numbered sections (Introduction with 1.1-1.3
 * sub-sections, Overall Description with 2.1-2.6 sub-sections, Functional
 * Requirements, Non-Functional Requirements, External Interface Requirements
 * with 5.1-5.4 sub-sections, System Features, Other Requirements, Appendix,
 * Revision History).
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

const BULLET_NUMBERING_REF = 'srs-bullet-list';

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
const DEFAULT_PURPOSE_AUDIENCE =
  '[Purpose of this SRS; intended audience — development and QA teams during implementation and testing.]';
const DEFAULT_SCOPE =
  '[Briefly restate the scope from the BRD, Part 3 Section 3, referencing it rather than repeating it in full.]';
const DEFAULT_REFERENCES = ['Project Brief (Part 2) v[x]', 'BRD (Part 3) v[x], approved [date]'];
const DEFAULT_PRODUCT_PERSPECTIVE =
  '[How this system fits into its environment — standalone, or integrated with other systems. See Part 6, Section 1.2 for the architecture diagram.]';
const DEFAULT_PRODUCT_FUNCTIONS = ['[Function 1 — one line]', '[Function 2 — one line]'];
const DEFAULT_USER_CLASSES = [
  { userClass: '[User class]', description: '[Description]', technicalLevel: '[Technical level]' },
];
const DEFAULT_OPERATING_ENV = [
  'Browsers: [supported browsers]',
  'Server: [OS/runtime]',
  'Devices: [desktop/mobile/tablet — see Part 7, Section 8 for responsive breakpoints]',
];
const DEFAULT_CONSTRAINTS = [
  '[e.g. Must use existing infrastructure]',
  '[Must comply with Part 6 §4 — Security Requirements]',
];
const DEFAULT_ASSUMPTIONS = ['[Assumption]'];
const DEFAULT_FUNCTIONAL_REQUIREMENTS = [
  { id: 'FR-MOD-001', requirement: '[The system shall...]', tracesTo: 'BR-001' },
  { id: 'FR-MOD-002', requirement: '[The system shall...]', tracesTo: 'BR-002' },
];
const DEFAULT_NFR = [
  {
    category: 'Performance',
    requirement: '[e.g. Dashboard loads within 2s for up to 50 concurrent users]',
  },
  { category: 'Security', requirement: '[See Part 6 §4 — Security Requirements]' },
  { category: 'Usability', requirement: '[e.g. Core workflow completable within 5 clicks]' },
  { category: 'Reliability', requirement: '[e.g. 99.5% uptime during business hours]' },
  {
    category: 'Scalability',
    requirement: '[e.g. Supports growth to N transactions/month without architecture changes]',
  },
];
const DEFAULT_USER_INTERFACES = [
  { screenId: 'SCR-01', screenName: '[Screen name]', relatedRequirements: 'FR-MOD-001' },
];
const DEFAULT_HARDWARE_INTERFACES =
  'None — standard web application with no dedicated hardware interfaces. [Update if not applicable.]';
const DEFAULT_SOFTWARE_INTERFACES =
  '[Other systems this software talks to — full contract in Part 6 §3, API Specification.]';
const DEFAULT_COMMUNICATION_INTERFACES =
  '[Network protocols, data formats, security — e.g. HTTPS only (TLS 1.2+), JSON payloads.]';
const DEFAULT_SYSTEM_FEATURE =
  'Feature: [Name] — Description: [...]. Precondition: [...]. Main Flow: [...]. Result: [...].';
const DEFAULT_OTHER_REQUIREMENTS = [
  '[Legal, compliance, or localisation requirement not covered elsewhere]',
];
const DEFAULT_APPENDIX = 'Open issue: [description, tracked as ISSUE-xx].';
const DEFAULT_REVISION_HISTORY = [
  { version: 'v0.1', date: '[date]', author: '[name]', changes: 'Initial draft' },
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
    method: 'setPurposeAudience',
    purpose: 'Section 1.1 purpose & audience.',
    example: 'This SRS defines testable requirements for engineering and QA.',
  },
  {
    method: 'setScope',
    purpose: 'Section 1.2 scope.',
    example: 'Covers the expense submission and approval modules only.',
  },
  {
    method: 'setReferences',
    purpose: 'Section 1.3 bulleted references.',
    example: ['Project Brief v1.0', 'BRD v1.2, approved Aug 1'],
  },
  {
    method: 'setProductPerspective',
    purpose: 'Section 2.1 product perspective.',
    example: 'Standalone SaaS product integrating with the company SSO.',
  },
  {
    method: 'setProductFunctions',
    purpose: 'Section 2.2 bulleted functions.',
    example: ['Submit expenses', 'Approve/reject expenses'],
  },
  {
    method: 'setUserClasses',
    purpose: 'Section 2.3 User Class/Description/Technical Level table.',
    example: [{ userClass: 'Employee', description: 'Submits expenses', technicalLevel: 'Basic' }],
  },
  {
    method: 'setOperatingEnvironment',
    purpose: 'Section 2.4 bulleted environment notes.',
    example: ['Browsers: Chrome, Safari, Edge (last 2 versions)'],
  },
  {
    method: 'setConstraints',
    purpose: 'Section 2.5 bulleted constraints.',
    example: ['Must use existing SSO infrastructure'],
  },
  {
    method: 'setAssumptionsDependencies',
    purpose: 'Section 2.6 bulleted assumptions.',
    example: ['Bank API is available 99.9% of the time'],
  },
  {
    method: 'setFunctionalRequirements',
    purpose: 'Section 3 ID/Requirement/Traces to table.',
    example: [
      {
        id: 'FR-MOD-001',
        requirement: 'The system shall allow users to submit expenses',
        tracesTo: 'BR-001',
      },
    ],
  },
  {
    method: 'setNonFunctionalRequirements',
    purpose: 'Section 4 Category/Requirement table.',
    example: [{ category: 'Performance', requirement: 'Dashboard loads within 2s' }],
  },
  {
    method: 'setUserInterfaces',
    purpose: 'Section 5.1 Screen ID/Name/Related Requirements table.',
    example: [
      {
        screenId: 'SCR-01',
        screenName: 'Expense submission form',
        relatedRequirements: 'FR-MOD-001',
      },
    ],
  },
  {
    method: 'setHardwareInterfaces',
    purpose: 'Section 5.2 hardware interfaces note.',
    example: 'None — standard web application.',
  },
  {
    method: 'setSoftwareInterfaces',
    purpose: 'Section 5.3 software interfaces note.',
    example: 'Integrates with the Xero accounting API.',
  },
  {
    method: 'setCommunicationInterfaces',
    purpose: 'Section 5.4 communication interfaces note.',
    example: 'HTTPS only (TLS 1.2+), JSON payloads.',
  },
  {
    method: 'setSystemFeature',
    purpose: 'Section 6 detailed feature spec paragraph.',
    example:
      'Feature: Expense OCR — Description: Extracts amount and vendor from receipt photos...',
  },
  {
    method: 'setOtherRequirements',
    purpose: 'Section 7 bulleted list.',
    example: ['Must comply with GDPR for EU customers'],
  },
  {
    method: 'setAppendix',
    purpose: 'Section 8 open issues paragraph.',
    example: 'Open issue: OCR accuracy for handwritten receipts, tracked as ISSUE-42.',
  },
  {
    method: 'setRevisionHistory',
    purpose: 'Section 9 Version/Date/Author/Changes table.',
    example: [
      { version: 'v0.1', date: 'Aug 31, 2026', author: 'Jane Doe', changes: 'Initial draft' },
    ],
  },
];

class SrsSDK {
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
  setPurposeAudience(text) {
    this.data.purposeAudience = text;
    return this;
  }
  setScope(text) {
    this.data.scope = text;
    return this;
  }
  setReferences(items) {
    this.data.references = items;
    return this;
  }
  setProductPerspective(text) {
    this.data.productPerspective = text;
    return this;
  }
  setProductFunctions(items) {
    this.data.productFunctions = items;
    return this;
  }
  setUserClasses(rows) {
    this.data.userClasses = rows;
    return this;
  }
  setOperatingEnvironment(items) {
    this.data.operatingEnvironment = items;
    return this;
  }
  setConstraints(items) {
    this.data.constraints = items;
    return this;
  }
  setAssumptionsDependencies(items) {
    this.data.assumptionsDependencies = items;
    return this;
  }
  setFunctionalRequirements(rows) {
    this.data.functionalRequirements = rows;
    return this;
  }
  setNonFunctionalRequirements(rows) {
    this.data.nonFunctionalRequirements = rows;
    return this;
  }
  setUserInterfaces(rows) {
    this.data.userInterfaces = rows;
    return this;
  }
  setHardwareInterfaces(text) {
    this.data.hardwareInterfaces = text;
    return this;
  }
  setSoftwareInterfaces(text) {
    this.data.softwareInterfaces = text;
    return this;
  }
  setCommunicationInterfaces(text) {
    this.data.communicationInterfaces = text;
    return this;
  }
  setSystemFeature(text) {
    this.data.systemFeature = text;
    return this;
  }
  setOtherRequirements(items) {
    this.data.otherRequirements = items;
    return this;
  }
  setAppendix(text) {
    this.data.appendix = text;
    return this;
  }
  setRevisionHistory(rows) {
    this.data.revisionHistory = rows;
    return this;
  }

  _buildPartTitlePage() {
    return [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [headingRun('Software Requirements Specification (SRS)')],
      }),
    ];
  }

  _buildContent() {
    const children = [];
    children.push(h1('Software Requirements Specification (SRS)'));

    const meta = Object.assign({}, DEFAULT_METADATA, this.data.metadata || {});
    children.push(this._metadataTable(meta));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));

    children.push(h2('1. Introduction'));
    children.push(h3('1.1 Purpose & Audience'));
    children.push(bodyPara(this.data.purposeAudience || DEFAULT_PURPOSE_AUDIENCE));
    children.push(h3('1.2 Scope'));
    children.push(bodyPara(this.data.scope || DEFAULT_SCOPE));
    children.push(h3('1.3 References'));
    (this.data.references && this.data.references.length
      ? this.data.references
      : DEFAULT_REFERENCES
    ).forEach((r) => children.push(bullet(r)));

    children.push(h2('2. Overall Description'));
    children.push(h3('2.1 Product Perspective'));
    children.push(bodyPara(this.data.productPerspective || DEFAULT_PRODUCT_PERSPECTIVE));
    children.push(h3('2.2 Product Functions'));
    (this.data.productFunctions && this.data.productFunctions.length
      ? this.data.productFunctions
      : DEFAULT_PRODUCT_FUNCTIONS
    ).forEach((f) => children.push(bullet(f)));
    children.push(h3('2.3 User Classes & Characteristics'));
    children.push(
      guidanceNote(
        "Same audience as the PRD's personas (Part 4, Section 2), reduced to the technical-access dimension that drives role-based design — don't re-describe needs/pain points here, just link the class to its persona.",
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'userClass', header: 'User Class', weight: 1 },
          { key: 'description', header: 'Description', weight: 1.6 },
          { key: 'technicalLevel', header: 'Technical Level', weight: 1 },
        ],
        this.data.userClasses,
        DEFAULT_USER_CLASSES,
      ),
    );
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('2.4 Operating Environment'));
    (this.data.operatingEnvironment && this.data.operatingEnvironment.length
      ? this.data.operatingEnvironment
      : DEFAULT_OPERATING_ENV
    ).forEach((o) => children.push(bullet(o)));
    children.push(h3('2.5 Constraints'));
    (this.data.constraints && this.data.constraints.length
      ? this.data.constraints
      : DEFAULT_CONSTRAINTS
    ).forEach((c) => children.push(bullet(c)));
    children.push(h3('2.6 Assumptions & Dependencies'));
    (this.data.assumptionsDependencies && this.data.assumptionsDependencies.length
      ? this.data.assumptionsDependencies
      : DEFAULT_ASSUMPTIONS
    ).forEach((a) => children.push(bullet(a)));

    children.push(h2('3. Functional Requirements'));
    children.push(
      guidanceNote(
        'Atomic (one behaviour per line), testable, and traceable to a BRD requirement ID.',
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'id', header: 'ID', weight: 0.9 },
          { key: 'requirement', header: 'Requirement', weight: 2.2 },
          { key: 'tracesTo', header: 'Traces to', weight: 0.9 },
        ],
        this.data.functionalRequirements,
        DEFAULT_FUNCTIONAL_REQUIREMENTS,
      ),
    );

    children.push(h2('4. Non-Functional Requirements'));
    children.push(
      buildZebraTable(
        [
          { key: 'category', header: 'Category', weight: 1 },
          { key: 'requirement', header: 'Requirement', weight: 2.8 },
        ],
        this.data.nonFunctionalRequirements,
        DEFAULT_NFR,
      ),
    );

    children.push(h2('5. External Interface Requirements'));
    children.push(h3('5.1 User Interfaces'));
    children.push(
      guidanceNote(
        'Screen inventory, navigation flow, responsive behaviour, and shared components — full visual specs live in Part 7 (UI/UX Documentation).',
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'screenId', header: 'Screen ID', weight: 0.9 },
          { key: 'screenName', header: 'Screen Name', weight: 1.6 },
          { key: 'relatedRequirements', header: 'Related Requirement(s)', weight: 1.3 },
        ],
        this.data.userInterfaces,
        DEFAULT_USER_INTERFACES,
      ),
    );
    children.push(h3('5.2 Hardware Interfaces'));
    children.push(bodyPara(this.data.hardwareInterfaces || DEFAULT_HARDWARE_INTERFACES));
    children.push(h3('5.3 Software Interfaces'));
    children.push(bodyPara(this.data.softwareInterfaces || DEFAULT_SOFTWARE_INTERFACES));
    children.push(h3('5.4 Communication Interfaces'));
    children.push(bodyPara(this.data.communicationInterfaces || DEFAULT_COMMUNICATION_INTERFACES));

    children.push(h2('6. System Features (Detailed Feature Specifications)'));
    children.push(
      guidanceNote(
        "For complex features where a single requirement row isn't enough detail: description, trigger/preconditions, main flow, and result. This bridges Section 3 above and the BRD's Use Cases section, Part 3 §7.",
      ),
    );
    children.push(bodyPara(this.data.systemFeature || DEFAULT_SYSTEM_FEATURE));

    children.push(h2('7. Other Requirements'));
    (this.data.otherRequirements && this.data.otherRequirements.length
      ? this.data.otherRequirements
      : DEFAULT_OTHER_REQUIREMENTS
    ).forEach((o) => children.push(bullet(o)));

    children.push(h2('8. Appendix'));
    children.push(bodyPara(this.data.appendix || DEFAULT_APPENDIX));

    children.push(h2('9. Revision History'));
    children.push(
      buildZebraTable(
        [
          { key: 'version', header: 'Version', weight: 0.8 },
          { key: 'date', header: 'Date', weight: 1 },
          { key: 'author', header: 'Author', weight: 1 },
          { key: 'changes', header: 'Changes', weight: 1.8 },
        ],
        this.data.revisionHistory,
        DEFAULT_REVISION_HISTORY,
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
          headers: { default: this._chapterHeader('SRS') },
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

module.exports = SrsSDK;
module.exports.SECTION_GUIDE = SECTION_GUIDE;
module.exports.buildZebraTable = buildZebraTable;
module.exports.BULLET_NUMBERING_REF = BULLET_NUMBERING_REF;
