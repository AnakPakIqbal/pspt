'use strict';
/**
 * =============================================================================
 *  PRODUCT REQUIREMENTS DOCUMENT / PRD (Part 4 of the Product Documentation Master)
 * =============================================================================
 * Regenerates "Part 4. Product Requirements Document (PRD)" exactly: title
 * page, metadata table, 10 numbered sections (Overview, Target Audience,
 * Requirements + Non-Functional sub-section, User Stories & Use Cases,
 * Design & UX, Technical Considerations, Success Metrics, Risks & Open
 * Questions, Roadmap, Revision History).
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

const BULLET_NUMBERING_REF = 'prd-bullet-list';

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
const DEFAULT_OVERVIEW =
  '[Summarise the product/feature, the problem it solves, and why it matters now.]';
const DEFAULT_AUDIENCE = [
  { persona: '[Primary persona]', needs: '[…]', painPoints: '[…]' },
  { persona: '[Secondary persona]', needs: '[…]', painPoints: '[…]' },
];
const DEFAULT_REQUIREMENTS = [
  {
    id: 'FR-01',
    feature: '[Feature name]',
    description: '[What it does and why]',
    priority: 'Must',
    status: 'Not started',
  },
  {
    id: 'FR-02',
    feature: '[Feature name]',
    description: '[What it does and why]',
    priority: 'Should',
    status: 'In progress',
  },
];
const DEFAULT_NFR_NOTES = [
  'Categories in scope for this feature: [e.g. Performance, Accessibility] — see Part 5 §4 for targets.',
  'Security & Privacy: see Part 6 §4 — Security Requirements.',
];
const DEFAULT_USER_STORIES = [
  {
    story: 'Story 1 — [As a [persona], I want to [action], so that [benefit].]',
    bullets: ['Given [context], when [action], then [result].', '[…]'],
  },
];
const DEFAULT_DESIGN_UX = [
  'Wireframes/mockups: [link]',
  'Prototype: [link]',
  'Design system components used: [list]',
];
const DEFAULT_TECHNICAL_CONSIDERATIONS = [
  'Architecture & dependencies: [high-level note; full detail in Part 6 §1]',
  'Data & analytics: [events to track, data model changes, reporting needs]',
  'Assumptions: [assumption 1]',
];
const DEFAULT_SUCCESS_METRICS = [
  {
    metric: '[e.g. Activation rate]',
    baseline: '[current %]',
    target: '[target %]',
    owner: '[owner]',
  },
];
const DEFAULT_RISK_REGISTER = [
  {
    risk: '[Risk description]',
    likelihood: 'Low/Med/High',
    impact: 'Low/Med/High',
    mitigation: '[Plan]',
  },
];
const DEFAULT_OPEN_QUESTIONS = ['[Question] — Owner: [name] — Due: [date]', '[…]'];
const DEFAULT_ROADMAP = [
  { horizon: 'Now (this quarter)', focus: '[Key initiatives]' },
  { horizon: 'Next (next quarter)', focus: '[Key initiatives]' },
  { horizon: 'Later (6–12 months)', focus: '[Directional themes]' },
];
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
    method: 'setOverview',
    purpose: 'Section 1: 2-3 sentence overview.',
    example: 'Acme Widget solves...',
  },
  {
    method: 'setTargetAudience',
    purpose: 'Section 2 Persona/Needs/Pain Points table.',
    example: [{ persona: 'Busy Manager', needs: 'Save time', painPoints: 'Manual process' }],
  },
  {
    method: 'setRequirements',
    purpose: 'Section 3 ID/Feature/Description/Priority/Status table.',
    example: [
      {
        id: 'FR-01',
        feature: 'Push notifications',
        description: 'Real-time alerts',
        priority: 'Must',
        status: 'Not started',
      },
    ],
  },
  {
    method: 'setNfrNotes',
    purpose: 'Section 3.1 bulleted NFR category notes.',
    example: ['Categories: Performance, Security'],
  },
  {
    method: 'setUserStories',
    purpose: 'Section 4 story blocks.',
    example: [
      {
        story: 'Story 1 — As a user, I want to export data, so that I can analyze it.',
        bullets: ['Given a filled report, when I click export, then a CSV downloads.'],
      },
    ],
  },
  {
    method: 'setDesignUx',
    purpose: 'Section 5 bulleted design links.',
    example: ['Wireframes: figma.com/...'],
  },
  {
    method: 'setTechnicalConsiderations',
    purpose: 'Section 6 bulleted technical notes.',
    example: ['Uses existing auth service'],
  },
  {
    method: 'setSuccessMetrics',
    purpose: 'Section 7 Metric/Baseline/Target/Owner table.',
    example: [{ metric: 'Activation rate', baseline: '40%', target: '65%', owner: 'PM' }],
  },
  {
    method: 'setRiskRegister',
    purpose: 'Section 8.1 Risk/Likelihood/Impact/Mitigation table.',
    example: [
      { risk: 'API rate limits', likelihood: 'Medium', impact: 'High', mitigation: 'Add caching' },
    ],
  },
  {
    method: 'setOpenQuestions',
    purpose: 'Section 8.2 bulleted questions.',
    example: ['Do we need SSO at launch? — Owner: Jane — Due: Sept 15'],
  },
  {
    method: 'setRoadmap',
    purpose: 'Section 9 Horizon/Focus table.',
    example: [{ horizon: 'Now', focus: 'MVP launch' }],
  },
  {
    method: 'setRevisionHistory',
    purpose: 'Section 10 Version/Date/Author/Changes table.',
    example: [
      { version: 'v0.1', date: 'Aug 31, 2026', author: 'Jane Doe', changes: 'Initial draft' },
    ],
  },
];

class PrdSDK {
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
  setOverview(text) {
    this.data.overview = text;
    return this;
  }
  setTargetAudience(rows) {
    this.data.targetAudience = rows;
    return this;
  }
  setRequirements(rows) {
    this.data.requirements = rows;
    return this;
  }
  setNfrNotes(items) {
    this.data.nfrNotes = items;
    return this;
  }
  /** @param {Array<{story:string, bullets:string[]}>} items */
  setUserStories(items) {
    this.data.userStories = items;
    return this;
  }
  setDesignUx(items) {
    this.data.designUx = items;
    return this;
  }
  setTechnicalConsiderations(items) {
    this.data.technicalConsiderations = items;
    return this;
  }
  setSuccessMetrics(rows) {
    this.data.successMetrics = rows;
    return this;
  }
  setRiskRegister(rows) {
    this.data.riskRegister = rows;
    return this;
  }
  setOpenQuestions(items) {
    this.data.openQuestions = items;
    return this;
  }
  setRoadmap(rows) {
    this.data.roadmap = rows;
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
        children: [headingRun('Product Requirements Document (PRD)')],
      }),
    ];
  }

  _buildContent() {
    const children = [];
    children.push(h1('Product Requirements Document (PRD)'));

    const meta = Object.assign({}, DEFAULT_METADATA, this.data.metadata || {});
    children.push(this._metadataTable(meta));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));

    children.push(h2('1. Overview'));
    children.push(
      guidanceNote(
        "Two to three sentences, a busy exec could read and understand the 'what' and 'why'.",
      ),
    );
    children.push(bodyPara(this.data.overview || DEFAULT_OVERVIEW));

    children.push(h2('2. Target Audience'));
    children.push(
      guidanceNote('Primary and secondary user personas. Link to full persona docs if they exist.'),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'persona', header: 'Persona', weight: 1 },
          { key: 'needs', header: 'Needs', weight: 1.2 },
          { key: 'painPoints', header: 'Pain Points', weight: 1.2 },
        ],
        this.data.targetAudience,
        DEFAULT_AUDIENCE,
      ),
    );

    children.push(h2('3. Requirements'));
    children.push(
      guidanceNote(
        "Keep one row per requirement. Priority uses MoSCoW: Must / Should / Could / Won't. Detailed, testable versions of these live in the SRS (Part 5) as FR-xxx and are traced back to this ID.",
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'id', header: 'ID', weight: 0.6 },
          { key: 'feature', header: 'Feature', weight: 1.2 },
          { key: 'description', header: 'Description', weight: 1.6 },
          { key: 'priority', header: 'Priority', weight: 0.7 },
          { key: 'status', header: 'Status', weight: 0.9 },
        ],
        this.data.requirements,
        DEFAULT_REQUIREMENTS,
      ),
    );

    children.push(h3('3.1 Non-Functional Requirements'));
    children.push(
      guidanceNote(
        "The full, testable NFR table (with concrete targets) is maintained once, in the SRS (Part 5, Section 4), so performance/accessibility/scalability targets don't drift between two versions. List only which categories matter most for this feature; leave the numbers to Part 5.",
      ),
    );
    (this.data.nfrNotes && this.data.nfrNotes.length
      ? this.data.nfrNotes
      : DEFAULT_NFR_NOTES
    ).forEach((n) => children.push(bullet(n)));

    children.push(h2('4. User Stories & Use Cases'));
    children.push(
      guidanceNote(
        "Format: As a [persona], I want to [action], so that [benefit]. Full step-by-step flows, alternate/exception paths, and a traceability matrix live in the BRD's Use Cases section (Part 3, Section 7).",
      ),
    );
    (this.data.userStories && this.data.userStories.length
      ? this.data.userStories
      : DEFAULT_USER_STORIES
    ).forEach((s) => {
      children.push(bodyPara(s.story));
      (s.bullets || []).forEach((b) => children.push(bullet(b)));
    });

    children.push(h2('5. Design & User Experience'));
    children.push(
      guidanceNote(
        'Link to Figma/mockups rather than pasting images — see Part 7 for the full UX Spec.',
      ),
    );
    (this.data.designUx && this.data.designUx.length
      ? this.data.designUx
      : DEFAULT_DESIGN_UX
    ).forEach((d) => children.push(bullet(d)));

    children.push(h2('6. Technical Considerations'));
    children.push(
      guidanceNote(
        'High-level notes only — see Part 6 (Technical Documentation) for the full details: System Architecture (§1), Data Model (§2), and API Specification (§3).',
      ),
    );
    (this.data.technicalConsiderations && this.data.technicalConsiderations.length
      ? this.data.technicalConsiderations
      : DEFAULT_TECHNICAL_CONSIDERATIONS
    ).forEach((t) => children.push(bullet(t)));

    children.push(h2('7. Success Metrics'));
    children.push(
      guidanceNote(
        "This is the project's single measurable success table — tie every metric back to a goal in Section 1, include a baseline, and treat this as canonical. The Project Brief's Success Criteria (Part 2, §10) and the BRD's Business Process KPIs (Part 3, §6.8) are qualitative/operational views that point back here rather than keeping their own targets.",
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'metric', header: 'Metric', weight: 1.3 },
          { key: 'baseline', header: 'Baseline', weight: 1 },
          { key: 'target', header: 'Target', weight: 1 },
          { key: 'owner', header: 'Owner', weight: 1 },
        ],
        this.data.successMetrics,
        DEFAULT_SUCCESS_METRICS[0],
      ),
    );

    children.push(h2('8. Risks & Open Questions'));
    children.push(h3('8.1 Risk Register'));
    children.push(
      guidanceNote(
        "The project's single Risk Register carries forward preliminary risks flagged in the Project Brief (Part 2, §11) and business risks flagged in the BRD (Part 3, §9). Log every new risk here going forward, whatever its source.",
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'risk', header: 'Risk', weight: 1.5 },
          { key: 'likelihood', header: 'Likelihood', weight: 1 },
          { key: 'impact', header: 'Impact', weight: 1 },
          { key: 'mitigation', header: 'Mitigation', weight: 1.3 },
        ],
        this.data.riskRegister,
        DEFAULT_RISK_REGISTER[0],
      ),
    );

    children.push(h3('8.2 Open Questions'));
    (this.data.openQuestions && this.data.openQuestions.length
      ? this.data.openQuestions
      : DEFAULT_OPEN_QUESTIONS
    ).forEach((q) => children.push(bullet(q)));

    children.push(h2('9. Roadmap'));
    children.push(
      guidanceNote(
        'High-level roadmap by horizon. Keep detailed sprint planning in your project tracker, not here.',
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'horizon', header: 'Horizon', weight: 1 },
          { key: 'focus', header: 'Focus', weight: 2.5 },
        ],
        this.data.roadmap,
        DEFAULT_ROADMAP,
      ),
    );

    children.push(h2('10. Revision History'));
    children.push(
      buildZebraTable(
        [
          { key: 'version', header: 'Version', weight: 0.8 },
          { key: 'date', header: 'Date', weight: 1 },
          { key: 'author', header: 'Author', weight: 1 },
          { key: 'changes', header: 'Changes', weight: 1.8 },
        ],
        this.data.revisionHistory,
        DEFAULT_REVISION_HISTORY[0],
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
          headers: { default: this._chapterHeader('PRD') },
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

module.exports = PrdSDK;
module.exports.SECTION_GUIDE = SECTION_GUIDE;
module.exports.buildZebraTable = buildZebraTable;
module.exports.BULLET_NUMBERING_REF = BULLET_NUMBERING_REF;
