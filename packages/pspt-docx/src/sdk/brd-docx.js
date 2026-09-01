'use strict';
/**
 * =============================================================================
 *  BUSINESS REQUIREMENTS DOCUMENT / BRD (Part 3 of the Product Documentation Master)
 * =============================================================================
 * Regenerates "Part 3. Business Requirements Document (BRD)" exactly: title
 * page, metadata table, 11 numbered sections including a Business
 * Requirements table, a Business Process sub-chapter (Overview / Scope /
 * RACI / Flow Diagram / Exceptions), a Use Cases sub-chapter (Actors / Use
 * Case Diagram / Use Case List / Detailed Specs / Traceability Matrix),
 * Assumptions & Dependencies, Risks, Approval/Sign-off, and Revision History.
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

const BULLET_NUMBERING_REF = 'brd-bullet-list';

/** Single-level bullet paragraph bound to this Part's numbering reference. */
const bullet = makeBullet(BULLET_NUMBERING_REF);

function boldPara(text) {
  return new Paragraph({
    spacing: { before: 160, after: 60, line: 276, lineRule: 'auto' },
    children: [run(text, { bold: true })],
  });
}

/** RACI-style table: "Step" column + one column per role, all navy-headered, zebra body. */
function buildRaciTable(roles, steps) {
  const columns = [{ key: 'step', header: 'Step', weight: 1.3 }].concat(
    roles.map((role, i) => ({ key: `role${i}`, header: role, weight: 1 })),
  );
  const rows = steps.map((s) => {
    const row = { step: s.step };
    (s.values || []).forEach((v, i) => {
      row[`role${i}`] = v;
    });
    return row;
  });
  return buildZebraTable(columns, rows, { step: PLACEHOLDER_ROW_TEXT });
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
const DEFAULT_PURPOSE = '[One-paragraph purpose statement.]';
const DEFAULT_BUSINESS_OBJECTIVES = [
  '[e.g. Reduce average process time from X to Y]',
  '[e.g. Achieve Z% adoption by month N post-launch]',
];
const DEFAULT_SCOPE = {
  inScope: '[business processes affected]',
  outOfScope: '[processes explicitly excluded]',
};
const DEFAULT_PROJECT_ROLES = [
  { name: '[Name]', role: '[Role]', responsibility: '[Responsibility]' },
];
const DEFAULT_BUSINESS_REQUIREMENTS = [
  {
    id: 'BR-001',
    requirement: '[The system shall...]',
    objective: 'Objective 1',
    priority: 'Must-have',
  },
  {
    id: 'BR-002',
    requirement: '[The system shall...]',
    objective: 'Objective 1',
    priority: 'Should-have',
  },
];
const DEFAULT_PROCESS_OVERVIEW =
  '[Name the process; describe its goal, trigger, and outcome in a short paragraph.]';
const DEFAULT_PROCESS_SCOPE =
  '[State exactly where this process starts and stops, since business processes often connect to other processes.]';
const DEFAULT_RACI_ROLES = ['[Role A]', '[Role B]', '[Role C]'];
const DEFAULT_RACI_STEPS = [
  { step: '[Step 1]', values: ['R/A', 'I', '—'] },
  { step: '[Step 2]', values: ['I', 'R/A', '—'] },
];
const DEFAULT_EXCEPTIONS = [
  '[What happens when a step is rejected, times out, or has missing information]',
];
const DEFAULT_ACTORS = [
  { actor: '[Actor]', description: '[Description — a human role or another system]' },
];
const DEFAULT_USE_CASE_LIST = [
  { id: 'UC-01', name: '[Use case name]', actor: '[Actor]', description: '[One-line description]' },
];
const DEFAULT_USE_CASE_SPECS = [
  {
    id: 'UC-01',
    name: '[Use Case Name]',
    primaryActor: '[actor]',
    preconditions: '[state that must be true before this use case starts]',
    postconditions: '[state after successful completion]',
    mainFlow: '1. [step] 2. [step] 3. [step]',
    alternateFlow: '[condition] → [what happens instead]',
    exceptionFlow: '[condition] → [error behaviour]',
    businessRules: '[rule, referencing Section 6.6 above]',
    relatedRequirements: 'FR-MOD-001 (Part 5)',
  },
];
const DEFAULT_TRACEABILITY = [{ useCase: 'UC-01', requirements: 'FR-MOD-001' }];
const DEFAULT_ASSUMPTIONS = [
  '[Assumption]',
  "[Dependency outside the dev team's control — e.g. another system's API availability, a policy decision, by [date]]",
];
const DEFAULT_RISKS = [
  "[Business risk to raise before sign-off, if any — otherwise 'None at this stage; see PRD Part 4 §8']",
];
const DEFAULT_APPROVALS = [
  { name: '[name]', role: '[role]', signature: '__________', date: '[date]' },
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
  { method: 'setPurpose', purpose: 'Section 1 purpose paragraph.', example: 'This BRD defines...' },
  {
    method: 'setBusinessObjectives',
    purpose: 'Section 2 bulleted objectives.',
    example: ['Reduce processing time from 3 days to 1'],
  },
  {
    method: 'setScope',
    purpose: 'Section 3 in/out of scope statement.',
    example: { inScope: 'Expense workflows', outOfScope: 'Payroll' },
  },
  {
    method: 'setProjectRoles',
    purpose: 'Section 4 Name/Role/Responsibility table.',
    example: [{ name: 'Jane Doe', role: 'Sponsor', responsibility: 'Approval' }],
  },
  {
    method: 'setBusinessRequirements',
    purpose: 'Section 5 ID/Requirement/Objective/Priority table.',
    example: [
      {
        id: 'BR-001',
        requirement: 'The system shall...',
        objective: 'Objective 1',
        priority: 'Must-have',
      },
    ],
  },
  {
    method: 'setProcessOverview',
    purpose: 'Section 6.1 process description.',
    example: 'The approval process starts when...',
  },
  {
    method: 'setProcessScope',
    purpose: 'Section 6.2 process boundaries.',
    example: 'Starts at submission, ends at payment.',
  },
  {
    method: 'setRaci',
    purpose: 'Section 6.3 RACI table.',
    example: { roles: ['Employee', 'Manager'], steps: [{ step: 'Submit', values: ['R/A', 'I'] }] },
  },
  {
    method: 'setExceptions',
    purpose: 'Section 6.5 bulleted exceptions/edge cases.',
    example: ['Missing receipt triggers manual review'],
  },
  {
    method: 'setActors',
    purpose: 'Section 7.1 Actor/Description table.',
    example: [{ actor: 'Employee', description: 'Submits expenses' }],
  },
  {
    method: 'setUseCaseList',
    purpose: 'Section 7.3 ID/Name/Actor/Description table.',
    example: [
      {
        id: 'UC-01',
        name: 'Submit expense',
        actor: 'Employee',
        description: 'Creates a new expense',
      },
    ],
  },
  {
    method: 'setUseCaseSpecs',
    purpose: 'Section 7.4 detailed use case blocks.',
    example: [
      {
        id: 'UC-01',
        name: 'Submit expense',
        primaryActor: 'Employee',
        preconditions: 'Logged in',
        postconditions: 'Saved',
        mainFlow: '1. Open 2. Fill 3. Submit',
        alternateFlow: 'n/a',
        exceptionFlow: 'Validation fails → show error',
        businessRules: 'Amount > 0',
        relatedRequirements: 'FR-EXP-001',
      },
    ],
  },
  {
    method: 'setTraceabilityMatrix',
    purpose: 'Section 7.5 Use Case/Requirement(s) table.',
    example: [{ useCase: 'UC-01', requirements: 'FR-EXP-001' }],
  },
  {
    method: 'setAssumptionsDependencies',
    purpose: 'Section 8 bulleted list.',
    example: ['Bank API is available 99.9% of the time'],
  },
  { method: 'setRisks', purpose: 'Section 9 bulleted list.', example: ['None at this stage'] },
  {
    method: 'setApprovals',
    purpose: 'Section 10 Name/Role/Signature/Date table.',
    example: [{ name: 'Jane Doe', role: 'Sponsor', signature: '__________', date: 'Aug 31, 2026' }],
  },
  {
    method: 'setRevisionHistory',
    purpose: 'Section 11 Version/Date/Author/Changes table.',
    example: [
      { version: 'v0.1', date: 'Aug 31, 2026', author: 'Jane Doe', changes: 'Initial draft' },
    ],
  },
];

class BrdSDK {
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
  setPurpose(text) {
    this.data.purpose = text;
    return this;
  }
  setBusinessObjectives(items) {
    this.data.businessObjectives = items;
    return this;
  }
  setScope(p = {}) {
    this.data.scope = p;
    return this;
  }
  setProjectRoles(rows) {
    this.data.projectRoles = rows;
    return this;
  }
  setBusinessRequirements(rows) {
    this.data.businessRequirements = rows;
    return this;
  }
  setProcessOverview(text) {
    this.data.processOverview = text;
    return this;
  }
  setProcessScope(text) {
    this.data.processScope = text;
    return this;
  }
  setRaci(p = {}) {
    this.data.raci = p;
    return this;
  }
  setExceptions(items) {
    this.data.exceptions = items;
    return this;
  }
  setActors(rows) {
    this.data.actors = rows;
    return this;
  }
  setUseCaseList(rows) {
    this.data.useCaseList = rows;
    return this;
  }
  setUseCaseSpecs(rows) {
    this.data.useCaseSpecs = rows;
    return this;
  }
  setTraceabilityMatrix(rows) {
    this.data.traceabilityMatrix = rows;
    return this;
  }
  setAssumptionsDependencies(items) {
    this.data.assumptionsDependencies = items;
    return this;
  }
  setRisks(items) {
    this.data.risks = items;
    return this;
  }
  setApprovals(rows) {
    this.data.approvals = rows;
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
        children: [headingRun('Business Requirements Document (BRD)')],
      }),
    ];
  }

  _buildContent() {
    const children = [];
    children.push(h1('Business Requirements Document (BRD)'));

    const meta = Object.assign({}, DEFAULT_METADATA, this.data.metadata || {});
    children.push(this._metadataTable(meta));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));

    children.push(h2('1. Purpose'));
    children.push(
      guidanceNote(
        "State why this document exists and who should read it. This BRD translates the Project Brief's goals (Part 2) into concrete, verifiable business requirements that Sales/Marketing/Ops and Development agree on, before detailed technical design in the SRS (Part 5).",
      ),
    );
    children.push(bodyPara(this.data.purpose || DEFAULT_PURPOSE));

    children.push(h2('2. Business Objectives'));
    children.push(
      guidanceNote(
        'Restate objectives from the Project Brief, tied to a business metric where possible (reduce X by Y%, cut process time from A to B).',
      ),
    );
    (this.data.businessObjectives && this.data.businessObjectives.length
      ? this.data.businessObjectives
      : DEFAULT_BUSINESS_OBJECTIVES
    ).forEach((o) => children.push(bullet(o)));

    children.push(h2('3. Scope'));
    const scope = Object.assign({}, DEFAULT_SCOPE, this.data.scope || {});
    children.push(bodyPara(`In scope: ${scope.inScope}. Out of scope: ${scope.outOfScope}.`));

    children.push(h2('4. Project Roles'));
    children.push(
      guidanceNote(
        'One row per person/role named there; explain each responsibility in the project.',
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'name', header: 'Name', weight: 1 },
          { key: 'role', header: 'Role', weight: 1 },
          { key: 'responsibility', header: 'Responsibility', weight: 2 },
        ],
        this.data.projectRoles,
        DEFAULT_PROJECT_ROLES[0],
      ),
    );

    children.push(h2('5. Business Requirements'));
    children.push(
      guidanceNote(
        "Numbered from the business's point of view ('the system shall allow...'), each is traceable to an objective above. Use a consistent ID scheme (BR-001, BR-002, ...) — these are referenced later from the SRS (as FR-xxx) and the Test Plan (as TC-xxx).",
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'id', header: 'ID', weight: 0.7 },
          { key: 'requirement', header: 'Requirement', weight: 2 },
          { key: 'objective', header: 'Related Objective', weight: 1 },
          { key: 'priority', header: 'Priority', weight: 0.8 },
        ],
        this.data.businessRequirements,
        DEFAULT_BUSINESS_REQUIREMENTS[0],
      ),
    );

    children.push(h2('6. Business Process'));
    children.push(
      guidanceNote(
        'Describes how work actually flows across roles and systems, end to end, to fulfil the requirements above — often spanning several use cases from Section 7 below. Especially useful for spotting handoff problems (a step no one owns, a bottleneck, a manual workaround).',
      ),
    );

    children.push(h3('6.1 Process Overview'));
    children.push(bodyPara(this.data.processOverview || DEFAULT_PROCESS_OVERVIEW));

    children.push(h3('6.2 Scope & Boundaries'));
    children.push(bodyPara(this.data.processScope || DEFAULT_PROCESS_SCOPE));

    children.push(h3('6.3 Roles & Responsibilities (RACI)'));
    children.push(
      guidanceNote(
        'List the business roles, departments, or systems involved in the process—not individual names. Use functional titles (e.g., Customer, Sales Representative, Finance Officer, System Administrator, ERP System). Each role should represent a participant who performs, approves, supports, or is informed about one or more process steps.',
      ),
    );
    const raci = this.data.raci || {};
    children.push(
      buildRaciTable(
        raci.roles && raci.roles.length ? raci.roles : DEFAULT_RACI_ROLES,
        raci.steps && raci.steps.length ? raci.steps : DEFAULT_RACI_STEPS,
      ),
    );

    children.push(h3('6.4 Process Flow Diagram'));
    children.push(
      guidanceNote(
        'A flowchart or swim-lane diagram (one lane per role/system) showing steps, decision points, and handoffs.',
      ),
    );
    children.push(bodyPara('[Insert swim-lane diagram here]'));

    children.push(h3('6.5 Exceptions & Edge Cases'));
    (this.data.exceptions && this.data.exceptions.length
      ? this.data.exceptions
      : DEFAULT_EXCEPTIONS
    ).forEach((e) => children.push(bullet(e)));

    children.push(h2('7. Use Cases'));
    children.push(
      guidanceNote(
        'Use cases describe how actors interact with the system to achieve a goal, step by step — they make the requirements in Section 5 and the process in Section 6 concrete, and are the direct source for test scenarios in the Test Plan (Part 11).',
      ),
    );

    children.push(h3('7.1 Actors'));
    children.push(
      buildZebraTable(
        [
          { key: 'actor', header: 'Actor', weight: 1 },
          { key: 'description', header: 'Description', weight: 2.5 },
        ],
        this.data.actors,
        DEFAULT_ACTORS[0],
      ),
    );

    children.push(h3('7.2 Use Case Diagram'));
    children.push(
      guidanceNote(
        'Insert a UML use case diagram showing actors, use cases, and <<include>>/<<extend>> relationships.',
      ),
    );
    children.push(bodyPara('[Insert diagram here]'));

    children.push(h3('7.3 Use Case List'));
    children.push(
      buildZebraTable(
        [
          { key: 'id', header: 'ID', weight: 0.6 },
          { key: 'name', header: 'Use Case Name', weight: 1.3 },
          { key: 'actor', header: 'Primary Actor', weight: 1 },
          { key: 'description', header: 'Description', weight: 1.6 },
        ],
        this.data.useCaseList,
        DEFAULT_USE_CASE_LIST[0],
      ),
    );

    children.push(h3('7.4 Detailed Use Case Specifications'));
    children.push(
      guidanceNote(
        'Repeat this block for every ID in Section 7.3. Detailed enough that a developer could implement it, and QA could write test cases directly from the flows.',
      ),
    );
    const specs =
      this.data.useCaseSpecs && this.data.useCaseSpecs.length
        ? this.data.useCaseSpecs
        : DEFAULT_USE_CASE_SPECS;
    specs.forEach((spec) => {
      children.push(boldPara(`${spec.id}: ${spec.name}`));
      children.push(bullet(`Primary Actor: ${spec.primaryActor}`));
      children.push(bullet(`Preconditions: ${spec.preconditions}`));
      children.push(bullet(`Postconditions (success): ${spec.postconditions}`));
      children.push(bullet(`Main Flow: ${spec.mainFlow}`));
      children.push(bullet(`Alternate Flow (A1): ${spec.alternateFlow}`));
      children.push(bullet(`Exception Flow (E1): ${spec.exceptionFlow}`));
      children.push(bullet(`Business Rules: ${spec.businessRules}`));
      children.push(bullet(`Related Requirements: ${spec.relatedRequirements}`));
    });

    children.push(h3('7.5 Traceability Matrix'));
    children.push(
      buildZebraTable(
        [
          { key: 'useCase', header: 'Use Case', weight: 1 },
          { key: 'requirements', header: 'Requirement(s)', weight: 2 },
        ],
        this.data.traceabilityMatrix,
        DEFAULT_TRACEABILITY[0],
      ),
    );

    children.push(h2('8. Assumptions & Dependencies'));
    (this.data.assumptionsDependencies && this.data.assumptionsDependencies.length
      ? this.data.assumptionsDependencies
      : DEFAULT_ASSUMPTIONS
    ).forEach((a) => children.push(bullet(a)));

    children.push(h2('9. Risks'));
    children.push(
      guidanceNote(
        "The project's one Risk Register lives in the PRD (Part 4, Section 8), which carries Likelihood/Impact/Mitigation for every risk regardless of type. Log business-level risks there directly rather than keeping a second list here; this section exists only to flag, at BRD sign-off time, whether any new business risk needs raising before requirements are approved.",
      ),
    );
    (this.data.risks && this.data.risks.length ? this.data.risks : DEFAULT_RISKS).forEach((r) =>
      children.push(bullet(r)),
    );

    children.push(h2('10. Approval / Sign-off'));
    children.push(
      guidanceNote(
        'Once signed, this BRD becomes the baseline for the SRS (Part 5); changes after sign-off should go through a change-request process.',
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'name', header: 'Name', weight: 1 },
          { key: 'role', header: 'Role', weight: 1 },
          { key: 'signature', header: 'Signature', weight: 1 },
          { key: 'date', header: 'Date', weight: 1 },
        ],
        this.data.approvals,
        DEFAULT_APPROVALS[0],
      ),
    );

    children.push(h2('11. Revision History'));
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
          headers: { default: this._chapterHeader('BRD') },
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

module.exports = BrdSDK;
module.exports.SECTION_GUIDE = SECTION_GUIDE;
module.exports.buildZebraTable = buildZebraTable;
module.exports.buildRaciTable = buildRaciTable;
module.exports.BULLET_NUMBERING_REF = BULLET_NUMBERING_REF;
