'use strict';
/**
 * =============================================================================
 *  DEPLOYMENT GUIDE (Part 9 of the Product Documentation Master)
 * =============================================================================
 * Regenerates "Part 9. Deployment Guide" exactly: title page, metadata
 * table, 8 numbered sections (Overview, Prerequisites, Environment Setup
 * with a shaded .env code block, Deployment Steps as a multi-command shell
 * code block, Configuration, Rollback Procedure, Post-Deployment
 * Verification checklist, Monitoring & Alerting).
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
  checklistItem,
  codeBlock,
  h1,
  h2,
  headingRun,
  outerBorderSet,
  run,
} = require('./pspt-core');

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
  "[What's being deployed, to where, and how often — continuous deployment vs manual releases.]";
const DEFAULT_PREREQUISITES = [
  { requirement: 'Server access', details: '[SSH key/cloud console access]' },
  { requirement: 'Runtime', details: '[e.g. Node.js v20+ installed on server]' },
  {
    requirement: 'Environment variables',
    details: '[.env values obtained from team password manager]',
  },
];
const DEFAULT_ENVIRONMENTS = [
  { environment: 'Development', purpose: 'Local development', branch: 'any feature branch' },
  { environment: 'Staging', purpose: 'Pre-release testing', branch: '`develop`' },
  { environment: 'Production', purpose: 'Live system', branch: '`main`' },
];
const DEFAULT_ENV_EXAMPLE = [
  '# .env.example — copy to .env and fill in real values',
  'DB_HOST=localhost',
  'DB_USER=app',
  'DB_PASSWORD=',
  'JWT_SECRET=',
];
const DEFAULT_DEPLOYMENT_STEPS = [
  '# 1. SSH into the server',
  'ssh deploy@production-server',
  '',
  '# 2. Pull the latest code',
  'cd /var/www/app && git pull origin main',
  '',
  '# 3. Install dependencies',
  'npm --prefix api install --production',
  'npm --prefix app install && npm --prefix app run build',
  '',
  '# 4. Run database migrations',
  'npm --prefix api run migrate',
  '',
  '# 5. Restart the service',
  'pm2 restart app-api',
];
const DEFAULT_CONFIGURATION = [
  { variable: '`DB_HOST`', description: 'Database host address', required: 'Yes' },
  { variable: '`JWT_SECRET`', description: 'Secret used to sign auth tokens', required: 'Yes' },
  { variable: '`SMTP_HOST`', description: 'Outbound email server', required: 'Yes' },
];
const DEFAULT_ROLLBACK_STEPS = [
  '# Roll back to the previous release tag',
  'git checkout <previous-release-tag>',
  'npm --prefix api install --production',
  'pm2 restart app-api',
];
const DEFAULT_ROLLBACK_NOTE =
  'If the rollback involves a database migration, restore from the pre-deployment backup taken in Step 4 above.';
const DEFAULT_POST_DEPLOY_CHECKLIST = [
  '`/api/v1/health` returns 200 OK',
  'Login works with a test account',
  'Recent log entries show no unexpected errors',
];
const DEFAULT_MONITORING = [
  {
    what: 'Server uptime',
    tool: '[e.g. PM2 + uptime monitor]',
    alertRecipient: 'Dev Lead (email/SMS)',
  },
  { what: 'Error logs', tool: '[Application log file]', alertRecipient: 'Dev team' },
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
    purpose: 'Section 1 overview paragraph.',
    example: 'Deployed to AWS ECS via GitHub Actions on every merge to main.',
  },
  {
    method: 'setPrerequisites',
    purpose: 'Section 2 Requirement/Details table.',
    example: [{ requirement: 'Server access', details: 'AWS console access with deploy role' }],
  },
  {
    method: 'setEnvironments',
    purpose: 'Section 3 Environment/Purpose/Branch table.',
    example: [{ environment: 'Staging', purpose: 'Pre-release testing', branch: 'develop' }],
  },
  {
    method: 'setEnvExample',
    purpose: 'Section 3 .env.example code block lines.',
    example: ['DB_HOST=localhost', 'API_KEY='],
  },
  {
    method: 'setDeploymentSteps',
    purpose:
      'Section 4 shell command code block lines (use empty strings for blank separator lines).',
    example: ['# 1. Build', 'npm run build', '', '# 2. Deploy', 'npm run deploy'],
  },
  {
    method: 'setConfiguration',
    purpose: 'Section 5 Variable/Description/Required table.',
    example: [{ variable: 'DB_HOST', description: 'Database host', required: 'Yes' }],
  },
  {
    method: 'setRollbackSteps',
    purpose: 'Section 6 rollback shell command code block lines.',
    example: ['git checkout v1.2.0', 'pm2 restart app'],
  },
  {
    method: 'setRollbackNote',
    purpose: 'Section 6 note after the code block.',
    example: 'Database rollbacks require restoring the pre-deploy snapshot.',
  },
  {
    method: 'setPostDeploymentChecklist',
    purpose: 'Section 7 checklist items.',
    example: ['Health check returns 200', 'No error spike in logs'],
  },
  {
    method: 'setMonitoring',
    purpose: 'Section 8 What/Tool/Alert Recipient table.',
    example: [{ what: 'Uptime', tool: 'Pingdom', alertRecipient: 'On-call engineer' }],
  },
];

class DeploymentGuideSDK {
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
  setOverview(t) {
    this.data.overview = t;
    return this;
  }
  setPrerequisites(rows) {
    this.data.prerequisites = rows;
    return this;
  }
  setEnvironments(rows) {
    this.data.environments = rows;
    return this;
  }
  setEnvExample(lines) {
    this.data.envExample = lines;
    return this;
  }
  setDeploymentSteps(lines) {
    this.data.deploymentSteps = lines;
    return this;
  }
  setConfiguration(rows) {
    this.data.configuration = rows;
    return this;
  }
  setRollbackSteps(lines) {
    this.data.rollbackSteps = lines;
    return this;
  }
  setRollbackNote(t) {
    this.data.rollbackNote = t;
    return this;
  }
  setPostDeploymentChecklist(items) {
    this.data.postDeploymentChecklist = items;
    return this;
  }
  setMonitoring(rows) {
    this.data.monitoring = rows;
    return this;
  }

  _buildPartTitlePage() {
    return [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [headingRun('Deployment Guide')],
      }),
    ];
  }

  _buildContent() {
    const children = [];
    children.push(h1('Deployment Guide'));
    const meta = Object.assign({}, DEFAULT_METADATA, this.data.metadata || {});
    children.push(this._metadataTable(meta));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));

    children.push(h2('1. Overview'));
    children.push(bodyPara(this.data.overview || DEFAULT_OVERVIEW));

    children.push(h2('2. Prerequisites'));
    children.push(
      buildZebraTable(
        [
          { key: 'requirement', header: 'Requirement', weight: 1.2 },
          { key: 'details', header: 'Details', weight: 2.5 },
        ],
        this.data.prerequisites,
        DEFAULT_PREREQUISITES,
      ),
    );

    children.push(h2('3. Environment Setup'));
    children.push(
      buildZebraTable(
        [
          { key: 'environment', header: 'Environment', weight: 1 },
          { key: 'purpose', header: 'Purpose', weight: 1.3 },
          { key: 'branch', header: 'Branch', weight: 1 },
        ],
        this.data.environments,
        DEFAULT_ENVIRONMENTS,
      ),
    );
    children.push(
      ...codeBlock(
        this.data.envExample && this.data.envExample.length
          ? this.data.envExample
          : DEFAULT_ENV_EXAMPLE,
      ),
    );

    children.push(h2('4. Deployment Steps'));
    children.push(
      ...codeBlock(
        this.data.deploymentSteps && this.data.deploymentSteps.length
          ? this.data.deploymentSteps
          : DEFAULT_DEPLOYMENT_STEPS,
      ),
    );

    children.push(h2('5. Configuration'));
    children.push(
      buildZebraTable(
        [
          { key: 'variable', header: 'Variable', weight: 1 },
          { key: 'description', header: 'Description', weight: 2 },
          { key: 'required', header: 'Required', weight: 0.8 },
        ],
        this.data.configuration,
        DEFAULT_CONFIGURATION,
      ),
    );

    children.push(h2('6. Rollback Procedure'));
    children.push(
      ...codeBlock(
        this.data.rollbackSteps && this.data.rollbackSteps.length
          ? this.data.rollbackSteps
          : DEFAULT_ROLLBACK_STEPS,
      ),
    );
    children.push(bodyPara(this.data.rollbackNote || DEFAULT_ROLLBACK_NOTE));

    children.push(h2('7. Post-Deployment Verification'));
    (this.data.postDeploymentChecklist && this.data.postDeploymentChecklist.length
      ? this.data.postDeploymentChecklist
      : DEFAULT_POST_DEPLOY_CHECKLIST
    ).forEach((c) => children.push(checklistItem(c)));

    children.push(h2('8. Monitoring & Alerting'));
    children.push(
      buildZebraTable(
        [
          { key: 'what', header: 'What', weight: 1 },
          { key: 'tool', header: 'Tool', weight: 1.4 },
          { key: 'alertRecipient', header: 'Alert Recipient', weight: 1.4 },
        ],
        this.data.monitoring,
        DEFAULT_MONITORING,
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
            run('Deployment Guide', {
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

module.exports = DeploymentGuideSDK;
module.exports.SECTION_GUIDE = SECTION_GUIDE;
module.exports.buildZebraTable = buildZebraTable;
module.exports.codeBlock = codeBlock;
