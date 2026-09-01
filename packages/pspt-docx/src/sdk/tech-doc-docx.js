'use strict';
/**
 * =============================================================================
 *  TECHNICAL DOCUMENTATION (Part 6 of the Product Documentation Master)
 * =============================================================================
 * Regenerates "Part 6. Technical Documentation" exactly. This Part is
 * unusual: it bundles THREE mini-documents, each with its OWN title-only
 * page and its own running header suffix:
 *   6a. Technical Documentation        — System Architecture + Security
 *       Requirements + Revision History (header: "Technical Documentation")
 *   6b. Data Model (ERD)               (header: "Technical Documentation [ERD]")
 *   6c. API Specification              (header: "Technical Documentation [API]")
 * The "[ERD]"/"[API]" suffix in the header is bold+italic navy, distinct
 * from the plain bold "Technical Documentation" label before it.
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
  checklistItem,
  codeBlock,
  guidanceNote,
  h1,
  h2,
  h3,
  headingRun,
  makeBullet,
  outerBorderSet,
  run,
} = require('./pspt-core');

const BULLET_NUMBERING_REF = 'techdoc-bullet-list';

/** Single-level bullet paragraph bound to this Part's numbering reference. */
const bullet = makeBullet(BULLET_NUMBERING_REF);

function boldPara(text) {
  return new Paragraph({
    spacing: { before: 120, after: 60, line: 276, lineRule: 'auto' },
    children: [run(text, { bold: true })],
  });
}

function metadataTable(meta) {
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
            children: [new Paragraph({ children: [run(r.value == null ? '' : String(r.value))] })],
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

// =============================================================================
// DEFAULTS — verbatim from the master template.
// =============================================================================

const DEFAULT_METADATA = {
  writer: '',
  status: 'Draft',
  version: 'V1 (Phase )',
  lastUpdate: 'Aug 20, 2026',
};

const DEFAULT_ARCH_OVERVIEW =
  '[Summarise the overall architecture style (monolith, microservices, client-server) and the reasoning behind it.]';
const DEFAULT_ARCH_DIAGRAM_NOTE =
  '[Insert a high-level box-and-arrow diagram showing major components and the direction of data flow between them.]';
const DEFAULT_COMPONENTS = [
  {
    component: '[Frontend]',
    responsibility: '[UI rendering, client-side routing, form validation]',
  },
  { component: '[Backend/API]', responsibility: '[Business logic, authentication, data access]' },
  { component: '[Database]', responsibility: '[Persistent data storage]' },
];
const DEFAULT_TECH_STACK = [
  {
    layer: 'Frontend',
    technology: '[e.g. React, TypeScript]',
    version: '[ver]',
    notes: '[rationale]',
  },
  {
    layer: 'Backend',
    technology: '[e.g. Node.js, Go, Java]',
    version: '[ver]',
    notes: '[rationale]',
  },
  {
    layer: 'Database',
    technology: '[e.g. PostgreSQL, DynamoDB]',
    version: '[ver]',
    notes: '[rationale]',
  },
  {
    layer: 'Infrastructure / Hosting',
    technology: '[e.g. AWS, GCP, Azure]',
    version: '—',
    notes: '[rationale]',
  },
  {
    layer: 'CI/CD',
    technology: '[e.g. GitHub Actions, Jenkins]',
    version: '—',
    notes: '[rationale]',
  },
  {
    layer: 'Monitoring & Logging',
    technology: '[e.g. Datadog, Grafana]',
    version: '—',
    notes: '[rationale]',
  },
];
const DEFAULT_CODEBASE_TREE = [
  'project-root/',
  '├── app/              # Frontend application',
  '├── api/              # Backend service',
  '└── docs/             # Project documentation',
];
const DEFAULT_NAMING_CONVENTIONS = [
  {
    item: 'Frontend component file',
    convention: '[e.g. PascalCase]',
    example: '[e.g. LeadCard.jsx]',
  },
  {
    item: 'Backend route file',
    convention: '[e.g. kebab-case, matches resource]',
    example: '[e.g. leads.routes.js]',
  },
];
const DEFAULT_CONFIG_NOTE =
  'Configuration: [.env holds local secrets, not committed; .env.example documents required variables with no real values].';
const DEFAULT_DATA_FLOW =
  '[e.g. When a user submits [action]: Frontend sends [METHOD] [endpoint] → API validates auth + request body → business logic executes → database write → response returned → UI renders success state.]';
const DEFAULT_INTEGRATION_POINTS = [
  {
    system: '[Service name]',
    purpose: '[Purpose/failure impact]',
    protocol: '[Protocol]',
    criticality: '[Critical / Non-critical]',
  },
];
const DEFAULT_SCALABILITY = [
  'Expected load: [current] concurrent users initially; re-evaluate above [threshold]',
  '[Connection pooling / caching notes]',
  '[Known bottlenecks and mitigation plan]',
];
const DEFAULT_DEPLOYMENT_TOPOLOGY =
  '[Insert deployment topology diagram — servers, containers, cloud regions, reverse proxy, etc.]';
const DEFAULT_SECURITY_OVERVIEW =
  "[Scope and purpose; compliance standards targeted (SOC2, ISO 27001, etc.) if any, or state 'None targeted in this phase'.]";
const DEFAULT_AUTH_REQUIREMENTS = [
  {
    id: 'SEC-AUTH-01',
    requirement: 'Passwords shall be hashed using bcrypt with a minimum cost factor of 10.',
  },
  {
    id: 'SEC-AUTH-02',
    requirement: 'Access tokens shall expire after [1 hour]; refresh tokens after [7 days].',
  },
  {
    id: 'SEC-AUTH-03',
    requirement: 'Role-based access control shall restrict admin-only endpoints from other roles.',
  },
];
const DEFAULT_DATA_PROTECTION = [
  { id: 'SEC-DATA-01', requirement: 'All traffic shall be served over HTTPS/TLS 1.2+.' },
  {
    id: 'SEC-DATA-02',
    requirement:
      'Credentials shall be stored in environment variables, never committed to source control.',
  },
  {
    id: 'SEC-DATA-03',
    requirement: 'Client PII shall not be logged in plaintext in application logs.',
  },
];
const DEFAULT_INPUT_VALIDATION = [
  {
    id: 'SEC-INPUT-01',
    requirement:
      'All API inputs shall be validated against a schema before reaching business logic.',
  },
  {
    id: 'SEC-INPUT-02',
    requirement:
      'All database queries shall use parameterised queries; string concatenation into SQL is prohibited.',
  },
];
const DEFAULT_INPUT_VALIDATION_CODE = [
  '// Good — parameterised',
  "db.query('SELECT * FROM users WHERE email = ?', [email]);",
];
const DEFAULT_SESSION_MANAGEMENT = [
  'Tokens are stored in memory / httpOnly cookies, not localStorage, to reduce XSS token theft.',
  'Logout invalidates the refresh token server-side.',
];
const DEFAULT_AUDIT_LOGGING = [
  { event: 'Login attempt (success/failure)', fields: 'user id/email, timestamp, IP, result' },
  { event: 'Data export', fields: 'user id, exported entity, timestamp' },
  { event: 'Permission-denied access attempt', fields: 'user id, endpoint, timestamp' },
];
const DEFAULT_COMPLIANCE = [
  'Data retention: [policy — e.g. indefinite while account active; deleted within 30 days of closure request]',
  "[Regulatory scope — e.g. GDPR, HIPAA, PCI-DSS — status/certification date, or 'None in this phase']",
];
const DEFAULT_VULN_MANAGEMENT = [
  { activity: 'Dependency vulnerability scan', frequency: 'Every release' },
  { activity: 'Manual penetration test', frequency: 'Annually, or before a major release' },
];
const DEFAULT_SECURITY_CHECKLIST = [
  'No secrets committed to the repository',
  'All endpoints enforce authentication/authorisation as specified',
  'Dependencies scanned with no unresolved high/critical vulnerabilities',
  'HTTPS enforced in production',
];
const DEFAULT_REVISION_HISTORY = [
  { version: 'v0.1', date: '[date]', author: '[name]', changes: 'Initial draft' },
];

const DEFAULT_ERD_OVERVIEW =
  '[What data domain does this ERD cover, and target database engine, e.g. PostgreSQL 16 / MySQL 8?]';
const DEFAULT_ENTITY_LIST = [{ entity: '`[table_name]`', description: '[One-line purpose]' }];
const DEFAULT_ERD_DIAGRAM_NOTE =
  "[Insert ERD diagram — crow's-foot or UML notation, showing entities, key attributes, and relationship cardinality.]";
const DEFAULT_ENTITY_DETAILS = [
  {
    tableName: '`[table_name]`',
    columns: [
      {
        column: '`id`',
        type: 'BIGINT',
        constraints: 'PK, auto-increment',
        description: 'Unique identifier',
      },
      {
        column: '`[fk]_id`',
        type: 'BIGINT',
        constraints: 'FK → `[table].id`, NOT NULL',
        description: '[relationship]',
      },
      {
        column: '`created_at`',
        type: 'DATETIME',
        constraints: 'NOT NULL',
        description: 'Record creation timestamp',
      },
    ],
  },
];
const DEFAULT_RELATIONSHIPS = [
  {
    from: '`[table_a]`',
    to: '`[table_b]`',
    cardinality: '1 : 0..*',
    rule: '[Business rule: this represents]',
  },
];
const DEFAULT_NORMALISATION_NOTES =
  '[Confirm the schema follows normal form, or explain any deliberate denormalisation and why — heads off review questions.]';
const DEFAULT_INDEXING_STRATEGY = [
  { table: '`[table]`', index: '`([columns])`', reason: '[Query pattern this serves]' },
];

const DEFAULT_API_OVERVIEW = '[API purpose, style (REST/GraphQL), and intended consumers.]';
const DEFAULT_API_BASE_URLS = [
  'Production: https://api.example.com/v1',
  'Staging:    https://staging-api.example.com/v1',
];
const DEFAULT_API_VERSIONING_NOTE =
  'Versioning is via URL path ([/v1], [/v2]); breaking changes require a new version.';
const DEFAULT_API_AUTH_NOTE = '[Auth scheme — JWT/API key/OAuth — and how to obtain a token.]';
const DEFAULT_API_AUTH_HEADER = ['Authorisation: Bearer <token>'];
const DEFAULT_API_ROLES_NOTE =
  'Role-based access: [roles]. See per-endpoint Access notes below, and Section 4.2 for the underlying auth requirements.';
const DEFAULT_API_RESPONSE_FORMAT = [
  '{',
  '  "success": true,',
  '  "data": { "...": "..." },',
  '  "error": null',
  '}',
];
const DEFAULT_API_ERROR_CODES = [
  { status: '200', meaning: 'Success' },
  { status: '400', meaning: 'Validation error (see error.fields)' },
  { status: '401', meaning: 'Missing/invalid token' },
  { status: '403', meaning: 'Authenticated but not authorised for this action' },
  { status: '404', meaning: 'Resource not found' },
];
const DEFAULT_ENDPOINTS = [
  {
    method: 'POST',
    path: '/v1/[resource]',
    description: 'Description: [Create a new resource]. Access: [roles].',
    requestBody: ['// Request body', '{', '  "field": "string, required"', '}'],
    responseBody: [
      '// Response 201 Created',
      '{',
      '  "success": true,',
      '  "data": { "id": 101 }',
      '}',
    ],
    errors: 'Errors: 400 if [field] is missing/invalid; 401 if unauthenticated.',
  },
];
const DEFAULT_PAGINATION_NOTE =
  'List endpoints accept ?page=1&limit=20&sort=-created_at&filter[field]=value. Default limit 20, max 100.';
const DEFAULT_RATE_LIMITING = [
  { limit: '[N] requests', window: 'per minute per token', header: '`X-RateLimit-Remaining`' },
];
const DEFAULT_API_CHANGELOG = [{ version: 'v1.1', date: '[date]', change: '[Change description]' }];

const SECTION_GUIDE = [
  {
    method: 'setHeaderFooterLabels',
    purpose: 'Overrides running header product name label.',
    example: { productNameLabel: 'Acme Widget' },
  },
  {
    method: 'setMetadata',
    purpose: 'Writer/Status/Version/Last Update metadata table (applies to 6a).',
    example: { writer: 'Jane Doe', status: 'Draft', version: 'V1', lastUpdate: 'Aug 31, 2026' },
  },
  {
    method: 'setArchitectureOverview',
    purpose: '6a §1.1 architecture style summary.',
    example: 'Microservices behind an API gateway...',
  },
  {
    method: 'setComponents',
    purpose: '6a §1.3 Component/Responsibility table.',
    example: [{ component: 'Frontend', responsibility: 'UI rendering' }],
  },
  {
    method: 'setTechStack',
    purpose: '6a §1.4 Layer/Technology/Version/Notes table.',
    example: [{ layer: 'Frontend', technology: 'React', version: '18', notes: 'Team familiarity' }],
  },
  {
    method: 'setCodebaseTree',
    purpose: '6a §1.5 repo tree lines (rendered as a shaded code block).',
    example: ['project-root/', '├── app/', '└── api/'],
  },
  {
    method: 'setNamingConventions',
    purpose: '6a §1.5 Item/Convention/Example table.',
    example: [{ item: 'Component file', convention: 'PascalCase', example: 'LeadCard.jsx' }],
  },
  {
    method: 'setDataFlow',
    purpose: '6a §1.6 data flow walkthrough.',
    example: 'When a user submits an expense: Frontend POSTs /expenses...',
  },
  {
    method: 'setIntegrationPoints',
    purpose: '6a §1.7 External System/Purpose/Protocol/Criticality table.',
    example: [
      { system: 'Stripe', purpose: 'Payments', protocol: 'REST/HTTPS', criticality: 'Critical' },
    ],
  },
  {
    method: 'setScalabilityNotes',
    purpose: '6a §1.8 bulleted scalability notes.',
    example: ['Expected load: 500 concurrent users'],
  },
  {
    method: 'setDeploymentTopologyNote',
    purpose: '6a §1.9 deployment topology note.',
    example: 'Deployed on AWS ECS across 2 availability zones.',
  },
  {
    method: 'setSecurityOverview',
    purpose: '6a §2.1 security scope/compliance.',
    example: 'Targets SOC2 Type II by end of year.',
  },
  {
    method: 'setAuthRequirements',
    purpose: '6a §2.2 ID/Requirement table.',
    example: [{ id: 'SEC-AUTH-01', requirement: 'Passwords hashed with bcrypt.' }],
  },
  {
    method: 'setDataProtection',
    purpose: '6a §2.3 ID/Requirement table.',
    example: [{ id: 'SEC-DATA-01', requirement: 'All traffic over TLS 1.2+.' }],
  },
  {
    method: 'setInputValidation',
    purpose: '6a §2.4 ID/Requirement table.',
    example: [{ id: 'SEC-INPUT-01', requirement: 'Schema validation on all inputs.' }],
  },
  {
    method: 'setInputValidationCode',
    purpose: '6a §2.4 code sample lines.',
    example: ["db.query('SELECT * FROM users WHERE id = ?', [id]);"],
  },
  {
    method: 'setSessionManagement',
    purpose: '6a §2.5 bulleted notes.',
    example: ['Tokens stored in httpOnly cookies'],
  },
  {
    method: 'setAuditLogging',
    purpose: '6a §2.6 Event/Logged Fields table.',
    example: [{ event: 'Login', fields: 'user id, timestamp' }],
  },
  {
    method: 'setComplianceRequirements',
    purpose: '6a §2.7 bulleted notes.',
    example: ['GDPR compliant as of Q1 2026'],
  },
  {
    method: 'setVulnerabilityManagement',
    purpose: '6a §2.8 Activity/Frequency table.',
    example: [{ activity: 'Pen test', frequency: 'Annually' }],
  },
  {
    method: 'setSecurityChecklist',
    purpose: '6a §2.9 checklist items.',
    example: ['No secrets committed', 'HTTPS enforced'],
  },
  {
    method: 'setRevisionHistory',
    purpose: '6a §3 Version/Date/Author/Changes table.',
    example: [
      { version: 'v0.1', date: 'Aug 31, 2026', author: 'Jane Doe', changes: 'Initial draft' },
    ],
  },
  {
    method: 'setErdOverview',
    purpose: 'ERD §1 data domain/engine.',
    example: 'Covers billing domain, PostgreSQL 16.',
  },
  {
    method: 'setEntityList',
    purpose: 'ERD §2 Entity/Description table.',
    example: [{ entity: '`users`', description: 'Registered accounts' }],
  },
  {
    method: 'setErdDiagramNote',
    purpose: 'ERD §3 diagram placeholder note.',
    example: '[See erd.png]',
  },
  {
    method: 'setEntityDetails',
    purpose: 'ERD §4 per-entity column tables.',
    example: [
      {
        tableName: '`users`',
        columns: [{ column: '`id`', type: 'BIGINT', constraints: 'PK', description: 'User ID' }],
      },
    ],
  },
  {
    method: 'setRelationships',
    purpose: 'ERD §5 From/To/Cardinality/Rule table.',
    example: [
      { from: '`users`', to: '`orders`', cardinality: '1:0..*', rule: 'A user has many orders' },
    ],
  },
  {
    method: 'setNormalisationNotes',
    purpose: 'ERD §6 note.',
    example: 'Schema follows 3NF throughout.',
  },
  {
    method: 'setIndexingStrategy',
    purpose: 'ERD §7 Table/Index/Reason table.',
    example: [{ table: '`orders`', index: '`(user_id)`', reason: 'Lookup by user' }],
  },
  {
    method: 'setApiOverview',
    purpose: 'API §1 purpose/style/consumers.',
    example: 'REST API consumed by the web and mobile apps.',
  },
  {
    method: 'setApiBaseUrls',
    purpose: 'API §2 base URL lines.',
    example: ['Production: https://api.acme.com/v1'],
  },
  {
    method: 'setApiVersioningNote',
    purpose: 'API §2 versioning note.',
    example: 'Versioned via URL path.',
  },
  {
    method: 'setApiAuthNote',
    purpose: 'API §3 auth scheme note.',
    example: 'JWT bearer tokens issued via OAuth2 client credentials.',
  },
  {
    method: 'setApiAuthHeader',
    purpose: 'API §3 auth header code line(s).',
    example: ['Authorization: Bearer <token>'],
  },
  {
    method: 'setApiRolesNote',
    purpose: 'API §3 roles note.',
    example: 'Roles: admin, member, viewer.',
  },
  {
    method: 'setApiResponseFormat',
    purpose: 'API §4 response format JSON lines.',
    example: ['{', '  "success": true', '}'],
  },
  {
    method: 'setApiErrorCodes',
    purpose: 'API §4 Status/Meaning table.',
    example: [{ status: '200', meaning: 'Success' }],
  },
  {
    method: 'setEndpoints',
    purpose: 'API §5 endpoint blocks.',
    example: [
      {
        method: 'GET',
        path: '/v1/users',
        description: 'List users.',
        requestBody: [],
        responseBody: ['{ "data": [] }'],
        errors: 'Errors: 401 if unauthenticated.',
      },
    ],
  },
  {
    method: 'setPaginationNote',
    purpose: 'API §6 pagination note.',
    example: 'Supports ?page and ?limit params.',
  },
  {
    method: 'setRateLimiting',
    purpose: 'API §7 Limit/Window/Header table.',
    example: [{ limit: '100 requests', window: 'per minute', header: 'X-RateLimit-Remaining' }],
  },
  {
    method: 'setApiChangelog',
    purpose: 'API §8 Version/Date/Change table.',
    example: [{ version: 'v1.1', date: 'Aug 31, 2026', change: 'Added pagination' }],
  },
];

class TechnicalDocumentationSDK {
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
  setArchitectureOverview(t) {
    this.data.architectureOverview = t;
    return this;
  }
  setComponents(rows) {
    this.data.components = rows;
    return this;
  }
  setTechStack(rows) {
    this.data.techStack = rows;
    return this;
  }
  setCodebaseTree(lines) {
    this.data.codebaseTree = lines;
    return this;
  }
  setNamingConventions(rows) {
    this.data.namingConventions = rows;
    return this;
  }
  setDataFlow(t) {
    this.data.dataFlow = t;
    return this;
  }
  setIntegrationPoints(rows) {
    this.data.integrationPoints = rows;
    return this;
  }
  setScalabilityNotes(items) {
    this.data.scalabilityNotes = items;
    return this;
  }
  setDeploymentTopologyNote(t) {
    this.data.deploymentTopologyNote = t;
    return this;
  }
  setSecurityOverview(t) {
    this.data.securityOverview = t;
    return this;
  }
  setAuthRequirements(rows) {
    this.data.authRequirements = rows;
    return this;
  }
  setDataProtection(rows) {
    this.data.dataProtection = rows;
    return this;
  }
  setInputValidation(rows) {
    this.data.inputValidation = rows;
    return this;
  }
  setInputValidationCode(lines) {
    this.data.inputValidationCode = lines;
    return this;
  }
  setSessionManagement(items) {
    this.data.sessionManagement = items;
    return this;
  }
  setAuditLogging(rows) {
    this.data.auditLogging = rows;
    return this;
  }
  setComplianceRequirements(items) {
    this.data.complianceRequirements = items;
    return this;
  }
  setVulnerabilityManagement(rows) {
    this.data.vulnerabilityManagement = rows;
    return this;
  }
  setSecurityChecklist(items) {
    this.data.securityChecklist = items;
    return this;
  }
  setRevisionHistory(rows) {
    this.data.revisionHistory = rows;
    return this;
  }
  setErdOverview(t) {
    this.data.erdOverview = t;
    return this;
  }
  setEntityList(rows) {
    this.data.entityList = rows;
    return this;
  }
  setErdDiagramNote(t) {
    this.data.erdDiagramNote = t;
    return this;
  }
  setEntityDetails(items) {
    this.data.entityDetails = items;
    return this;
  }
  setRelationships(rows) {
    this.data.relationships = rows;
    return this;
  }
  setNormalisationNotes(t) {
    this.data.normalisationNotes = t;
    return this;
  }
  setIndexingStrategy(rows) {
    this.data.indexingStrategy = rows;
    return this;
  }
  setApiOverview(t) {
    this.data.apiOverview = t;
    return this;
  }
  setApiBaseUrls(lines) {
    this.data.apiBaseUrls = lines;
    return this;
  }
  setApiVersioningNote(t) {
    this.data.apiVersioningNote = t;
    return this;
  }
  setApiAuthNote(t) {
    this.data.apiAuthNote = t;
    return this;
  }
  setApiAuthHeader(lines) {
    this.data.apiAuthHeader = lines;
    return this;
  }
  setApiRolesNote(t) {
    this.data.apiRolesNote = t;
    return this;
  }
  setApiResponseFormat(lines) {
    this.data.apiResponseFormat = lines;
    return this;
  }
  setApiErrorCodes(rows) {
    this.data.apiErrorCodes = rows;
    return this;
  }
  setEndpoints(items) {
    this.data.endpoints = items;
    return this;
  }
  setPaginationNote(t) {
    this.data.paginationNote = t;
    return this;
  }
  setRateLimiting(rows) {
    this.data.rateLimiting = rows;
    return this;
  }
  setApiChangelog(rows) {
    this.data.apiChangelog = rows;
    return this;
  }

  _buildTechDocContent() {
    const children = [];
    children.push(h1('Technical Documentation'));
    const meta = Object.assign({}, DEFAULT_METADATA, this.data.metadata || {});
    children.push(metadataTable(meta));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));

    children.push(h2('1. System Architecture'));
    children.push(h3('1.1 Overview'));
    children.push(bodyPara(this.data.architectureOverview || DEFAULT_ARCH_OVERVIEW));
    children.push(h3('1.2 Architecture Diagram'));
    children.push(bodyPara(DEFAULT_ARCH_DIAGRAM_NOTE));
    children.push(h3('1.3 Components / Layers'));
    children.push(
      buildZebraTable(
        [
          { key: 'component', header: 'Component', weight: 1 },
          { key: 'responsibility', header: 'Responsibility', weight: 2.5 },
        ],
        this.data.components,
        DEFAULT_COMPONENTS,
      ),
    );
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('1.4 Technology Stack'));
    children.push(
      buildZebraTable(
        [
          { key: 'layer', header: 'Layer', weight: 1 },
          { key: 'technology', header: 'Technology', weight: 1.4 },
          { key: 'version', header: 'Version', weight: 0.7 },
          { key: 'notes', header: 'Notes', weight: 1.4 },
        ],
        this.data.techStack,
        DEFAULT_TECH_STACK,
      ),
    );
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('1.5 Codebase Structure'));
    children.push(
      guidanceNote(
        'Repository layout style (monorepo vs separate repos), root-level tree, and naming conventions.',
      ),
    );
    children.push(
      ...codeBlock(
        this.data.codebaseTree && this.data.codebaseTree.length
          ? this.data.codebaseTree
          : DEFAULT_CODEBASE_TREE,
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'item', header: 'Item', weight: 1.2 },
          { key: 'convention', header: 'Convention', weight: 1.3 },
          { key: 'example', header: 'Example', weight: 1.3 },
        ],
        this.data.namingConventions,
        DEFAULT_NAMING_CONVENTIONS,
      ),
    );
    children.push(bodyPara(DEFAULT_CONFIG_NOTE, { italics: true }));
    children.push(h3('1.6 Data Flow'));
    children.push(
      guidanceNote(
        'Walk through one or two representative request flows end-to-end, beyond the static component diagram above.',
      ),
    );
    children.push(bodyPara(this.data.dataFlow || DEFAULT_DATA_FLOW));
    children.push(h3('1.7 Integration Points'));
    children.push(
      buildZebraTable(
        [
          { key: 'system', header: 'External System', weight: 1.2 },
          { key: 'purpose', header: 'Purpose', weight: 1.4 },
          { key: 'protocol', header: 'Protocol', weight: 0.9 },
          { key: 'criticality', header: 'Criticality', weight: 1 },
        ],
        this.data.integrationPoints,
        DEFAULT_INTEGRATION_POINTS,
      ),
    );
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('1.8 Scalability & Performance Considerations'));
    (this.data.scalabilityNotes && this.data.scalabilityNotes.length
      ? this.data.scalabilityNotes
      : DEFAULT_SCALABILITY
    ).forEach((s) => children.push(bullet(s)));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('1.9 Deployment Topology'));
    children.push(
      guidanceNote(
        'Where components physically/logically run — complements the step-by-step Deployment Guide (Part 9). Data privacy and security controls are not repeated here — see Section 4 below.',
      ),
    );
    children.push(bodyPara(this.data.deploymentTopologyNote || DEFAULT_DEPLOYMENT_TOPOLOGY));

    children.push(h2('2. Security Requirements'));
    children.push(
      guidanceNote(
        'Specifies the security posture the system must meet — treat every row as a testable requirement, not an aspirational goal. Each should be verifiable during code review, testing, or a security audit.',
      ),
    );
    children.push(h3('2.1 Overview'));
    children.push(bodyPara(this.data.securityOverview || DEFAULT_SECURITY_OVERVIEW));
    children.push(h3('2.2 Authentication & Authorization Requirements'));
    children.push(
      buildZebraTable(
        [
          { key: 'id', header: 'ID', weight: 1 },
          { key: 'requirement', header: 'Requirement', weight: 3 },
        ],
        this.data.authRequirements,
        DEFAULT_AUTH_REQUIREMENTS,
      ),
    );
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('2.3 Data Protection'));
    children.push(
      buildZebraTable(
        [
          { key: 'id', header: 'ID', weight: 1 },
          { key: 'requirement', header: 'Requirement', weight: 3 },
        ],
        this.data.dataProtection,
        DEFAULT_DATA_PROTECTION,
      ),
    );
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('2.4 Input Validation & Sanitisation'));
    children.push(
      buildZebraTable(
        [
          { key: 'id', header: 'ID', weight: 1 },
          { key: 'requirement', header: 'Requirement', weight: 3 },
        ],
        this.data.inputValidation,
        DEFAULT_INPUT_VALIDATION,
      ),
    );
    children.push(
      ...codeBlock(
        this.data.inputValidationCode && this.data.inputValidationCode.length
          ? this.data.inputValidationCode
          : DEFAULT_INPUT_VALIDATION_CODE,
      ),
    );
    children.push(h3('2.5 Session Management'));
    (this.data.sessionManagement && this.data.sessionManagement.length
      ? this.data.sessionManagement
      : DEFAULT_SESSION_MANAGEMENT
    ).forEach((s) => children.push(bullet(s)));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('2.6 Audit Logging'));
    children.push(
      buildZebraTable(
        [
          { key: 'event', header: 'Event', weight: 1.4 },
          { key: 'fields', header: 'Logged Fields', weight: 2 },
        ],
        this.data.auditLogging,
        DEFAULT_AUDIT_LOGGING,
      ),
    );
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('2.7 Compliance Requirements'));
    (this.data.complianceRequirements && this.data.complianceRequirements.length
      ? this.data.complianceRequirements
      : DEFAULT_COMPLIANCE
    ).forEach((c) => children.push(bullet(c)));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('2.8 Vulnerability Management'));
    children.push(
      buildZebraTable(
        [
          { key: 'activity', header: 'Activity', weight: 1.5 },
          { key: 'frequency', header: 'Frequency', weight: 1.5 },
        ],
        this.data.vulnerabilityManagement,
        DEFAULT_VULN_MANAGEMENT,
      ),
    );
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('2.9 Security Checklist'));
    (this.data.securityChecklist && this.data.securityChecklist.length
      ? this.data.securityChecklist
      : DEFAULT_SECURITY_CHECKLIST
    ).forEach((c) => children.push(checklistItem(c)));

    children.push(h2('3. Revision History'));
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

  _buildErdContent() {
    const children = [];
    children.push(h2('Data Model (ERD)'));
    children.push(
      guidanceNote(
        'Documents the actual source of truth for what tables exist, what they contain, and how they relate. Keep this in sync with the real database schema — a stale ERD actively misleads.',
      ),
    );

    children.push(h3('1. Overview'));
    children.push(bodyPara(this.data.erdOverview || DEFAULT_ERD_OVERVIEW));

    children.push(h3('2.  Entity List'));
    children.push(
      buildZebraTable(
        [
          { key: 'entity', header: 'Entity', weight: 1 },
          { key: 'description', header: 'Description', weight: 2.5 },
        ],
        this.data.entityList,
        DEFAULT_ENTITY_LIST,
      ),
    );

    children.push(h3('3.  ERD Diagram'));
    children.push(
      guidanceNote(
        'The primary artefact of this section — everything else here supports and explains it.',
      ),
    );
    children.push(bodyPara(this.data.erdDiagramNote || DEFAULT_ERD_DIAGRAM_NOTE));

    children.push(h3('4. Entity Detail'));
    children.push(
      guidanceNote(
        'Repeat per entity — precise enough that a developer could write the CREATE TABLE statement directly from it.',
      ),
    );
    const entities =
      this.data.entityDetails && this.data.entityDetails.length
        ? this.data.entityDetails
        : DEFAULT_ENTITY_DETAILS;
    entities.forEach((entity) => {
      children.push(boldPara(entity.tableName));
      children.push(
        buildZebraTable(
          [
            { key: 'column', header: 'Column', weight: 1 },
            { key: 'type', header: 'Type', weight: 0.8 },
            { key: 'constraints', header: 'Constraints', weight: 1.4 },
            { key: 'description', header: 'Description', weight: 1.4 },
          ],
          entity.columns,
          entity.columns,
        ),
      );
      children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    });

    children.push(h3('5. Relationships'));
    children.push(
      buildZebraTable(
        [
          { key: 'from', header: 'From', weight: 1 },
          { key: 'to', header: 'To', weight: 1 },
          { key: 'cardinality', header: 'Cardinality', weight: 0.9 },
          { key: 'rule', header: 'Rule', weight: 1.6 },
        ],
        this.data.relationships,
        DEFAULT_RELATIONSHIPS,
      ),
    );

    children.push(h3('6. Normalisation Notes'));
    children.push(bodyPara(this.data.normalisationNotes || DEFAULT_NORMALISATION_NOTES));

    children.push(h3('7. Indexing Strategy'));
    children.push(
      buildZebraTable(
        [
          { key: 'table', header: 'Table', weight: 1 },
          { key: 'index', header: 'Index', weight: 1.2 },
          { key: 'reason', header: 'Reason', weight: 1.6 },
        ],
        this.data.indexingStrategy,
        DEFAULT_INDEXING_STRATEGY,
      ),
    );

    return children;
  }

  _buildApiContent() {
    const children = [];
    children.push(h2('API Specification'));
    children.push(
      guidanceNote(
        'The contract between the backend and any consumer (frontend, mobile app, third-party integration). Ideally generated from, or kept in lockstep with, an OpenAPI/Swagger spec — this section is the human-readable companion.',
      ),
    );

    children.push(h3('1. Overview'));
    children.push(bodyPara(this.data.apiOverview || DEFAULT_API_OVERVIEW));

    children.push(h3('2. Base URL & Versioning'));
    children.push(
      ...codeBlock(
        this.data.apiBaseUrls && this.data.apiBaseUrls.length
          ? this.data.apiBaseUrls
          : DEFAULT_API_BASE_URLS,
      ),
    );
    children.push(bodyPara(this.data.apiVersioningNote || DEFAULT_API_VERSIONING_NOTE));

    children.push(h3('3. Authentication & Authorisation'));
    children.push(bodyPara(this.data.apiAuthNote || DEFAULT_API_AUTH_NOTE));
    children.push(
      ...codeBlock(
        this.data.apiAuthHeader && this.data.apiAuthHeader.length
          ? this.data.apiAuthHeader
          : DEFAULT_API_AUTH_HEADER,
      ),
    );
    children.push(bodyPara(this.data.apiRolesNote || DEFAULT_API_ROLES_NOTE));

    children.push(h3('4. Common Response Format & Error Codes'));
    children.push(
      ...codeBlock(
        this.data.apiResponseFormat && this.data.apiResponseFormat.length
          ? this.data.apiResponseFormat
          : DEFAULT_API_RESPONSE_FORMAT,
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'status', header: 'Status', weight: 1 },
          { key: 'meaning', header: 'Meaning', weight: 2.5 },
        ],
        this.data.apiErrorCodes,
        DEFAULT_API_ERROR_CODES,
      ),
    );

    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('5. Endpoints'));
    children.push(
      guidanceNote(
        'Repeat this block per endpoint, grouped by resource, mirroring the Data Model structure in Section 2 above.',
      ),
    );
    const endpoints =
      this.data.endpoints && this.data.endpoints.length ? this.data.endpoints : DEFAULT_ENDPOINTS;
    endpoints.forEach((ep) => {
      children.push(boldPara(`\`${ep.method} ${ep.path}\``));
      children.push(bodyPara(ep.description));
      if (ep.requestBody && ep.requestBody.length) children.push(...codeBlock(ep.requestBody));
      if (ep.responseBody && ep.responseBody.length) children.push(...codeBlock(ep.responseBody));
      children.push(bodyPara(ep.errors));
    });

    children.push(h3('6. Pagination, Filtering & Sorting'));
    children.push(bodyPara(this.data.paginationNote || DEFAULT_PAGINATION_NOTE));

    children.push(h3('7.  Rate Limiting'));
    children.push(
      buildZebraTable(
        [
          { key: 'limit', header: 'Limit', weight: 1.2 },
          { key: 'window', header: 'Window', weight: 1.2 },
          { key: 'header', header: 'Header', weight: 1.4 },
        ],
        this.data.rateLimiting,
        DEFAULT_RATE_LIMITING,
      ),
    );

    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('8. API Changelog'));
    children.push(
      guidanceNote(
        "The API's own version history (which endpoint changed, when) — distinct from the product-wide Changelog (Part 11), which tracks user-visible features rather than the wire contract.",
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'version', header: 'Version', weight: 0.8 },
          { key: 'date', header: 'Date', weight: 1 },
          { key: 'change', header: 'Change', weight: 2 },
        ],
        this.data.apiChangelog,
        DEFAULT_API_CHANGELOG,
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

  _chapterHeader(bracketSuffix) {
    const labels = this.data.headerFooterLabels || {};
    const runs = [
      run(`${labels.productNameLabel || 'Product name/logo'}\t`, {
        bold: true,
        size: SIZE.headerFooter,
        color: COLOR.headerFooterText,
      }),
      run('Technical Documentation ', {
        bold: true,
        size: SIZE.headerFooter,
        color: COLOR.headerFooterText,
      }),
    ];
    if (bracketSuffix) {
      runs.push(
        run(bracketSuffix, {
          bold: true,
          italics: true,
          size: SIZE.headerFooter,
          color: COLOR.headerFooterText,
        }),
      );
    }
    runs.push(run('|', { size: SIZE.headerFooter, color: COLOR.mutedText }));
    runs.push(
      run('Confidential', { bold: true, size: SIZE.headerFooter, color: COLOR.confidential }),
    );
    return new Header({
      children: [
        new Paragraph({
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR.tableBorder, space: 4 },
          },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: runs,
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

  _titlePageSection(headingText) {
    return {
      properties: {
        titlePage: true,
        page: { size: { width: PAGE.width, height: PAGE.height }, margin: PAGE.margin },
      },
      headers: { default: this._blankHeader(), first: this._blankHeader() },
      footers: { default: this._blankFooter(), first: this._blankFooter() },
      children: [
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [headingRun(headingText)] }),
      ],
    };
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
        this._titlePageSection('Technical Documentation'),
        {
          properties: {
            page: { size: { width: PAGE.width, height: PAGE.height }, margin: PAGE.margin },
          },
          headers: { default: this._chapterHeader() },
          footers: { default: this._chapterFooter() },
          children: this._buildTechDocContent(),
        },
        this._titlePageSection('Data Model (ERD)'),
        {
          properties: {
            page: { size: { width: PAGE.width, height: PAGE.height }, margin: PAGE.margin },
          },
          headers: { default: this._chapterHeader('[ERD]') },
          footers: { default: this._chapterFooter() },
          children: this._buildErdContent(),
        },
        this._titlePageSection('API Specification'),
        {
          properties: {
            page: { size: { width: PAGE.width, height: PAGE.height }, margin: PAGE.margin },
          },
          headers: { default: this._chapterHeader('[API]') },
          footers: { default: this._chapterFooter() },
          children: this._buildApiContent(),
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

module.exports = TechnicalDocumentationSDK;
module.exports.SECTION_GUIDE = SECTION_GUIDE;
module.exports.buildZebraTable = buildZebraTable;
module.exports.codeBlock = codeBlock;
module.exports.BULLET_NUMBERING_REF = BULLET_NUMBERING_REF;
