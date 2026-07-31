'use strict';
/**
 * =============================================================================
 *  PRODUCT SPECIFICATION DOCUMENT SDK
 * =============================================================================
 *
 * Regenerates the "Product Specification" Word template with byte-for-byte
 * matched STYLE: Arial throughout, the same heading colors/sizes, the same
 * navy-header/white-text table style, the same US-Letter page size & margins,
 * and the same header/footer/cover-page layout as the source template.
 *
 * WHO THIS IS FOR
 * ----------------
 * This file is meant to be driven by another AI (or a human) that has the
 * actual product content. You do NOT need to touch styling, fonts, colors,
 * margins, or table borders — that is all baked in and matches the template
 * exactly. You only call the `set...()` method for whichever section(s) you
 * have content for.
 *
 * HOW TO USE (quick start)
 * -------------------------
 *   const ProductSpecSDK = require('pspt-docx');
 *   const doc = new ProductSpecSDK();
 *
 *   doc.setCoverPage({
 *     productName: 'Acme Widget',
 *     shortDescription: 'A widget that connects other widgets.',
 *     lastUpdated: '2026-07-22',
 *     status: 'Draft',
 *   });
 *   doc.setExecutiveSummary('Acme Widget lets teams ...');
 *   // ... call whichever other set*() methods you have content for ...
 *
 *   await doc.generate('/mnt/user-data/outputs/product-spec.docx');
 *
 * IMPORTANT FOR AI CALLERS
 * -------------------------
 * - Every `set...()` method has a JSDoc comment directly above it describing
 *   (a) what real-world information belongs in that section and (b) the
 *   exact shape of the object/array it expects. READ IT before calling.
 * - You do not have to call every method. Any section you skip is rendered
 *   with the same "[bracketed placeholder]" prompt text the original
 *   template uses, in italics, so a human reviewer can see what's still
 *   missing. This means calling `generate()` with zero setter calls
 *   reproduces the blank template exactly.
 * - Call `ProductSpecSDK.sectionGuide()` (a static method, no instance
 *   needed) to get a machine-readable JSON array of every section: its
 *   method name, its purpose, and an example payload. Use this if you want
 *   to enumerate sections programmatically instead of reading the source.
 * - Tables: pass an array of plain objects. Keys are given in each method's
 *   JSDoc. Omit a method entirely (don't call it) to get the placeholder row.
 * - `setEnableHardwareSection(true)` turns on the optional 2.4 Hardware
 *   Specification block and the Hardware Validation table in Quality
 *   Control. Leave it off (default) for pure-software products.
 * =============================================================================
 */

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, VerticalAlign, HeightRule,
  Header, Footer, PageNumber, TabStopType, TabStopPosition, ImageRun, PageBreak,
} = require('docx');
const {
  DOCX_FONT: FONT,
  DOCX_COLOR: COLOR,
  DOCX_SIZE: SIZE,
  DOCX_PAGE: PAGE,
  DOCX_CONTENT_WIDTH: CONTENT_WIDTH,
  PLACEHOLDER_ROW_TEXT,
  distributeColumnWidths,
  resolveRows,
} = require('pspt-core');

// =============================================================================
// LOW-LEVEL BUILDERS (not exported — internal rendering helpers)
// =============================================================================

function run(text, opts = {}) {
  return new TextRun(Object.assign({ text: String(text), font: FONT, size: SIZE.body }, opts));
}

/** A normal Arial 11pt body paragraph. Pass {italics:true} for placeholder text. */
function bodyPara(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 160, line: 276, lineRule: 'auto' },
    children: [run(text, opts)],
  });
}

/** Splits a multi-line string into one bodyPara per non-empty line. */
function textBlock(content, fallback) {
  const hasContent = content != null && String(content).trim() !== '';
  const source = hasContent ? String(content) : String(fallback);
  const italics = !hasContent;
  const lines = source.split(/\r?\n+/).filter((l) => l.trim() !== '');
  if (lines.length === 0) lines.push(source);
  return lines.map((line) => bodyPara(line, italics ? { italics: true } : {}));
}

/** The recurring "Purpose: <description>" line that appears under every major heading. */
function purposeLine(text) {
  return new Paragraph({
    spacing: { after: 200, line: 276, lineRule: 'auto' },
    children: [
      run('Purpose: ', { bold: true, italics: true, color: COLOR.purposeLabel }),
      run(text, { italics: true, color: COLOR.purposeText }),
    ],
  });
}

/** heading levels map */
const H = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
};

function heading(text, level) {
  return new Paragraph({ heading: H[level], children: [run(text)] });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function borderSet() {
  const b = { style: BorderStyle.SINGLE, size: 4, color: COLOR.border };
  return { top: b, bottom: b, left: b, right: b, insideHorizontal: b, insideVertical: b };
}

function headerCell(text, width) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { fill: COLOR.tableHeaderBg, type: ShadingType.CLEAR, color: 'auto' },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [run(text, { bold: true, color: COLOR.tableHeaderText })],
    })],
  });
}

function bodyCell(text, width, opts = {}) {
  const italics = !!opts.placeholder;
  const value = text == null || text === '' ? (opts.placeholder ? PLACEHOLDER_ROW_TEXT : '') : String(text);
  const lines = value.split(/\r?\n+/);
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    margins: { top: 70, bottom: 70, left: 100, right: 100 },
    verticalAlign: VerticalAlign.CENTER,
    children: lines.map((line) => new Paragraph({ children: [run(line, { italics })] })),
  });
}

/**
 * Builds a table matching the template's table style exactly: navy header
 * row with centered white bold text, single 0.5pt black borders throughout,
 * columns proportioned across the full page content width.
 *
 * @param {Array<{key:string, header:string, weight?:number}>} columns
 * @param {Array<Object>|undefined} rows - real data rows, keyed by column.key
 * @param {Array<Object>} placeholderRows - fallback row(s) shown (in italics)
 *        when `rows` is empty/undefined, so the doc still shows the reader
 *        what belongs in each column.
 */
function buildTable(columns, rows, placeholderRows) {
  const { hasData, rows: useRows } = resolveRows(rows, placeholderRows);
  const widths = distributeColumnWidths(columns, CONTENT_WIDTH);

  const headerRow = new TableRow({
    tableHeader: true,
    children: columns.map((c, i) => headerCell(c.header, widths[i])),
  });
  const bodyRows = (useRows || []).map((r) => new TableRow({
    children: columns.map((c, i) => bodyCell(r ? r[c.key] : undefined, widths[i], { placeholder: !hasData })),
  }));
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: widths,
    borders: borderSet(),
    rows: [headerRow, ...bodyRows],
  });
}

/** A numbered "1. / a." style block, used by the Security section. */
function numberedBlock(items) {
  const paras = [];
  (items || []).forEach((item, i) => {
    paras.push(new Paragraph({
      spacing: { after: 60, line: 276, lineRule: 'auto' },
      indent: { left: 360, hanging: 360 },
      children: [run(`${i + 1}. ${item.title || item}`)],
    }));
    if (item && Array.isArray(item.children)) {
      item.children.forEach((sub, j) => {
        const letters = 'abcdefghijklmnopqrstuvwxyz';
        paras.push(new Paragraph({
          spacing: { after: 60, line: 276, lineRule: 'auto' },
          indent: { left: 1080, hanging: 360 },
          children: [run(`${letters[j] || j + 1}. ${sub}`)],
        }));
      });
    }
  });
  return paras;
}

function inferImageType(filePath) {
  const ext = path.extname(filePath).toLowerCase().replace('.', '');
  if (ext === 'jpg') return 'jpeg';
  return ext || 'png';
}

/** Embeds an image if a real path is given; otherwise draws a bordered
 *  placeholder box (matching the cover page's "IMAGE" placeholder graphic). */
function imageOrPlaceholder(imagePath, widthPx, heightPx, placeholderLabel) {
  if (imagePath && fs.existsSync(imagePath)) {
    const data = fs.readFileSync(imagePath);
    return new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new ImageRun({
        data,
        type: inferImageType(imagePath),
        transformation: { width: widthPx, height: heightPx },
      })],
    });
  }
  // Placeholder box: single-cell shaded/bordered table, same visual role as
  // the template's light-blue "IMAGE" placeholder shape.
  const w = Math.min(CONTENT_WIDTH, Math.round(widthPx * 15)); // rough px->dxa
  return new Table({
    width: { size: w, type: WidthType.DXA },
    alignment: AlignmentType.CENTER,
    columnWidths: [w],
    borders: borderSet(),
    rows: [new TableRow({
      height: { value: Math.max(1440, Math.round(heightPx * 15)), rule: HeightRule.ATLEAST },
      children: [new TableCell({
        width: { size: w, type: WidthType.DXA },
        shading: { fill: COLOR.imageBoxFill, type: ShadingType.CLEAR, color: 'auto' },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [run(placeholderLabel || '[Insert image]', { italics: true })],
        })],
      })],
    })],
  });
}

// =============================================================================
// SECTION GUIDE — machine-readable manifest of every section.
// An AI can call `ProductSpecSDK.sectionGuide()` to enumerate sections
// programmatically (method name, what it's for, example payload) instead of
// reading JSDoc source.
// =============================================================================

const SECTION_GUIDE = [
  { method: 'setCoverPage', purpose: 'Cover page: product name, one-line description, last-updated date, and status.', example: { productName: 'Acme Widget', shortDescription: 'A widget that connects other widgets.', lastUpdated: '2026-07-22', status: 'Draft', logoImagePath: null } },
  { method: 'setProductInfo', purpose: 'Small identity table at the top of Product Overview: product name, version, current status.', example: { productName: 'Acme Widget', version: 'v1.0', status: 'In development' } },
  { method: 'setExecutiveSummary', purpose: 'One-page narrative overview for any stakeholder: what the product is, why it exists, how success is measured.', example: 'Acme Widget is a ... It exists because ... Success is measured by ...' },
  { method: 'setProductRoadmap', purpose: 'High-level roadmap of major phases/releases.', example: [{ phase: 'v1.0', theme: 'Initial launch', timeframe: 'Q3 2026', status: 'In progress' }] },
  { method: 'setTargetMarket', purpose: 'Who this product is for, segmented.', example: [{ segment: 'End users', description: 'Individuals who need ...' }] },
  { method: 'setCustomerPersonas', purpose: 'Representative buyer/user personas.', example: [{ persona: 'Busy Manager', role: 'Team lead', goals: 'Save time', painPoints: 'Manual process', buyingBehavior: 'Free trial then upgrade' }] },
  { method: 'setUserJourney', purpose: 'A user-journey diagram/screenshot, with optional caption.', example: { imagePath: '/path/to/journey.png', caption: 'Guest booking flow' } },
  { method: 'setUseCases', purpose: 'Concrete use cases: actor, trigger, outcome.', example: [{ useCase: 'Sign up', actor: 'New user', trigger: 'Clicks "Sign up"', outcome: 'Account created' }] },
  { method: 'setCompetitiveAnalysis', purpose: 'Competitor comparison.', example: [{ competitor: 'Competitor A', strengths: '...', weaknesses: '...', pricing: '...', differentiator: '...' }] },
  { method: 'setPricingModel', purpose: 'Pricing approach and tier breakdown.', example: { modelDescription: 'Subscription tiers, monthly billing.', tiers: [{ tier: 'Basic', price: '$9/mo', includes: '...', targetSegment: '...' }] } },
  { method: 'setRevenueModel', purpose: 'How the product actually makes money.', example: 'Subscription SaaS revenue plus a reseller channel.' },
  { method: 'setFeatures', purpose: 'Feature list with priority (Must/Should/Could).', example: [{ name: 'Push notifications', description: 'Real-time alerts', priority: 'Must' }] },
  { method: 'setSystemArchitecture', purpose: 'Architecture narrative (monolith/microservices/serverless) plus optional diagram.', example: { description: 'Microservices behind an API gateway ...', diagramImagePath: null } },
  { method: 'setTechnologyStack', purpose: 'Tech choices per layer with justification.', example: [{ layer: 'Frontend', technology: 'React', justification: 'Team familiarity' }] },
  { method: 'setApis', purpose: 'Key API endpoints plus a link to full API docs.', example: { rows: [{ endpoint: '/api/v1/users', method: 'GET', description: 'List users', authRequired: 'Yes', role: 'Admin' }], docsLink: 'https://...' } },
  { method: 'setDatabaseDesign', purpose: 'Tables/collections with purpose and key fields, plus optional ER diagram.', example: { rows: [{ table: 'Users', purpose: 'User accounts', keyFields: 'id, email', url: '' }], diagramImagePath: null } },
  { method: 'setAuthentication', purpose: 'Auth method (OAuth2/SSO/SAML/JWT/API keys), MFA, session management.', example: 'OAuth2 with JWT access tokens; optional TOTP MFA.' },
  { method: 'setSecurity', purpose: 'Security controls as a nested numbered list (tool -> sub-features), plus free-text notes on encryption/pen-testing/dependency scanning.', example: { measures: [{ title: 'Cloudflare', children: ['WAF', 'Turnstile'] }], notes: 'Data encrypted at rest and in transit ...' } },
  { method: 'setMonitoring', purpose: 'Monitoring tools, key alerts/dashboards.', example: 'Datadog dashboards for latency, error rate, and saturation.' },
  { method: 'setLogging', purpose: 'Log levels, retention, centralized logging, PII handling.', example: '30-day retention in Google Cloud Logging; PII redacted at source.' },
  { method: 'setBackupStrategy', purpose: 'Backup frequency/retention/RTO/RPO.', example: 'Nightly backups, 30-day retention, RTO 4h, RPO 1h.' },
  { method: 'setEnableHardwareSection', purpose: 'Toggle the optional 2.4 Hardware Specification block (and Hardware Validation table). Leave off for pure software.', example: true },
  { method: 'setHardwareOverview', purpose: 'High-level device description, form factor, role in the product.', example: 'A pocket-sized BLE sensor puck.' },
  { method: 'setHardwareComponents', purpose: 'Major hardware components with vendor/backup vendor.', example: [{ component: 'MCU', function: 'Main compute', vendor: 'Nordic', altVendor: 'Espressif' }] },
  { method: 'setBillOfMaterials', purpose: 'Full BOM: part numbers, qty, unit cost, vendor, lead time.', example: [{ partNo: 'R-1001', description: '10k resistor', qty: '4', unitCost: '$0.01', vendor: 'Digikey', leadTime: '2 weeks' }] },
  { method: 'setMechanicalDesign', purpose: 'Enclosure material, dimensions, weight, tolerances, IP rating, plus optional CAD render.', example: { description: 'ABS enclosure, IP54.', diagramImagePath: null } },
  { method: 'setElectricalSpecification', purpose: 'Voltage, current draw, PCB notes, plus optional schematic.', example: { description: '5V input, 200mA typical draw.', diagramImagePath: null } },
  { method: 'setSensors', purpose: 'Onboard sensors with range/accuracy/purpose.', example: [{ sensor: 'Accelerometer', type: 'MEMS', range: '+-16g', accuracy: '1%', purpose: 'Motion detection' }] },
  { method: 'setConnectivity', purpose: 'Wireless/wired protocols, pairing, OTA support.', example: 'BLE 5.0 for pairing, Wi-Fi for OTA updates.' },
  { method: 'setPowerRequirements', purpose: 'Power source, battery life, charging method, consumption profile.', example: 'Li-ion 500mAh, ~7 days typical use, USB-C charging.' },
  { method: 'setFirmware', purpose: 'Firmware architecture, OTA mechanism, version control, rollback.', example: 'A/B partition OTA with automatic rollback on boot failure.' },
  { method: 'setCertifications', purpose: 'Regulatory certifications with status/target date.', example: [{ certification: 'FCC', region: 'US', status: 'Pending', targetDate: 'Q4 2026' }] },
  { method: 'setEnvironmentalRequirements', purpose: 'Operating temp/humidity, ingress protection, drop/shock tolerance.', example: '-10C to 45C operating range, IP54, 1m drop tested.' },
  { method: 'setManufacturingNotes', purpose: 'Assembly summary, contract manufacturer, tooling, capacity/yield.', example: 'Contract manufactured in Shenzhen; target yield 98%.' },
  { method: 'setMaintenance', purpose: 'Field-serviceable parts, calibration schedule, lifespan, RMA process.', example: 'Battery is field-replaceable; 2-year expected lifespan.' },
  { method: 'setDesignTools', purpose: 'Design/documentation tooling with links.', example: [{ category: 'Design File', tool: 'Figma', link: 'https://figma.com/...' }] },
  { method: 'setDesignPrinciples', purpose: 'Core UX philosophy guiding design decisions.', example: 'Clarity and speed first: minimal-friction interactions.' },
  { method: 'setLayoutGrid', purpose: 'Breakpoint/grid/margin/gutter values.', example: [{ property: 'Grid System', value: '12-column grid' }] },
  { method: 'setTypography', purpose: 'Font family and available weights/sizes.', example: [{ weight: 'Font family', sizes: 'Inter' }, { weight: 'Regular', sizes: '12px, 16px, 20px' }] },
  { method: 'setColorPalette', purpose: 'Brand/UI color roles with hex values.', example: [{ role: 'Primary', color: 'Blue', hex: '#0055FF' }] },
  { method: 'setComponentsStates', purpose: 'UI components and their visual states.', example: [{ component: 'Primary Button', state: 'Default', behavior: 'Blue background, white text' }] },
  { method: 'setResponsiveBehavior', purpose: 'Breakpoints and device-specific notes.', example: [{ breakpoint: '<768px', device: 'Mobile', notes: 'Single column layout' }] },
  { method: 'setInteractionAnimation', purpose: 'Transitions, micro-interactions, loading states.', example: [{ aspect: 'Transitions', notes: 'Fade in, 200ms' }] },
  { method: 'setRevisionHistory', purpose: 'Version history of the product/design.', example: [{ version: 'v1.0', date: '2026-01-01', changes: 'Initial release' }] },
  { method: 'setTestStrategy', purpose: 'Overall QA approach: manual vs automated, environments, tools.', example: 'Automated unit + integration tests in CI; manual exploratory testing before release.' },
  { method: 'setTestPlan', purpose: 'Test phases with entry/exit criteria.', example: [{ phase: 'Regression', scope: 'Full app', entryCriteria: 'Feature complete', exitCriteria: 'Zero P1 bugs' }] },
  { method: 'setTestCases', purpose: 'Individual test cases with steps and expected results.', example: [{ id: 'TC-001', description: 'Login with valid credentials', steps: '1. Open app 2. Enter creds 3. Submit', expectedResult: 'User is logged in', status: 'Pass' }] },
  { method: 'setBugTracking', purpose: 'Severity definitions and SLA to fix.', example: [{ severity: 'Critical', definition: 'Data loss or outage', sla: '4 hours' }] },
  { method: 'setSecurityTesting', purpose: 'Pen-test schedule, vulnerability scanning tools, findings tracker link.', example: 'Quarterly third-party pen test; Snyk for dependency scanning.' },
  { method: 'setHardwareValidation', purpose: 'Hardware validation tests (EMC, drop, thermal, etc.) and results. Only rendered if hardware section is enabled.', example: [{ type: 'Drop Test', description: '1m onto concrete', standard: 'IEC 60068-2-32', result: 'Pass' }] },
];

// =============================================================================
// MAIN SDK CLASS
// =============================================================================

class ProductSpecSDK {
  constructor() {
    this.data = {};
    this.hardwareEnabled = false;
  }

  /** Returns the machine-readable section manifest described above. */
  static sectionGuide() {
    return SECTION_GUIDE;
  }

  // ---------------------------------------------------------------------
  // COVER PAGE
  // ---------------------------------------------------------------------
  /**
   * Cover page. This is the very first thing a reader sees.
   * @param {Object} p
   * @param {string} p.productName - The product/application name (large title text).
   * @param {string} p.shortDescription - One sentence describing the product, shown italic under the title.
   * @param {string} p.lastUpdated - Date this document was last updated, e.g. "2026-07-22".
   * @param {string} p.status - Document status, e.g. "Draft" / "In Review" / "Approved".
   * @param {string} [p.logoImagePath] - Optional path to a logo/hero image to embed on the cover. If omitted, a placeholder box is drawn (matching the template).
   */
  setCoverPage(p = {}) {
    this.data.cover = p;
    return this;
  }

  /**
   * The small identity table at the top of "Product Overview".
   * @param {Object} p
   * @param {string} p.productName
   * @param {string} p.version - e.g. "v1.0", or "v2.1 (revision — prior: v2.0)".
   * @param {string} p.status - Free-text current build/rollout status.
   */
  setProductInfo(p = {}) {
    this.data.productInfo = p;
    return this;
  }

  /**
   * Executive Summary: a one-page overview for any stakeholder (executive to
   * new hire) — what the product is, why it exists, and how success is measured.
   * @param {string} text
   */
  setExecutiveSummary(text) {
    this.data.executiveSummary = text;
    return this;
  }

  /**
   * Product Roadmap Summary table.
   * @param {Array<{phase:string, theme:string, timeframe:string, status:string}>} rows
   */
  setProductRoadmap(rows) {
    this.data.productRoadmap = rows;
    return this;
  }

  // ---------------------------------------------------------------------
  // 2.1 BUSINESS PERSPECTIVE
  // ---------------------------------------------------------------------
  /**
   * Target Market: who this product is for, segmented.
   * @param {Array<{segment:string, description:string}>} rows
   */
  setTargetMarket(rows) {
    this.data.targetMarket = rows;
    return this;
  }

  /**
   * Customer Personas: representative buyers/users.
   * @param {Array<{persona:string, role:string, goals:string, painPoints:string, buyingBehavior:string}>} rows
   */
  setCustomerPersonas(rows) {
    this.data.customerPersonas = rows;
    return this;
  }

  /**
   * User Journey: a diagram/screenshot showing the end-to-end flow.
   * @param {Object} p
   * @param {string} [p.imagePath] - Path to the journey image. If omitted, a placeholder box is drawn.
   * @param {string} [p.caption] - Optional caption under the image.
   * @param {number} [p.widthPx=450] @param {number} [p.heightPx=540]
   */
  setUserJourney(p = {}) {
    this.data.userJourney = p;
    return this;
  }

  /**
   * Use Cases: concrete actor/trigger/outcome scenarios.
   * @param {Array<{useCase:string, actor:string, trigger:string, outcome:string}>} rows
   */
  setUseCases(rows) {
    this.data.useCases = rows;
    return this;
  }

  /**
   * Competitive Analysis: how this product stacks up against alternatives.
   * @param {Array<{competitor:string, strengths:string, weaknesses:string, pricing:string, differentiator:string}>} rows
   */
  setCompetitiveAnalysis(rows) {
    this.data.competitiveAnalysis = rows;
    return this;
  }

  /**
   * Pricing Model: overall approach plus tier breakdown.
   * @param {Object} p
   * @param {string} p.modelDescription - e.g. "Subscription tiers / usage-based / freemium / hybrid, etc."
   * @param {Array<{tier:string, price:string, includes:string, targetSegment:string}>} p.tiers
   */
  setPricingModel(p = {}) {
    this.data.pricingModel = p;
    return this;
  }

  /**
   * Revenue Model: how the product actually makes money (narrative).
   * @param {string} text
   */
  setRevenueModel(text) {
    this.data.revenueModel = text;
    return this;
  }

  // ---------------------------------------------------------------------
  // 2.2 FUNCTIONAL SPECIFICATION
  // ---------------------------------------------------------------------
  /**
   * Features list: what the product does, with priority.
   * @param {Array<{name:string, description:string, priority:('Must'|'Should'|'Could'|string)}>} rows
   */
  setFeatures(rows) {
    this.data.features = rows;
    return this;
  }

  // ---------------------------------------------------------------------
  // 2.3 TECHNICAL SPECIFICATION
  // ---------------------------------------------------------------------
  /**
   * System Architecture: narrative description (monolith/microservices/
   * serverless, etc.) plus an optional architecture diagram.
   * @param {Object} p
   * @param {string} p.description
   * @param {string} [p.diagramImagePath]
   */
  setSystemArchitecture(p = {}) {
    this.data.systemArchitecture = p;
    return this;
  }

  /**
   * Technology Stack: what's used at each layer, and why.
   * @param {Array<{layer:string, technology:string, justification:string}>} rows
   */
  setTechnologyStack(rows) {
    this.data.technologyStack = rows;
    return this;
  }

  /**
   * APIs: key endpoints, plus a link to full API documentation.
   * @param {Object} p
   * @param {Array<{endpoint:string, method:string, description:string, authRequired:string, role:string}>} p.rows
   * @param {string} [p.docsLink] - Link to full Swagger/OpenAPI/Postman docs.
   */
  setApis(p = {}) {
    this.data.apis = p;
    return this;
  }

  /**
   * Database Design: tables/collections, plus an optional ER diagram.
   * @param {Object} p
   * @param {Array<{table:string, purpose:string, keyFields:string, url:string}>} p.rows
   * @param {string} [p.diagramImagePath]
   */
  setDatabaseDesign(p = {}) {
    this.data.databaseDesign = p;
    return this;
  }

  /**
   * Authentication method: OAuth2 / SSO / SAML / JWT / API keys, MFA, sessions.
   * @param {string} text
   */
  setAuthentication(text) {
    this.data.authentication = text;
    return this;
  }

  /**
   * Security: controls as a nested list (tool -> sub-features) plus free-text
   * notes on encryption, secrets management, pen-testing cadence, dependency scanning.
   * @param {Object} p
   * @param {Array<{title:string, children?:string[]}>} p.measures
   * @param {string} [p.notes]
   */
  setSecurity(p = {}) {
    this.data.security = p;
    return this;
  }

  /** Monitoring tools and key alerts/dashboards. @param {string} text */
  setMonitoring(text) {
    this.data.monitoring = text;
    return this;
  }

  /** Log levels, retention policy, centralized logging tool, PII handling. @param {string} text */
  setLogging(text) {
    this.data.logging = text;
    return this;
  }

  /** Backup frequency, retention, RTO, RPO. @param {string} text */
  setBackupStrategy(text) {
    this.data.backupStrategy = text;
    return this;
  }

  // ---------------------------------------------------------------------
  // 2.4 HARDWARE SPECIFICATION (optional — pure SaaS products can skip this)
  // ---------------------------------------------------------------------
  /** Turn the Hardware Specification section (and Hardware Validation) on/off. @param {boolean} enabled */
  setEnableHardwareSection(enabled) {
    this.hardwareEnabled = !!enabled;
    return this;
  }

  /** High-level device description, form factor, role in the product. @param {string} text */
  setHardwareOverview(text) {
    this.data.hardwareOverview = text;
    return this;
  }

  /**
   * Major hardware components.
   * @param {Array<{component:string, function:string, vendor:string, altVendor:string}>} rows
   */
  setHardwareComponents(rows) {
    this.data.hardwareComponents = rows;
    return this;
  }

  /**
   * Bill of Materials.
   * @param {Array<{partNo:string, description:string, qty:string, unitCost:string, vendor:string, leadTime:string}>} rows
   */
  setBillOfMaterials(rows) {
    this.data.billOfMaterials = rows;
    return this;
  }

  /**
   * Mechanical Design: enclosure material, dimensions, weight, tolerances, IP rating.
   * @param {Object} p @param {string} p.description @param {string} [p.diagramImagePath]
   */
  setMechanicalDesign(p = {}) {
    this.data.mechanicalDesign = p;
    return this;
  }

  /**
   * Electrical Specification: voltage, current draw, PCB notes.
   * @param {Object} p @param {string} p.description @param {string} [p.diagramImagePath]
   */
  setElectricalSpecification(p = {}) {
    this.data.electricalSpecification = p;
    return this;
  }

  /**
   * Sensors onboard the device.
   * @param {Array<{sensor:string, type:string, range:string, accuracy:string, purpose:string}>} rows
   */
  setSensors(rows) {
    this.data.sensors = rows;
    return this;
  }

  /** Wi-Fi/BLE/LTE/Zigbee/LoRa/Ethernet, protocols, pairing, OTA support. @param {string} text */
  setConnectivity(text) {
    this.data.connectivity = text;
    return this;
  }

  /** Power source, battery life, charging method, consumption profile. @param {string} text */
  setPowerRequirements(text) {
    this.data.powerRequirements = text;
    return this;
  }

  /** Firmware architecture, OTA mechanism, version control, rollback support. @param {string} text */
  setFirmware(text) {
    this.data.firmware = text;
    return this;
  }

  /**
   * Regulatory certifications.
   * @param {Array<{certification:string, region:string, status:string, targetDate:string}>} rows
   */
  setCertifications(rows) {
    this.data.certifications = rows;
    return this;
  }

  /** Operating temp/humidity range, ingress protection, drop/shock tolerance. @param {string} text */
  setEnvironmentalRequirements(text) {
    this.data.environmentalRequirements = text;
    return this;
  }

  /** Assembly summary, contract manufacturer, tooling, capacity/yield targets. @param {string} text */
  setManufacturingNotes(text) {
    this.data.manufacturingNotes = text;
    return this;
  }

  /** Field-serviceable parts, calibration schedule, expected lifespan, RMA process. @param {string} text */
  setMaintenance(text) {
    this.data.maintenance = text;
    return this;
  }

  // ---------------------------------------------------------------------
  // 3. UI/UX SPECIFICATION
  // ---------------------------------------------------------------------
  /**
   * Design & documentation tooling.
   * @param {Array<{category:string, tool:string, link:string}>} rows
   */
  setDesignTools(rows) {
    this.data.designTools = rows;
    return this;
  }

  /** Core UX philosophy guiding design decisions. @param {string} text */
  setDesignPrinciples(text) {
    this.data.designPrinciples = text;
    return this;
  }

  /**
   * Layout & Grid values (breakpoint, grid system, margin, gutter, etc.)
   * @param {Array<{property:string, value:string}>} rows
   */
  setLayoutGrid(rows) {
    this.data.layoutGrid = rows;
    return this;
  }

  /**
   * Typography: font family and the weights/sizes available.
   * @param {Array<{weight:string, sizes:string}>} rows - first row is typically {weight:'Font family', sizes:'<font name>'}.
   */
  setTypography(rows) {
    this.data.typography = rows;
    return this;
  }

  /**
   * Color palette.
   * @param {Array<{role:string, color:string, hex:string}>} rows
   */
  setColorPalette(rows) {
    this.data.colorPalette = rows;
    return this;
  }

  /**
   * UI components and their states.
   * @param {Array<{component:string, state:string, behavior:string}>} rows
   */
  setComponentsStates(rows) {
    this.data.componentsStates = rows;
    return this;
  }

  /**
   * Responsive breakpoints and device-specific notes.
   * @param {Array<{breakpoint:string, device:string, notes:string}>} rows
   */
  setResponsiveBehavior(rows) {
    this.data.responsiveBehavior = rows;
    return this;
  }

  /**
   * Interaction & animation notes.
   * @param {Array<{aspect:string, notes:string}>} rows
   */
  setInteractionAnimation(rows) {
    this.data.interactionAnimation = rows;
    return this;
  }

  /**
   * Design/product revision history.
   * @param {Array<{version:string, date:string, changes:string}>} rows
   */
  setRevisionHistory(rows) {
    this.data.revisionHistory = rows;
    return this;
  }

  // ---------------------------------------------------------------------
  // 4. QUALITY CONTROL
  // ---------------------------------------------------------------------
  /** Overall QA approach: manual vs automated, environments, tools used. @param {string} text */
  setTestStrategy(text) {
    this.data.testStrategy = text;
    return this;
  }

  /**
   * Test plan phases with entry/exit criteria.
   * @param {Array<{phase:string, scope:string, entryCriteria:string, exitCriteria:string}>} rows
   */
  setTestPlan(rows) {
    this.data.testPlan = rows;
    return this;
  }

  /**
   * Individual test cases.
   * @param {Array<{id:string, description:string, steps:string, expectedResult:string, status:string}>} rows
   */
  setTestCases(rows) {
    this.data.testCases = rows;
    return this;
  }

  /**
   * Bug severity definitions and SLA to fix.
   * @param {Array<{severity:string, definition:string, sla:string}>} rows
   */
  setBugTracking(rows) {
    this.data.bugTracking = rows;
    return this;
  }

  /** Pen-test schedule, vulnerability scanning tools, findings tracker link. @param {string} text */
  setSecurityTesting(text) {
    this.data.securityTesting = text;
    return this;
  }

  /**
   * Hardware validation tests. Only rendered if the hardware section is enabled.
   * @param {Array<{type:string, description:string, standard:string, result:string}>} rows
   */
  setHardwareValidation(rows) {
    this.data.hardwareValidation = rows;
    return this;
  }

  // =========================================================================
  // BUILD / GENERATE
  // =========================================================================

  _buildCover() {
    const c = this.data.cover || {};
    const children = [];
    children.push(new Paragraph({
      heading: HeadingLevel.TITLE,
      alignment: AlignmentType.CENTER,
      children: [run(c.productName || '[Application Name]', { font: FONT, size: SIZE.appName, bold: false })],
    }));
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [run(c.shortDescription || '[short description of the product]', { italics: true, color: COLOR.purposeText })],
    }));
    children.push(new Paragraph({ children: [] }));
    children.push(imageOrPlaceholder(c.logoImagePath, 300, 220, 'IMAGE'));
    children.push(new Paragraph({ children: [] }));
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 300 },
      children: [run('Product Specification', { size: SIZE.docTitle })],
    }));
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        run('Last Updated: ', { bold: true }),
        run(c.lastUpdated || '[YYYY-MM-DD]', { italics: true, color: COLOR.coverValue }),
      ],
    }));
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        run('Status: ', { bold: true }),
        run(c.status || '[Draft / In Review / Approved]', { italics: true, color: COLOR.coverValue }),
      ],
    }));
    children.push(pageBreak());
    return children;
  }

  _buildProductOverview() {
    const children = [];
    children.push(heading('1. Product Overview', 1));
    children.push(purposeLine('Give any stakeholder — from an executive to a new hire — a one-page overview of what this product is, why it exists, and how success is measured.'));

    // The identity table has bold labels in column 1 and no header row, so
    // it's built directly rather than via buildTable().
    const info = this.data.productInfo || {};
    children.push(this._identityTable(info));

    children.push(heading('Executive Summary', 2));
    children.push(...textBlock(this.data.executiveSummary, '[Write a concise executive summary: what the product is, why it exists, and how success is measured.]'));

    children.push(heading('Product Roadmap Summary', 2));
    children.push(buildTable(
      [
        { key: 'phase', header: 'Phase', weight: 1 },
        { key: 'theme', header: 'Theme', weight: 1.3 },
        { key: 'timeframe', header: 'Timeframe', weight: 1 },
        { key: 'status', header: 'Status', weight: 1 },
      ],
      this.data.productRoadmap,
      [{ phase: PLACEHOLDER_ROW_TEXT, theme: PLACEHOLDER_ROW_TEXT, timeframe: PLACEHOLDER_ROW_TEXT, status: PLACEHOLDER_ROW_TEXT }],
    ));
    children.push(pageBreak());
    return children;
  }

  _identityTable(info) {
    const rows = [
      ['Product Name', info.productName || '[Product Name]'],
      ['Version', info.version || '[Version]'],
      ['Status', info.status || '[Status]'],
    ];
    const w1 = Math.round(CONTENT_WIDTH * 0.25);
    const w2 = CONTENT_WIDTH - w1;
    const isPlaceholder = !info || (!info.productName && !info.version && !info.status);
    return new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [w1, w2],
      borders: borderSet(),
      rows: rows.map(([label, value]) => new TableRow({
        children: [
          new TableCell({
            width: { size: w1, type: WidthType.DXA },
            margins: { top: 70, bottom: 70, left: 100, right: 100 },
            children: [new Paragraph({ children: [run(label, { bold: true })] })],
          }),
          new TableCell({
            width: { size: w2, type: WidthType.DXA },
            margins: { top: 70, bottom: 70, left: 100, right: 100 },
            children: [new Paragraph({ children: [run(value, { italics: isPlaceholder })] })],
          }),
        ],
      })),
    });
  }

  _buildBusinessPerspective() {
    const children = [];
    children.push(heading('2. Product Specification', 1));
    children.push(heading('2.1 Business Perspective', 2));
    children.push(purposeLine('Establish the commercial rationale — market, customers, competition, and economics — before diving into functionality.'));

    children.push(heading('Target Market', 3));
    children.push(buildTable(
      [{ key: 'segment', header: 'Segment', weight: 1 }, { key: 'description', header: 'Description', weight: 2.5 }],
      this.data.targetMarket,
      [{ segment: PLACEHOLDER_ROW_TEXT, description: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Customer Personas', 3));
    children.push(buildTable(
      [
        { key: 'persona', header: 'Persona', weight: 1 },
        { key: 'role', header: 'Role', weight: 1 },
        { key: 'goals', header: 'Goals', weight: 1.3 },
        { key: 'painPoints', header: 'Pain Points', weight: 1.3 },
        { key: 'buyingBehavior', header: 'Buying Behavior', weight: 1.3 },
      ],
      this.data.customerPersonas,
      [{ persona: PLACEHOLDER_ROW_TEXT, role: PLACEHOLDER_ROW_TEXT, goals: PLACEHOLDER_ROW_TEXT, painPoints: PLACEHOLDER_ROW_TEXT, buyingBehavior: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('User Journey', 3));
    const uj = this.data.userJourney || {};
    children.push(imageOrPlaceholder(uj.imagePath, uj.widthPx || 450, uj.heightPx || 540, '[Insert user-journey diagram/screenshot]'));
    if (uj.caption) children.push(bodyPara(uj.caption, { italics: true }));

    children.push(heading('Use Cases', 3));
    children.push(buildTable(
      [
        { key: 'useCase', header: 'Use Case', weight: 1.3 },
        { key: 'actor', header: 'Actor', weight: 1 },
        { key: 'trigger', header: 'Trigger', weight: 1.3 },
        { key: 'outcome', header: 'Outcome', weight: 1.3 },
      ],
      this.data.useCases,
      [{ useCase: PLACEHOLDER_ROW_TEXT, actor: PLACEHOLDER_ROW_TEXT, trigger: PLACEHOLDER_ROW_TEXT, outcome: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Competitive Analysis', 3));
    children.push(buildTable(
      [
        { key: 'competitor', header: 'Competitor', weight: 1 },
        { key: 'strengths', header: 'Strengths', weight: 1 },
        { key: 'weaknesses', header: 'Weaknesses', weight: 1 },
        { key: 'pricing', header: 'Pricing', weight: 1 },
        { key: 'differentiator', header: 'Differentiator', weight: 1.2 },
      ],
      this.data.competitiveAnalysis,
      [{ competitor: PLACEHOLDER_ROW_TEXT, strengths: PLACEHOLDER_ROW_TEXT, weaknesses: PLACEHOLDER_ROW_TEXT, pricing: PLACEHOLDER_ROW_TEXT, differentiator: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Pricing Model', 3));
    const pm = this.data.pricingModel || {};
    children.push(...textBlock(pm.modelDescription, '[Subscription tiers / usage-based / one-time / freemium / hardware+subscription hybrid, etc.]'));
    children.push(buildTable(
      [
        { key: 'tier', header: 'Tier', weight: 1 },
        { key: 'price', header: 'Price', weight: 1 },
        { key: 'includes', header: 'Includes', weight: 1.5 },
        { key: 'targetSegment', header: 'Target Segment', weight: 1.2 },
      ],
      pm.tiers,
      [{ tier: PLACEHOLDER_ROW_TEXT, price: PLACEHOLDER_ROW_TEXT, includes: PLACEHOLDER_ROW_TEXT, targetSegment: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Revenue Model', 3));
    children.push(...textBlock(this.data.revenueModel, '[Describe how the product makes money: subscription, transaction fees, licensing, hardware margin, etc.]'));

    return children;
  }

  _buildFunctionalSpecification() {
    const children = [];
    children.push(heading('2.2 Functional Specification', 2));
    children.push(purposeLine('Define exactly what the product does — the behaviors, rules, and logic that engineering must build and QA must verify.'));

    children.push(heading('Features', 3));
    children.push(buildTable(
      [
        { key: 'name', header: 'Feature Name', weight: 1 },
        { key: 'description', header: 'Description', weight: 2 },
        { key: 'priority', header: 'Priority', weight: 0.7 },
      ],
      this.data.features,
      [{ name: PLACEHOLDER_ROW_TEXT, description: PLACEHOLDER_ROW_TEXT, priority: PLACEHOLDER_ROW_TEXT }],
    ));
    return children;
  }

  _buildTechnicalSpecification() {
    const children = [];
    children.push(heading('2.3 Technical Specification', 2));
    children.push(purposeLine('The engineering blueprint — architecture, stack, and non-functional requirements needed to build, scale, and operate the system.'));

    children.push(heading('System Architecture', 3));
    const sa = this.data.systemArchitecture || {};
    children.push(bodyPara('🖼 Insert architecture diagram — components, services, data flow', { italics: true, color: COLOR.purposeText }));
    if (sa.diagramImagePath) children.push(imageOrPlaceholder(sa.diagramImagePath, 500, 320));
    children.push(...textBlock(sa.description, '[Narrative description of architecture style: microservices/monolith/serverless, etc.]'));

    children.push(heading('Technology Stack', 3));
    children.push(buildTable(
      [
        { key: 'layer', header: 'Layer', weight: 1 },
        { key: 'technology', header: 'Technology', weight: 1.2 },
        { key: 'justification', header: 'Justification', weight: 1.6 },
      ],
      this.data.technologyStack,
      [{ layer: PLACEHOLDER_ROW_TEXT, technology: PLACEHOLDER_ROW_TEXT, justification: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('APIs', 3));
    const apis = this.data.apis || {};
    children.push(buildTable(
      [
        { key: 'endpoint', header: 'Endpoint', weight: 1.2 },
        { key: 'method', header: 'Method', weight: 0.8 },
        { key: 'description', header: 'Description', weight: 1.5 },
        { key: 'authRequired', header: 'Auth Required', weight: 0.9 },
        { key: 'role', header: 'Role', weight: 0.8 },
      ],
      apis.rows,
      [{ endpoint: PLACEHOLDER_ROW_TEXT, method: PLACEHOLDER_ROW_TEXT, description: PLACEHOLDER_ROW_TEXT, authRequired: PLACEHOLDER_ROW_TEXT, role: PLACEHOLDER_ROW_TEXT }],
    ));
    children.push(...textBlock(apis.docsLink, '[Link to full API documentation — Swagger/OpenAPI/Postman collection]'));

    children.push(heading('Database Design', 3));
    const db = this.data.databaseDesign || {};
    children.push(bodyPara('🖼 Insert ER diagram', { italics: true, color: COLOR.purposeText }));
    if (db.diagramImagePath) children.push(imageOrPlaceholder(db.diagramImagePath, 500, 320));
    children.push(buildTable(
      [
        { key: 'table', header: 'Table/Collection', weight: 1 },
        { key: 'purpose', header: 'Purpose', weight: 1.3 },
        { key: 'keyFields', header: 'Key Main Fields', weight: 1 },
        { key: 'url', header: 'URL', weight: 1 },
      ],
      db.rows,
      [{ table: PLACEHOLDER_ROW_TEXT, purpose: PLACEHOLDER_ROW_TEXT, keyFields: PLACEHOLDER_ROW_TEXT, url: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Authentication', 3));
    children.push(...textBlock(this.data.authentication, '[Method: OAuth2 / SSO / SAML / JWT / API keys. MFA requirements. Session management.]'));

    children.push(heading('Security', 3));
    const sec = this.data.security || {};
    if (sec.measures && sec.measures.length) {
      children.push(...numberedBlock(sec.measures));
    } else {
      children.push(...numberedBlock([{ title: '[Security tool/control]', children: ['[Specific feature enabled]'] }]));
    }
    children.push(...textBlock(sec.notes, '[Data encryption at rest/in transit, secrets management, pen testing cadence, dependency scanning.]'));

    children.push(heading('Monitoring', 3));
    children.push(...textBlock(this.data.monitoring, '[Tools — Datadog/Grafana/New Relic. Key alerts and dashboards.]'));

    children.push(heading('Logging', 3));
    children.push(...textBlock(this.data.logging, '[Log levels, retention policy, centralized logging tool, PII handling in logs.]'));

    children.push(heading('Backup Strategy', 3));
    children.push(...textBlock(this.data.backupStrategy, '[Backup frequency, retention, recovery time objective (RTO), recovery point objective (RPO).]'));

    return children;
  }

  _buildHardwareSpecification() {
    const children = [];
    children.push(heading('2.4 Hardware Specification', 2));
    children.push(bodyPara('🔧 HARDWARE ONLY — optional for pure SaaS', { bold: true }));
    children.push(purposeLine('Full technical and manufacturing definition of the physical device(s) that make up the product.'));

    children.push(heading('Hardware Overview', 3));
    children.push(...textBlock(this.data.hardwareOverview, '[High-level description of the device, form factor, and role in the overall product.]'));

    children.push(heading('Components', 3));
    children.push(buildTable(
      [
        { key: 'component', header: 'Component', weight: 1 },
        { key: 'function', header: 'Function', weight: 1.2 },
        { key: 'vendor', header: 'Vendor', weight: 1 },
        { key: 'altVendor', header: 'Alternative/Backup Vendor', weight: 1.2 },
      ],
      this.data.hardwareComponents,
      [{ component: PLACEHOLDER_ROW_TEXT, function: PLACEHOLDER_ROW_TEXT, vendor: PLACEHOLDER_ROW_TEXT, altVendor: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Bill of Materials (BOM)', 3));
    children.push(buildTable(
      [
        { key: 'partNo', header: 'Part No.', weight: 0.8 },
        { key: 'description', header: 'Description', weight: 1.4 },
        { key: 'qty', header: 'Qty', weight: 0.6 },
        { key: 'unitCost', header: 'Unit Cost', weight: 0.8 },
        { key: 'vendor', header: 'Vendor', weight: 0.9 },
        { key: 'leadTime', header: 'Lead Time', weight: 0.8 },
      ],
      this.data.billOfMaterials,
      [{ partNo: PLACEHOLDER_ROW_TEXT, description: PLACEHOLDER_ROW_TEXT, qty: PLACEHOLDER_ROW_TEXT, unitCost: PLACEHOLDER_ROW_TEXT, vendor: PLACEHOLDER_ROW_TEXT, leadTime: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Mechanical Design', 3));
    const md = this.data.mechanicalDesign || {};
    children.push(bodyPara('🖼 Insert CAD renders / mechanical drawings', { italics: true, color: COLOR.purposeText }));
    if (md.diagramImagePath) children.push(imageOrPlaceholder(md.diagramImagePath, 500, 320));
    children.push(...textBlock(md.description, '[Enclosure material, dimensions, weight, tolerances, IP rating.]'));

    children.push(heading('Electrical Specification', 3));
    const es = this.data.electricalSpecification || {};
    children.push(...textBlock(es.description, '[Voltage, current draw, PCB design notes, schematics reference.]'));
    children.push(bodyPara('🖼 Insert circuit diagram / PCB layout', { italics: true, color: COLOR.purposeText }));
    if (es.diagramImagePath) children.push(imageOrPlaceholder(es.diagramImagePath, 500, 320));

    children.push(heading('Sensors', 3));
    children.push(buildTable(
      [
        { key: 'sensor', header: 'Sensor', weight: 1 },
        { key: 'type', header: 'Type', weight: 1 },
        { key: 'range', header: 'Range', weight: 1 },
        { key: 'accuracy', header: 'Accuracy', weight: 1 },
        { key: 'purpose', header: 'Purpose', weight: 1 },
      ],
      this.data.sensors,
      [{ sensor: PLACEHOLDER_ROW_TEXT, type: PLACEHOLDER_ROW_TEXT, range: PLACEHOLDER_ROW_TEXT, accuracy: PLACEHOLDER_ROW_TEXT, purpose: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Connectivity', 3));
    children.push(...textBlock(this.data.connectivity, '[Wi-Fi / BLE / LTE / Zigbee / LoRa / Ethernet. Protocols, pairing method, OTA update support.]'));

    children.push(heading('Power Requirements', 3));
    children.push(...textBlock(this.data.powerRequirements, '[Power source — battery/mains/PoE. Battery life estimates, charging method, power consumption profile.]'));

    children.push(heading('Firmware', 3));
    children.push(...textBlock(this.data.firmware, '[Firmware architecture, update mechanism (OTA), version control, rollback support.]'));

    children.push(heading('Certifications', 3));
    children.push(buildTable(
      [
        { key: 'certification', header: 'Certification', weight: 1.1 },
        { key: 'region', header: 'Region', weight: 0.8 },
        { key: 'status', header: 'Status', weight: 0.9 },
        { key: 'targetDate', header: 'Target Date', weight: 0.9 },
      ],
      this.data.certifications,
      [{ certification: '[FCC/CE/RoHS/UL/etc.]', region: PLACEHOLDER_ROW_TEXT, status: '[Pending/Obtained]', targetDate: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Environmental Requirements', 3));
    children.push(...textBlock(this.data.environmentalRequirements, '[Operating temperature/humidity range, ingress protection, drop/shock tolerance.]'));

    children.push(heading('Manufacturing Notes', 3));
    children.push(...textBlock(this.data.manufacturingNotes, '[Assembly instructions summary, contract manufacturer, tooling requirements, production capacity/yield targets.]'));

    children.push(heading('Maintenance', 3));
    children.push(...textBlock(this.data.maintenance, '[Field-serviceable parts, calibration schedule, expected device lifespan, RMA process.]'));

    children.push(pageBreak());
    return children;
  }

  _buildUiUx() {
    const children = [];
    children.push(heading('3. UI/UX Specification', 1));
    children.push(purposeLine('Define the experience layer — how users interact with the product visually and functionally.'));

    children.push(heading('Design & Documentation Tools', 2));
    children.push(buildTable(
      [
        { key: 'category', header: 'Category', weight: 1 },
        { key: 'tool', header: 'Tool Used', weight: 1 },
        { key: 'link', header: 'Link / Location', weight: 1.5 },
      ],
      this.data.designTools,
      [{ category: PLACEHOLDER_ROW_TEXT, tool: PLACEHOLDER_ROW_TEXT, link: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Design Principles', 2));
    children.push(...textBlock(this.data.designPrinciples, '[Describe the core design philosophy guiding UX decisions — e.g. clarity first, minimal friction, accessibility.]'));

    children.push(heading('Design System', 2));

    children.push(heading('Layout & Grid', 3));
    children.push(buildTable(
      [{ key: 'property', header: 'Property', weight: 1 }, { key: 'value', header: 'Value', weight: 1.5 }],
      this.data.layoutGrid,
      [{ property: PLACEHOLDER_ROW_TEXT, value: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Typography', 3));
    children.push(buildTable(
      [{ key: 'weight', header: 'Weight', weight: 1 }, { key: 'sizes', header: 'Sizes', weight: 1.5 }],
      this.data.typography,
      [{ weight: '[Font family]', sizes: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Color Palette', 3));
    children.push(buildTable(
      [
        { key: 'role', header: 'Role', weight: 1 },
        { key: 'color', header: 'Color', weight: 1 },
        { key: 'hex', header: 'Hex', weight: 1 },
      ],
      this.data.colorPalette,
      [{ role: PLACEHOLDER_ROW_TEXT, color: PLACEHOLDER_ROW_TEXT, hex: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Components & States', 3));
    children.push(buildTable(
      [
        { key: 'component', header: 'Component', weight: 1 },
        { key: 'state', header: 'State', weight: 1 },
        { key: 'behavior', header: 'Behavior / Description', weight: 1.6 },
      ],
      this.data.componentsStates,
      [{ component: PLACEHOLDER_ROW_TEXT, state: PLACEHOLDER_ROW_TEXT, behavior: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Responsive Behavior', 2));
    children.push(buildTable(
      [
        { key: 'breakpoint', header: 'Breakpoint', weight: 1 },
        { key: 'device', header: 'Device', weight: 1 },
        { key: 'notes', header: 'Notes', weight: 1.6 },
      ],
      this.data.responsiveBehavior,
      [{ breakpoint: PLACEHOLDER_ROW_TEXT, device: PLACEHOLDER_ROW_TEXT, notes: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Interaction & Animation', 2));
    children.push(buildTable(
      [{ key: 'aspect', header: 'Aspect', weight: 1 }, { key: 'notes', header: 'Notes', weight: 2 }],
      this.data.interactionAnimation,
      [{ aspect: PLACEHOLDER_ROW_TEXT, notes: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Revision History', 2));
    children.push(buildTable(
      [
        { key: 'version', header: 'Version', weight: 0.8 },
        { key: 'date', header: 'Date', weight: 0.8 },
        { key: 'changes', header: 'Changes Made', weight: 2 },
      ],
      this.data.revisionHistory,
      [{ version: PLACEHOLDER_ROW_TEXT, date: PLACEHOLDER_ROW_TEXT, changes: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(pageBreak());
    return children;
  }

  _buildQualityControl() {
    const children = [];
    children.push(heading('4. Quality Control', 1));
    children.push(purposeLine('Define how quality is verified before release, across software and (if applicable) hardware.'));

    children.push(heading('Test Strategy', 2));
    children.push(...textBlock(this.data.testStrategy, '[Overall approach: manual vs automated, environments, tools used.]'));

    children.push(heading('Test Plan', 2));
    children.push(buildTable(
      [
        { key: 'phase', header: 'Test Phase', weight: 1 },
        { key: 'scope', header: 'Scope', weight: 1.2 },
        { key: 'entryCriteria', header: 'Entry Criteria', weight: 1.2 },
        { key: 'exitCriteria', header: 'Exit Criteria', weight: 1.2 },
      ],
      this.data.testPlan,
      [{ phase: PLACEHOLDER_ROW_TEXT, scope: PLACEHOLDER_ROW_TEXT, entryCriteria: PLACEHOLDER_ROW_TEXT, exitCriteria: PLACEHOLDER_ROW_TEXT }],
    ));

    children.push(heading('Test Cases', 2));
    children.push(buildTable(
      [
        { key: 'id', header: 'Test Case ID', weight: 0.8 },
        { key: 'description', header: 'Description', weight: 1.4 },
        { key: 'steps', header: 'Steps', weight: 1.4 },
        { key: 'expectedResult', header: 'Expected Result', weight: 1.2 },
        { key: 'status', header: 'Status', weight: 0.8 },
      ],
      this.data.testCases,
      [{ id: 'TC-001', description: PLACEHOLDER_ROW_TEXT, steps: PLACEHOLDER_ROW_TEXT, expectedResult: PLACEHOLDER_ROW_TEXT, status: '[Pass/Fail/Blocked]' }],
    ));

    children.push(heading('Bug Tracking', 2));
    children.push(buildTable(
      [
        { key: 'severity', header: 'Severity', weight: 1 },
        { key: 'definition', header: 'Definition', weight: 1.6 },
        { key: 'sla', header: 'SLA to Fix', weight: 1 },
      ],
      this.data.bugTracking,
      ['Critical', 'High', 'Medium', 'Low'].map((severity) => ({ severity, definition: PLACEHOLDER_ROW_TEXT, sla: PLACEHOLDER_ROW_TEXT })),
    ));

    children.push(heading('Security Testing', 2));
    children.push(...textBlock(this.data.securityTesting, '[Pen test schedule, vulnerability scanning tools, findings tracker link.]'));

    if (this.hardwareEnabled) {
      children.push(heading('Hardware Validation', 2));
      children.push(bodyPara('🔧 HARDWARE ONLY', { bold: true }));
      children.push(buildTable(
        [
          { key: 'type', header: 'Validation Type', weight: 1 },
          { key: 'description', header: 'Description', weight: 1.4 },
          { key: 'standard', header: 'Standard/Reference', weight: 1.2 },
          { key: 'result', header: 'Result', weight: 1 },
        ],
        this.data.hardwareValidation,
        [{ type: '[EMC/EMI Testing]', description: PLACEHOLDER_ROW_TEXT, standard: PLACEHOLDER_ROW_TEXT, result: PLACEHOLDER_ROW_TEXT }],
      ));
    }

    return children;
  }

  /** Assembles the full document body (array of docx elements). */
  _buildChildren() {
    return [
      ...this._buildCover(),
      ...this._buildProductOverview(),
      ...this._buildBusinessPerspective(),
      ...this._buildFunctionalSpecification(),
      ...this._buildTechnicalSpecification(),
      ...(this.hardwareEnabled ? this._buildHardwareSpecification() : []),
      ...this._buildUiUx(),
      ...this._buildQualityControl(),
    ];
  }

  _buildHeader() {
    return new Header({
      children: [new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          run('Logo/[Product Name] ', { size: SIZE.headerFooter, bold: true, color: COLOR.headerFooterText }),
          run('\tProduct Specification | ', { size: SIZE.headerFooter, color: COLOR.headerFooterText }),
          run('CONFIDENTIAL', { size: SIZE.headerFooter, bold: true, color: COLOR.confidential }),
        ],
      })],
    });
  }

  _buildFooter() {
    return new Footer({
      children: [
        new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            run('Internal Only', { size: SIZE.headerFooter, color: COLOR.headerFooterText }),
            run('\t', { size: SIZE.headerFooter }),
            run('CONFIDENTIAL', { size: SIZE.headerFooter, bold: true, color: COLOR.confidential }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            run('Page ', { size: SIZE.headerFooter, color: COLOR.headerFooterText }),
            new TextRun({ children: [PageNumber.CURRENT], size: SIZE.headerFooter, color: COLOR.headerFooterText, font: FONT }),
            run(' of ', { size: SIZE.headerFooter, color: COLOR.headerFooterText }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: SIZE.headerFooter, color: COLOR.headerFooterText, font: FONT }),
          ],
        }),
      ],
    });
  }

  /** Builds the underlying `docx` Document object. */
  toDocument() {
    return new Document({
      styles: {
        default: { document: { run: { font: FONT, size: SIZE.body } } },
        paragraphStyles: [
          { id: 'Title', name: 'Title', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { font: FONT, size: SIZE.appName, bold: false, color: COLOR.black } },
          { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { font: FONT, size: SIZE.h1, bold: true, color: COLOR.h1 },
            paragraph: { spacing: { before: 400, after: 200 } } },
          { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { font: FONT, size: SIZE.h2, bold: true, color: COLOR.h2 },
            paragraph: { spacing: { before: 300, after: 150 } } },
          { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { font: FONT, size: SIZE.h3, bold: true, color: COLOR.h3 },
            paragraph: { spacing: { before: 200, after: 100 } } },
          { id: 'Heading4', name: 'Heading 4', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { font: FONT, size: SIZE.h4, italics: true, color: COLOR.h4 },
            paragraph: { spacing: { before: 0, after: 0 } } },
        ],
      },
      sections: [{
        properties: {
          page: {
            size: { width: PAGE.width, height: PAGE.height },
            margin: PAGE.margin,
          },
        },
        headers: { default: this._buildHeader() },
        footers: { default: this._buildFooter() },
        children: this._buildChildren(),
      }],
    });
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

module.exports = ProductSpecSDK;
module.exports.SECTION_GUIDE = SECTION_GUIDE;
module.exports.imageOrPlaceholder = imageOrPlaceholder;
module.exports.numberedBlock = numberedBlock;
module.exports.buildTable = buildTable;
