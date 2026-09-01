'use strict';
/**
 * =============================================================================
 *  UI/UX DOCUMENTATION (Part 7 of the Product Documentation Master)
 * =============================================================================
 * Regenerates "Part 7. UI/UX Documentation" exactly: title page, metadata
 * table, 8 numbered sections (Overview, User Flow, Design System/UI Kit
 * with 3.1-3.4 sub-sections, Interaction Details with 4.1-4.3 sub-sections,
 * Content & Copy, Accessibility, Responsive & Platform Notes, Revision
 * History).
 *
 * NOTE: Section 6's heading is reproduced verbatim from the source template
 * INCLUDING its authoring glitch — "6. AccessibilityColourr contrast meets
 * WCAG 2.1 AA: [Yes/No — notes]" — since this SDK's job is byte-for-byte
 * fidelity to the original, not to silently fix upstream typos. Override
 * via setAccessibilityHeading() if you want the corrected heading instead.
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

const BULLET_NUMBERING_REF = 'uiux-bullet-list';

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
const DEFAULT_OVERVIEW_BULLETS = [
  'Primary user goal: [...]',
  'Entry points: [where the user encounters this flow]',
  "Success state: [what 'done' looks like for the user]",
];
const DEFAULT_USER_FLOW_NOTE =
  '[Flow diagram/prototype link] — narrate the happy path step by step, referencing screen names used in Section 3.';
const DEFAULT_COLOR_PALETTE = [
  { role: 'Primary', hex: '[#______]', usage: 'Primary buttons, key actions' },
  { role: 'Secondary', hex: '[#______]', usage: 'Secondary actions, accents' },
  { role: 'Background', hex: '[#______]', usage: 'Page/app background' },
  { role: 'Surface', hex: '[#______]', usage: 'Cards, panels, modals' },
  { role: 'Text — Primary', hex: '[#______]', usage: 'Main body text' },
  { role: 'Text — Secondary', hex: '[#______]', usage: 'Muted/helper text' },
  {
    role: 'Success / Warning / Error',
    hex: '[#___/#___/#___]',
    usage: 'Status indicators, alerts',
  },
];
const DEFAULT_TYPOGRAPHY = [
  { style: 'Display / H1', spec: '[Font, size, weight]', usage: 'Page titles' },
  { style: 'H2 / H3', spec: '[Font, size, weight]', usage: 'Section headers' },
  { style: 'Body', spec: '[Font, size, weight]', usage: 'Default text' },
  { style: 'Caption / Label', spec: '[Font, size, weight]', usage: 'Form labels, helper text' },
];
const DEFAULT_COMPONENTS = [
  {
    component: 'Button — Primary',
    states: 'Default / Hover / Disabled / Loading',
    notes: '[Library/source link]',
  },
  {
    component: 'Input field',
    states: 'Default / Focus / Error / Disabled',
    notes: '[Library/source link]',
  },
  { component: 'Card', states: 'Default / Selected / Hover', notes: '[Library/source link]' },
  { component: 'Modal / Dialogue', states: 'Open / Closing', notes: '[Library/source link]' },
];
const DEFAULT_SPACING_GRID_NOTE =
  '[e.g. 8pt spacing scale (4/8/16/24/32/48px), 12-column grid, max content width 1200px]';
const DEFAULT_COMPONENT_BEHAVIOUR = ['[Component] — [expected interaction, e.g. hover, tap, drag]'];
const DEFAULT_EDGE_CASES = [
  '[e.g. What happens with a 200-character title?]',
  '[e.g. What happens if the network request times out mid-flow?]',
];
const DEFAULT_ANIMATION_NOTE =
  '[Transition specs, duration, easing — or link to a motion prototype]';
const DEFAULT_CONTENT_COPY = [
  { element: '[Button label]', copy: '"[Exact copy]"', notes: '[Character limit, tone notes]' },
  { element: '[Error message]', copy: '"[Exact copy]"', notes: '[When it appears]' },
];
const DEFAULT_ACCESSIBILITY_HEADING =
  '6. AccessibilityColourr contrast meets WCAG 2.1 AA: [Yes/No — notes]';
const DEFAULT_ACCESSIBILITY_BULLETS = [
  'Keyboard navigation/focus order: [...]',
  'Screen reader labels (alt text, ARIA): [...]',
  'Touch target sizes (mobile): [...]',
];
const DEFAULT_RESPONSIVE_NOTES = [
  { platform: 'Desktop / Web', notes: '[Layout differences]' },
  { platform: 'Mobile Web', notes: '[Layout differences]' },
  { platform: 'Native iOS / Android', notes: '[Platform-specific patterns]' },
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
    method: 'setOverviewBullets',
    purpose: 'Section 1 bulleted overview (goal/entry points/success state).',
    example: ['Primary user goal: submit an expense in under 2 minutes'],
  },
  {
    method: 'setUserFlowNote',
    purpose: 'Section 2 flow/prototype narration.',
    example: 'See figma.com/... — user opens app, taps +, fills form, submits.',
  },
  {
    method: 'setColorPalette',
    purpose: 'Section 3.1 Role/Hex/Usage table.',
    example: [{ role: 'Primary', hex: '#2F6690', usage: 'Primary buttons' }],
  },
  {
    method: 'setTypography',
    purpose: 'Section 3.2 Style/Font-Size-Weight/Usage table.',
    example: [{ style: 'Display / H1', spec: 'Inter, 28px, Bold', usage: 'Page titles' }],
  },
  {
    method: 'setComponents',
    purpose: 'Section 3.3 Component/States/Notes table.',
    example: [{ component: 'Button', states: 'Default/Hover', notes: 'shadcn/ui' }],
  },
  {
    method: 'setSpacingGridNote',
    purpose: 'Section 3.4 spacing/grid note.',
    example: '8pt spacing scale, 12-column grid.',
  },
  {
    method: 'setComponentBehaviour',
    purpose: 'Section 4.1 bulleted interaction notes.',
    example: ['Card — hover lifts with shadow'],
  },
  {
    method: 'setEdgeCases',
    purpose: 'Section 4.2 bulleted edge cases.',
    example: ['What happens with an empty state?'],
  },
  {
    method: 'setAnimationNote',
    purpose: 'Section 4.3 animation/motion note.',
    example: '200ms ease-out for all transitions.',
  },
  {
    method: 'setContentCopy',
    purpose: 'Section 5 Element/Copy/Notes table.',
    example: [{ element: 'Submit button', copy: '"Submit expense"', notes: 'Max 20 chars' }],
  },
  {
    method: 'setAccessibilityHeading',
    purpose:
      "Overrides Section 6's heading text (defaults to the source template's exact wording, including its glitch).",
    example: '6. Accessibility',
  },
  {
    method: 'setAccessibilityBullets',
    purpose: 'Section 6 bulleted accessibility notes.',
    example: ['Keyboard navigation fully supported'],
  },
  {
    method: 'setResponsiveNotes',
    purpose: 'Section 7 Platform/Notes table.',
    example: [{ platform: 'Desktop / Web', notes: 'Sidebar collapses below 1024px' }],
  },
  {
    method: 'setRevisionHistory',
    purpose: 'Section 8 Version/Date/Author/Changes table.',
    example: [
      { version: 'v0.1', date: 'Aug 31, 2026', author: 'Jane Doe', changes: 'Initial draft' },
    ],
  },
];

class UiUxSDK {
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
  setOverviewBullets(items) {
    this.data.overviewBullets = items;
    return this;
  }
  setUserFlowNote(t) {
    this.data.userFlowNote = t;
    return this;
  }
  setColorPalette(rows) {
    this.data.colorPalette = rows;
    return this;
  }
  setTypography(rows) {
    this.data.typography = rows;
    return this;
  }
  setComponents(rows) {
    this.data.components = rows;
    return this;
  }
  setSpacingGridNote(t) {
    this.data.spacingGridNote = t;
    return this;
  }
  setComponentBehaviour(items) {
    this.data.componentBehaviour = items;
    return this;
  }
  setEdgeCases(items) {
    this.data.edgeCases = items;
    return this;
  }
  setAnimationNote(t) {
    this.data.animationNote = t;
    return this;
  }
  setContentCopy(rows) {
    this.data.contentCopy = rows;
    return this;
  }
  setAccessibilityHeading(t) {
    this.data.accessibilityHeading = t;
    return this;
  }
  setAccessibilityBullets(items) {
    this.data.accessibilityBullets = items;
    return this;
  }
  setResponsiveNotes(rows) {
    this.data.responsiveNotes = rows;
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
        children: [headingRun('UI/UX Documentation')],
      }),
    ];
  }

  _buildContent() {
    const children = [];
    children.push(h1('UI/UX Documentation'));
    const meta = Object.assign({}, DEFAULT_METADATA, this.data.metadata || {});
    children.push(this._metadataTable(meta));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));

    children.push(h2('1. Overview'));
    children.push(
      guidanceNote(
        'Summarise the experience in plain language before diving into screens. What is the user trying to accomplish?',
      ),
    );
    (this.data.overviewBullets && this.data.overviewBullets.length
      ? this.data.overviewBullets
      : DEFAULT_OVERVIEW_BULLETS
    ).forEach((o) => children.push(bullet(o)));

    children.push(h2('2. User Flow'));
    children.push(
      guidanceNote(
        "Link the interactive prototype rather than pasting static screenshots — specs go stale fast if images aren't kept in sync.",
      ),
    );
    children.push(bodyPara(this.data.userFlowNote || DEFAULT_USER_FLOW_NOTE));

    children.push(h2('3. Design System / UI Kit'));
    children.push(h3('3.1 Colour Palette'));
    children.push(
      buildZebraTable(
        [
          { key: 'role', header: 'Role', weight: 1 },
          { key: 'hex', header: 'Hex', weight: 1 },
          { key: 'usage', header: 'Usage', weight: 1.6 },
        ],
        this.data.colorPalette,
        DEFAULT_COLOR_PALETTE,
      ),
    );
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('3.2 Typography'));
    children.push(
      buildZebraTable(
        [
          { key: 'style', header: 'Style', weight: 1 },
          { key: 'spec', header: 'Font / Size / Weight', weight: 1.4 },
          { key: 'usage', header: 'Usage', weight: 1.2 },
        ],
        this.data.typography,
        DEFAULT_TYPOGRAPHY,
      ),
    );
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('3.3 Components'));
    children.push(
      guidanceNote(
        'If using a component library (Figma AI, shadcn, Material), reference the source library instead of redefining from scratch.',
      ),
    );
    children.push(
      buildZebraTable(
        [
          { key: 'component', header: 'Component', weight: 1.1 },
          { key: 'states', header: 'States', weight: 1.6 },
          { key: 'notes', header: 'Notes / Source', weight: 1.1 },
        ],
        this.data.components,
        DEFAULT_COMPONENTS,
      ),
    );
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('3.4 Spacing & Grid'));
    children.push(bodyPara(this.data.spacingGridNote || DEFAULT_SPACING_GRID_NOTE));

    children.push(h2('4. Interaction Details'));
    children.push(h3('4.1 Component Behaviour'));
    (this.data.componentBehaviour && this.data.componentBehaviour.length
      ? this.data.componentBehaviour
      : DEFAULT_COMPONENT_BEHAVIOUR
    ).forEach((c) => children.push(bullet(c)));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('4.2 Edge Cases'));
    (this.data.edgeCases && this.data.edgeCases.length
      ? this.data.edgeCases
      : DEFAULT_EDGE_CASES
    ).forEach((e) => children.push(bullet(e)));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
    children.push(h3('4.3 Animation / Motion'));
    children.push(bodyPara(this.data.animationNote || DEFAULT_ANIMATION_NOTE));

    children.push(h2('5. Content & Copy'));
    children.push(
      buildZebraTable(
        [
          { key: 'element', header: 'Element', weight: 1 },
          { key: 'copy', header: 'Copy', weight: 1.4 },
          { key: 'notes', header: 'Notes', weight: 1.4 },
        ],
        this.data.contentCopy,
        DEFAULT_CONTENT_COPY,
      ),
    );

    // Section 6 is rendered as a Heading2 with the (unusual) full-sentence
    // heading exactly as it appears in the source template.
    children.push(h2(this.data.accessibilityHeading || DEFAULT_ACCESSIBILITY_HEADING));
    (this.data.accessibilityBullets && this.data.accessibilityBullets.length
      ? this.data.accessibilityBullets
      : DEFAULT_ACCESSIBILITY_BULLETS
    ).forEach((a) => children.push(bullet(a)));

    children.push(h2('7. Responsive & Platform Notes'));
    children.push(
      buildZebraTable(
        [
          { key: 'platform', header: 'Platform', weight: 1 },
          { key: 'notes', header: 'Notes', weight: 2.5 },
        ],
        this.data.responsiveNotes,
        DEFAULT_RESPONSIVE_NOTES,
      ),
    );

    children.push(h2('8. Revision History'));
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
            // Source has an extra space between "UI/UX" and "Documentation" in the header.
            run('UI/UX  Documentation', {
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

module.exports = UiUxSDK;
module.exports.SECTION_GUIDE = SECTION_GUIDE;
module.exports.buildZebraTable = buildZebraTable;
module.exports.BULLET_NUMBERING_REF = BULLET_NUMBERING_REF;
