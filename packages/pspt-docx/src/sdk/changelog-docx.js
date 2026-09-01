'use strict';
/**
 * =============================================================================
 *  CHANGELOG (Part 11 of the Product Documentation Master)
 * =============================================================================
 * Regenerates "Part 11. Changelog" exactly: title page, metadata table,
 * "1. How to Use This Changelog" bullets, "2. Category Legend" table,
 * "3. Unreleased" bullets, then a repeatable version-heading block per
 * release: "vX.Y.Z [YYYY-MM-DD] — Optional theme" as Heading 3, followed by
 * "TAG — description" bullets.
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

const BULLET_NUMBERING_REF = 'changelog-bullet-list';

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
const DEFAULT_HOW_TO_USE = [
  'Newest release goes at the top — read top to bottom, most recent first.',
  'One entry per user-visible change. Skip internal refactors unless they affect behaviour.',
  'Follow Semantic Versioning: MAJOR.MINOR.PATCH (breaking change → MAJOR; new feature → MINOR; bug fix → PATCH).',
  'Write for the reader of the release, not the author of the code.',
];
const DEFAULT_CATEGORY_LEGEND = [
  { tag: 'ADDED', useFor: 'New features or capabilities' },
  { tag: 'CHANGED', useFor: 'Changes to existing functionality' },
  { tag: 'FIXED', useFor: 'Bug fixes' },
  { tag: 'DEPRECATED', useFor: 'Features are still working but are scheduled for removal' },
  { tag: 'REMOVED', useFor: 'Features that were removed' },
  { tag: 'SECURITY', useFor: 'Security-related fixes or hardening' },
];
const DEFAULT_UNRELEASED = [
  'ADDED — [Description of new feature]',
  'FIXED — [Description of bug fix]',
];
const DEFAULT_RELEASES = [
  {
    heading: 'v1.2.0 [YYYY-MM-DD] — [Optional one-line theme]',
    entries: [
      'ADDED — [e.g. Added CSV export for reports]',
      'CHANGED — [e.g. Search results now load 40% faster]',
      'FIXED — [e.g. Fixed crash when uploading files over 50MB]',
      'SECURITY — [e.g. Patched dependency with known CVE]',
    ],
  },
  {
    heading: 'v1.1.0 [YYYY-MM-DD]',
    entries: [
      'ADDED — [...]',
      'DEPRECATED — [e.g. Legacy API v0 endpoints — removal planned v2.0]',
    ],
  },
  {
    heading: 'v1.0.0 [YYYY-MM-DD] — Initial release',
    entries: ['ADDED — Initial public release.'],
  },
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
    method: 'setUnreleased',
    purpose: 'Section 3 bulleted unreleased changes.',
    example: ['ADDED — Dark mode', 'FIXED — Crash on export'],
  },
  {
    method: 'setReleases',
    purpose:
      'Repeatable version blocks (newest first). Each has a heading and TAG — description bullets.',
    example: [
      {
        heading: 'v1.3.0 [2026-09-01] — Dark mode',
        entries: [
          'ADDED — Dark mode toggle in settings',
          'FIXED — Sidebar overflow on small screens',
        ],
      },
    ],
  },
];

class ChangelogSDK {
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
  setUnreleased(items) {
    this.data.unreleased = items;
    return this;
  }
  /** @param {Array<{heading:string, entries:string[]}>} items */
  setReleases(items) {
    this.data.releases = items;
    return this;
  }

  _buildPartTitlePage() {
    return [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [headingRun('Changelog')] }),
    ];
  }

  _buildContent() {
    const children = [];
    children.push(h1('Changelog'));
    const meta = Object.assign({}, DEFAULT_METADATA, this.data.metadata || {});
    children.push(this._metadataTable(meta));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));

    children.push(h2('1. How to Use This Changelog'));
    DEFAULT_HOW_TO_USE.forEach((h) => children.push(bullet(h)));

    children.push(h2('2. Category Legend'));
    children.push(
      buildZebraTable(
        [
          { key: 'tag', header: 'Tag', weight: 1 },
          { key: 'useFor', header: 'Use for', weight: 2.5 },
        ],
        DEFAULT_CATEGORY_LEGEND,
      ),
    );

    children.push(h2('3. Unreleased'));
    children.push(
      guidanceNote(
        'Changes merged but not yet shipped go here. Move them under a dated version heading at the time of release.',
      ),
    );
    (this.data.unreleased && this.data.unreleased.length
      ? this.data.unreleased
      : DEFAULT_UNRELEASED
    ).forEach((u) => children.push(bullet(u)));

    const releases =
      this.data.releases && this.data.releases.length ? this.data.releases : DEFAULT_RELEASES;
    releases.forEach((r) => {
      children.push(h3(r.heading));
      (r.entries || []).forEach((e) => children.push(bullet(e)));
    });

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
            run('Changelog', {
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

module.exports = ChangelogSDK;
module.exports.SECTION_GUIDE = SECTION_GUIDE;
module.exports.buildZebraTable = buildZebraTable;
module.exports.BULLET_NUMBERING_REF = BULLET_NUMBERING_REF;
