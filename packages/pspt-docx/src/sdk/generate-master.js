'use strict';
/**
 * =============================================================================
 *  PRODUCT DOCUMENTATION MASTER — single-document assembler
 * =============================================================================
 * Builds the whole 87-page master document in one run instead of generating
 * 15 separate .docx files and merging them by hand.
 *
 * It does NOT merge packed .docx files. Every Part is built by the same `docx`
 * library with the same style ids, so this concatenates each Part's *sections*
 * into one Document — no OOXML surgery, no relationship/numbering remapping,
 * and Word's "Page X of Y" fields resolve against the real total.
 *
 * HOW TO USE
 * -----------
 *   const { MasterDocument } = require('./generate-master');
 *
 *   const master = new MasterDocument();
 *   master.setHeaderFooterLabels({ productNameLabel: 'Acme Widget' }); // all Parts
 *
 *   master.part('projectBrief')
 *     .setOverview('Acme Widget lets small teams submit expenses from their phones.')
 *     .setObjectives(['Cut expense processing time by 70%']);
 *
 *   master.part('brd').setPurpose('...');
 *
 *   await master.generate('./product-documentation-master.docx');
 *
 * Parts you never touch still render — as the template's own placeholder text,
 * exactly as they do standalone. `new MasterDocument().generate(path)` gives
 * you the complete blank master.
 * =============================================================================
 */

const { Document, Packer, AlignmentType } = require('docx');
const fs = require('fs');

const { DOCX_FONT: FONT, DOCX_COLOR: COLOR, DOCX_SIZE: SIZE } = require('./pspt-core');

/**
 * Every Part, in the exact order the master document assembles them.
 * `key` is what you pass to `master.part(key)`; `heading` is the literal heading
 * text that Part renders, so a caller can build a table of contents from this list.
 */
const PARTS = [
  {
    key: 'picMatrix',
    part: null,
    title: 'PIC Matrix & Documentation SOP',
    heading: 'Name of Product Documentation Flow & PIC Matrix',
    module: './pic-matrix-docx',
  },
  {
    key: 'styleGuide',
    part: 1,
    title: 'Documentation Style Guide',
    heading: 'Documentation Style Guide',
    module: './style-guide-docx',
  },
  {
    key: 'projectBrief',
    part: 2,
    title: 'Project Brief',
    heading: 'Project Brief',
    module: './project-brief-docx',
  },
  {
    key: 'brd',
    part: 3,
    title: 'Business Requirements Document',
    heading: 'Business Requirements Document (BRD)',
    module: './brd-docx',
  },
  {
    key: 'prd',
    part: 4,
    title: 'Product Requirements Document',
    heading: 'Product Requirements Document (PRD)',
    module: './prd-docx',
  },
  {
    key: 'srs',
    part: 5,
    title: 'Software Requirements Specification',
    heading: 'Software Requirements Specification (SRS)',
    module: './srs-docx',
  },
  {
    key: 'techDoc',
    part: 6,
    title: 'Technical Documentation, Data Model & API Spec',
    heading: 'Technical Documentation',
    module: './tech-doc-docx',
  },
  {
    key: 'uiux',
    part: 7,
    title: 'UI/UX Documentation',
    heading: 'UI/UX Documentation',
    module: './uiux-docx',
  },
  {
    key: 'uat',
    part: 8,
    title: 'User Acceptance Testing',
    heading: 'User Acceptance Testing(UAT)',
    module: './uat-docx',
  },
  {
    key: 'deploymentGuide',
    part: 9,
    title: 'Deployment Guide',
    heading: 'Deployment Guide',
    module: './deployment-guide-docx',
  },
  {
    key: 'userManual',
    part: 10,
    title: 'User Manual',
    heading: 'User Manual',
    module: './user-manual-docx',
  },
  {
    key: 'changelog',
    part: 11,
    title: 'Changelog',
    heading: 'Changelog',
    module: './changelog-docx',
  },
  {
    key: 'changeRequestLog',
    part: 12,
    title: 'Change Request Log',
    heading: 'Change Request Log',
    module: './change-request-log-docx',
  },
  { key: 'glossary', part: 13, title: 'Glossary', heading: 'Glossary', module: './glossary-docx' },
  { key: 'appendix', part: 14, title: 'Appendix', heading: 'Appendix', module: './appendix-docx' },
];

/**
 * The master document's paragraph styles. Fourteen of the fifteen Parts already
 * define exactly this set standalone; the PIC Matrix shifts each level up by one
 * because it has no Part-title page above it. That divergence is cosmetic only —
 * its single heading writes its own run size and color directly, so it renders
 * identically under either style set.
 */
function masterStyles() {
  return {
    default: { document: { run: { font: FONT, size: SIZE.body } } },
    paragraphStyles: [
      {
        id: 'Title',
        name: 'Title',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: { font: FONT, size: SIZE.title, bold: true, color: COLOR.h1 },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 200 } },
      },
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
      {
        id: 'Heading4',
        name: 'Heading 4',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: { font: FONT, size: SIZE.h4, bold: true, color: COLOR.h1Rendered },
        paragraph: { spacing: { before: 200, after: 100 } },
      },
    ],
  };
}

class MasterDocument {
  /**
   * @param {Object} [opts]
   * @param {string[]} [opts.include] - only assemble these Part keys, in PARTS order.
   *   Defaults to all fifteen.
   */
  constructor(opts = {}) {
    const include = opts.include;
    if (include) {
      const unknown = include.filter((k) => !PARTS.some((p) => p.key === k));
      if (unknown.length) throw new Error(`Unknown Part key(s): ${unknown.join(', ')}`);
    }
    this.order = PARTS.filter((p) => !include || include.includes(p.key));
    this.instances = new Map();
    this.sharedLabels = null;
  }

  /** The list of Part keys this master will assemble, in order. */
  static parts() {
    return PARTS.map((p) => ({ key: p.key, part: p.part, title: p.title, heading: p.heading }));
  }

  /**
   * The SDK instance for one Part, created on first access. Call its `set...()`
   * methods to fill it in; anything left unset renders as template placeholder text.
   * @param {string} key - a key from `MasterDocument.parts()`
   */
  part(key) {
    const meta = this.order.find((p) => p.key === key);
    if (!meta) throw new Error(`Part "${key}" is not part of this master document`);
    if (!this.instances.has(key)) {
      const SDK = require(meta.module);
      const instance = new SDK();
      if (this.sharedLabels) instance.setHeaderFooterLabels(this.sharedLabels);
      this.instances.set(key, instance);
    }
    return this.instances.get(key);
  }

  /**
   * Applies the running-header labels to every Part at once — the one call the
   * suite otherwise makes you repeat fifteen times.
   * @param {Object} labels - e.g. `{ productNameLabel: 'Acme Widget' }`
   */
  setHeaderFooterLabels(labels = {}) {
    this.sharedLabels = labels;
    for (const instance of this.instances.values()) instance.setHeaderFooterLabels(labels);
    return this;
  }

  documentOptions() {
    const sections = [];
    const numbering = [];
    const seenRefs = new Set();

    for (const meta of this.order) {
      const options = this.part(meta.key).documentOptions();
      sections.push(...options.sections);
      for (const config of (options.numbering && options.numbering.config) || []) {
        // Every Part namespaces its own reference, so a collision would mean a
        // genuine authoring bug rather than something to silently paper over.
        if (seenRefs.has(config.reference)) {
          throw new Error(
            `Duplicate numbering reference "${config.reference}" — two Parts share a bullet-list id`,
          );
        }
        seenRefs.add(config.reference);
        numbering.push(config);
      }
    }

    return { styles: masterStyles(), numbering: { config: numbering }, sections };
  }

  toDocument() {
    return new Document(this.documentOptions());
  }

  async toBuffer() {
    return Packer.toBuffer(this.toDocument());
  }

  /** Renders the assembled master document to `outputPath`. */
  async generate(outputPath) {
    fs.writeFileSync(outputPath, await this.toBuffer());
    return outputPath;
  }
}

module.exports = { MasterDocument, PARTS, masterStyles };
