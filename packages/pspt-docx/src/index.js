'use strict';
/**
 * =============================================================================
 *  pspt-docx — Product Documentation Master SDK suite
 * =============================================================================
 * The package entry point. Every Part's SDK class, the shared style tokens and
 * primitives, and the single-document assembler, in one place:
 *
 *   const { ProjectBriefSDK } = require('pspt-docx');
 *   const doc = new ProjectBriefSDK();
 *   doc.setHeaderFooterLabels({ productNameLabel: 'Acme Widget' });
 *   doc.setOverview('...');
 *   await doc.generate('./02-project-brief.docx');
 *
 * Or assemble the whole 87-page master in one run:
 *
 *   const { MasterDocument } = require('pspt-docx');
 *   const master = new MasterDocument();
 *   master.setHeaderFooterLabels({ productNameLabel: 'Acme Widget' });
 *   master.part('projectBrief').setOverview('...');
 *   await master.generate('./product-documentation-master.docx');
 *
 * The modules under `src/sdk/` stay individually requirable and self-contained,
 * so the folder can still be copied out and used on its own — see src/README.md.
 * =============================================================================
 */

const { MasterDocument, PARTS, masterStyles } = require('./sdk/generate-master');
const core = require('./sdk/pspt-core');

module.exports = {
  // Front matter
  PicMatrixSDK: require('./sdk/pic-matrix-docx'),

  // Parts 1-14
  StyleGuideSDK: require('./sdk/style-guide-docx'),
  ProjectBriefSDK: require('./sdk/project-brief-docx'),
  BrdSDK: require('./sdk/brd-docx'),
  PrdSDK: require('./sdk/prd-docx'),
  SrsSDK: require('./sdk/srs-docx'),
  TechnicalDocumentationSDK: require('./sdk/tech-doc-docx'),
  UiUxSDK: require('./sdk/uiux-docx'),
  UatSDK: require('./sdk/uat-docx'),
  DeploymentGuideSDK: require('./sdk/deployment-guide-docx'),
  UserManualSDK: require('./sdk/user-manual-docx'),
  ChangelogSDK: require('./sdk/changelog-docx'),
  ChangeRequestLogSDK: require('./sdk/change-request-log-docx'),
  GlossarySDK: require('./sdk/glossary-docx'),
  AppendixSDK: require('./sdk/appendix-docx'),

  // Whole-document assembly
  MasterDocument,
  PARTS,
  masterStyles,

  // Shared style tokens and docx primitives
  core,
};
