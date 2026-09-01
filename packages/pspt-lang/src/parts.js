'use strict';
/**
 * The bridge between the `.pspt` DSL and the pspt-docx SDK suite.
 *
 * Nothing here is hand-maintained. Every Part's setter list, its data-key
 * names, and the payload shape each setter expects are read straight out of
 * that SDK's own `sectionGuide()` at compile time — so adding a setter to a
 * Part makes it available in the DSL immediately, and removing one turns into
 * a compile-time warning rather than a runtime crash. (The previous codegen
 * carried a hand-written 50-entry map, which is exactly what went stale when
 * the docx package was replaced.)
 */

const docx = require('pspt-docx');

/** `setBackgroundPains` -> `backgroundPains` */
function methodToDataKey(method) {
  const bare = method.replace(/^set/, '');
  return bare.charAt(0).toLowerCase() + bare.slice(1);
}

/**
 * Infers what the DSL construct for a setter has to look like, from the
 * example payload the SDK publishes for it.
 *
 *   'scalar'     a plain `field: value` line
 *   'stringList' a `list <name> { item "..." }` block
 *   'rows'       a `table <name> { ... }` + `rows <name> [ ... ]` pair
 *   'nestedRows' rows whose entries carry nested `item` children
 *   'object'     an `object <name> { key: value }` block
 */
function inferShape(example) {
  if (example == null) return 'unknown';
  const t = typeof example;
  if (t === 'string' || t === 'number' || t === 'boolean') return 'scalar';
  if (Array.isArray(example)) {
    if (example.length === 0) return 'rows';
    if (example.every((e) => typeof e === 'string')) return 'stringList';
    if (example.some((e) => e && typeof e === 'object' && 'children' in e)) return 'nestedRows';
    return 'rows';
  }
  if (t === 'object') return 'object';
  return 'unknown';
}

/**
 * Builds `{ dataKey: { method, shape, example } }` for one SDK class.
 * @param {Function} SDK - a class exposing `static sectionGuide()`
 */
function setterRegistry(SDK) {
  const registry = new Map();
  for (const entry of SDK.sectionGuide()) {
    if (typeof entry.method !== 'string') continue;
    registry.set(methodToDataKey(entry.method), {
      method: entry.method,
      shape: inferShape(entry.example),
      example: entry.example,
    });
  }
  return registry;
}

/**
 * Every valid `type=` value for a docx `.pspt` file: one per Part, plus
 * `master` for the assembled document.
 * @returns {Map<string, {key, part, title, className, SDK}>}
 */
function buildPartTypes() {
  const byKey = new Map();
  const CLASS_BY_KEY = {
    picMatrix: 'PicMatrixSDK',
    styleGuide: 'StyleGuideSDK',
    projectBrief: 'ProjectBriefSDK',
    brd: 'BrdSDK',
    prd: 'PrdSDK',
    srs: 'SrsSDK',
    techDoc: 'TechnicalDocumentationSDK',
    uiux: 'UiUxSDK',
    uat: 'UatSDK',
    deploymentGuide: 'DeploymentGuideSDK',
    userManual: 'UserManualSDK',
    changelog: 'ChangelogSDK',
    changeRequestLog: 'ChangeRequestLogSDK',
    glossary: 'GlossarySDK',
    appendix: 'AppendixSDK',
  };
  for (const meta of docx.PARTS) {
    const className = CLASS_BY_KEY[meta.key];
    if (!className || !docx[className]) {
      throw new Error(
        `pspt-docx exports no class for Part "${meta.key}" — the DSL bridge is out of date`,
      );
    }
    byKey.set(meta.key, { ...meta, className, SDK: docx[className] });
  }
  return byKey;
}

const PART_TYPES = buildPartTypes();

/** All accepted `type=` values, in the order they appear in the master document. */
const DOC_TYPES = ['master', ...PART_TYPES.keys(), 'xlsx'];

/** Levenshtein-ish suggestion for an unrecognised name, or null. */
function suggest(name, candidates) {
  const lower = String(name).toLowerCase();
  let best = null;
  let bestScore = Infinity;
  for (const c of candidates) {
    const a = c.toLowerCase();
    if (a === lower) return c;
    const rows = [];
    for (let i = 0; i <= lower.length; i++) rows[i] = [i];
    for (let j = 0; j <= a.length; j++) rows[0][j] = j;
    for (let i = 1; i <= lower.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        rows[i][j] = Math.min(
          rows[i - 1][j] + 1,
          rows[i][j - 1] + 1,
          rows[i - 1][j - 1] + (lower[i - 1] === a[j - 1] ? 0 : 1),
        );
      }
    }
    const score = rows[lower.length][a.length];
    if (score < bestScore) {
      bestScore = score;
      best = c;
    }
  }
  return bestScore <= Math.max(2, Math.floor(lower.length / 3)) ? best : null;
}

module.exports = { PART_TYPES, DOC_TYPES, setterRegistry, methodToDataKey, inferShape, suggest };
