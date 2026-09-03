'use strict';
/**
 * =============================================================================
 *  Generates llms/docx-sdk.md from the SDKs themselves.
 * =============================================================================
 * The previous agent-facing docs were a hand-written snapshot, so when the docx
 * package was replaced they silently became instructions for an API that no
 * longer exists. Every setter, its purpose and its payload shape already live
 * in each Part's `sectionGuide()`, so this reference is derived from that
 * rather than transcribed — it cannot drift.
 *
 *   node scripts/generate-llms-docs.js           # write llms/docx-sdk.md
 *   node scripts/generate-llms-docs.js --check   # fail if it is out of date
 *
 * Prose that a machine can't derive — what a Part is *for*, the authoring
 * order, the placeholder rule — lives in PART_NOTES below, next to the
 * generator, so there is exactly one place to edit when a Part changes shape.
 * =============================================================================
 */

const fs = require('fs');
const path = require('path');
const { version } = require('pspt-docx/package.json');
const { PART_TYPES, setterRegistry } = require('pspt-lang/src/parts');

const OUT = path.join(__dirname, '..', 'llms', 'docx-sdk.md');

/** One line per Part explaining what it is for — the only hand-written part. */
const PART_NOTES = {
  picMatrix:
    'Front matter for the whole document: the cover identity table, the team and their RACI-style roles, the PIC matrix saying who makes/checks/approves each Part, and the storage guide. Its `picMatrix` and `storageGuide` tables ship with sensible defaults — leave them alone unless your organisation genuinely differs.',
  styleGuide:
    "Documents the template's own design system. Usually correct as-is; override only if your product deliberately uses a different visual language.",
  projectBrief:
    'The elevator pitch: what is being built, the pains that justify it, measurable objectives, the module list that seeds the BRD and PRD, a rough phase timeline, deliverables and early risks.',
  brd: 'Business Requirements. The largest Part. Business-language requirements (BR-xxx) each traced to an objective, the process and its RACI, actors, use cases and their detailed specs, and the traceability matrix the SRS reads back.',
  prd: 'Product Requirements. Personas, MoSCoW-prioritised features (which the SRS turns into testable FR-xxx), user stories, the single canonical success-metrics table and the single Risk Register.',
  srs: 'Software Requirements. Atomic, testable functional requirements traced back to BRD ids, non-functional targets, interfaces, and the screen inventory.',
  techDoc:
    'Three mini-documents in one Part, generated back to back: **Technical Documentation** (architecture, components, stack, repo tree, data flow, integrations, security), **Data Model** (entity list, per-entity columns, relationships, indexing) and **API Specification** (auth, response envelope, error codes, endpoints, rate limiting, changelog).',
  uiux: 'Colour palette, typography, components and states, spacing, interaction, accessibility and responsive behaviour. Reference a component library rather than redefining one.',
  uat: 'Objectives, features in scope, test levels and owners, environments, the test-case table traced to BRD use cases, non-functional criteria, severity definitions and sign-off.',
  deploymentGuide:
    'What is deployed where, prerequisites, environments, the `.env` example, the deploy and rollback procedures, every configuration variable, the post-deploy checklist and monitoring.',
  userManual:
    'Written for the end user in plain language: getting started, navigation, one walkthrough per feature, FAQ, troubleshooting and support contact.',
  changelog:
    'Semantic-versioned release history with tagged entries (ADDED / CHANGED / FIXED / DEPRECATED / REMOVED / SECURITY), newest release first, plus an Unreleased section.',
  changeRequestLog:
    'One row per change request as they arrive: id, date, requester, description, affected flow and status.',
  glossary:
    'One consolidated glossary for the whole document — business terms, acronyms and technical jargon.',
  appendix:
    'Links out to living resources (tracker, design files, OpenAPI, runbooks, dashboards) instead of duplicating them here.',
};

/** How each inferred payload shape is written in a `.pspt` file. */
const DSL_FORM = {
  scalar: (k) => `\`${k}: "..."\``,
  stringList: (k) => `\`list ${k} { item "..." }\``,
  rows: (k) => `\`table ${k} { ... }\` + \`rows ${k} [ ... ]\``,
  nestedRows: (k) => `\`list ${k} { item "..." { item "..." } }\``,
  object: (k) => `\`object ${k} { key: value }\``,
  unknown: () => '—',
};

/** A compact, single-line rendering of an example payload. */
// Deep enough to reveal a nested field's own object keys (e.g. `columns` in
// setEntityDetails' `{tableName, columns: [{column, type, ...}]}` example) —
// this used to collapse at depth 1, so docx-sdk.md said `columns: […]` and
// never told a reader those items had to be objects, not strings.
const MAX_SHAPE_DEPTH = 3;

function shapeOf(example) {
  const render = (value, depth) => {
    if (value === null || value === undefined) return 'null';
    if (Array.isArray(value)) {
      if (!value.length) return '[]';
      return depth > MAX_SHAPE_DEPTH ? '[…]' : `${render(value[0], depth + 1)}[]`;
    }
    if (typeof value === 'object') {
      if (depth > MAX_SHAPE_DEPTH) return '{…}';
      const keys = Object.keys(value);
      const rendered = keys.map((key) => {
        const nested = value[key];
        const isNested = Array.isArray(nested) || (nested && typeof nested === 'object');
        return isNested ? `${key}: ${render(nested, depth + 1)}` : key;
      });
      return `{ ${rendered.join(', ')} }`;
    }
    return typeof value;
  };
  return render(example, 0);
}

function escapePipes(text) {
  return String(text).replace(/\|/g, '\\|');
}

function build() {
  const lines = [];
  const parts = [...PART_TYPES.values()];
  const totalSetters = parts.reduce((total, meta) => total + setterRegistry(meta.SDK).size, 0);

  lines.push('<!-- GENERATED FILE — do not edit by hand.');
  lines.push("     Produced by `npm run llms:build` from each Part SDK's own sectionGuide().");
  lines.push('     Edit scripts/generate-llms-docs.js (or the SDK) and re-run instead. -->');
  lines.push('');
  lines.push('# pspt-docx');
  lines.push('');
  lines.push(`**This reference describes pspt-docx ${version}.** An agent that installed`);
  lines.push('`pspt-docx@^1` has the removed `ProductSpecSDK` instead — check with');
  lines.push("`require('pspt-docx/package.json').version` before relying on anything below.");
  lines.push('');
  lines.push('`pspt-docx` generates the **Product Documentation Master** — a Word (.docx) suite');
  lines.push('made of PIC Matrix front matter followed by 14 numbered Parts. Each Part is a');
  lines.push('self-contained mini-document with its own title page and running header, and');
  lines.push('can be generated on its own or assembled with every other Part into one file.');
  lines.push('');
  lines.push(
    `There are **${totalSetters} setters across ${parts.length} Part SDKs**. You never touch styling — fonts,`,
  );
  lines.push('colours, table styles, page setup and header/footer chrome are fixed by the');
  lines.push('template.');
  lines.push('');
  lines.push('> **Replaces `ProductSpecSDK`.** The earlier single-document class, and its');
  lines.push('> `setCoverPage` / `setExecutiveSummary` / `setEnableHardwareSection` API, no');
  lines.push('> longer exist. There is no hardware section — this suite targets software');
  lines.push('> products. If you were told otherwise by an older copy of these docs, that');
  lines.push('> copy is stale.');
  lines.push('');
  lines.push('## The two ways in');
  lines.push('');
  lines.push('```js');
  lines.push("const { ProjectBriefSDK, MasterDocument } = require('pspt-docx');");
  lines.push('');
  lines.push('// One Part on its own');
  lines.push('const doc = new ProjectBriefSDK();');
  lines.push("doc.setHeaderFooterLabels({ productNameLabel: 'Acme Widget' });");
  lines.push("doc.setOverview('...');");
  lines.push("await doc.generate('./02-project-brief.docx');");
  lines.push('');
  lines.push('// Or every Part assembled into one document');
  lines.push('const master = new MasterDocument();');
  lines.push(
    "master.setHeaderFooterLabels({ productNameLabel: 'Acme Widget' }); // all 15 at once",
  );
  lines.push("master.part('projectBrief').setOverview('...');");
  lines.push("master.part('brd').setPurpose('...');");
  lines.push("await master.generate('./product-documentation-master.docx');");
  lines.push('```');
  lines.push('');
  lines.push('Every setter returns `this`, so calls chain.');
  lines.push('');
  lines.push('## The placeholder rule');
  lines.push('');
  lines.push('**Skipping a setter is fine and is the intended behaviour.** The section still');
  lines.push("renders, using the template's own `[bracketed placeholder]` text, so a reviewer");
  lines.push('sees exactly what is still missing. `new AnySDK().generate()` with zero calls');
  lines.push("reproduces that Part's blank template exactly, and a `MasterDocument` with only");
  lines.push('two Parts filled in still produces a complete, clean document.');
  lines.push('');
  lines.push('Corollary for an agent: **never invent content to avoid an empty section.** If');
  lines.push('the codebase gives no signal for a field, leave the setter uncalled, or pass an');
  lines.push('explicit `[define …]` note. A placeholder is information; a fabricated value is');
  lines.push('a defect.');
  lines.push('');
  lines.push('## Payload shapes');
  lines.push('');
  lines.push('Every setter takes one of five shapes. The third column is how it is written in');
  lines.push('a `.pspt` file (see [dsl.md](./dsl.md)); calling the SDK directly, pass the');
  lines.push('JavaScript value shown in the second column.');
  lines.push('');
  lines.push('| Shape | JavaScript | `.pspt` form |');
  lines.push('| --- | --- | --- |');
  lines.push('| `scalar` | a string (or number/boolean) | `overview: "..."` |');
  lines.push('| `stringList` | `string[]` | `list objectives { item "..." }` |');
  lines.push(
    '| `rows` | `Object[]` — one object per table row | `table x { ... }` + `rows x [ ... ]` |',
  );
  lines.push('| `nestedRows` | `{ title, children? }[]` | `list x { item "a" { item "b" } }` |');
  lines.push('| `object` | a single plain object | `object metadata { writer: "..." }` |');
  lines.push('');
  lines.push('A row cell may itself be a list or an object — that is what makes');
  lines.push("`setEndpoints`' request/response bodies and `setEntityDetails`' column lists");
  lines.push('reachable. Extra keys in a row are ignored; a missing key renders as a blank');
  lines.push('cell rather than throwing.');
  lines.push('');
  lines.push('## Discovering the API at runtime');
  lines.push('');
  lines.push('Every class carries its own machine-readable guide — prefer it over trusting');
  lines.push('any documentation, including this file:');
  lines.push('');
  lines.push('```js');
  lines.push('ProjectBriefSDK.sectionGuide();');
  lines.push("// [{ method: 'setOverview', purpose: '...', example: '...' }, ...]");
  lines.push('');
  lines.push("const { MasterDocument } = require('pspt-docx');");
  lines.push('MasterDocument.parts(); // [{ key, part, title, heading }, ...] in document order');
  lines.push('```');
  lines.push('');
  lines.push('## The Parts');
  lines.push('');
  lines.push('| Key | | Class | Setters |');
  lines.push('| --- | --- | --- | --- |');
  for (const meta of parts) {
    const count = setterRegistry(meta.SDK).size;
    const label = meta.part == null ? 'front matter' : `Part ${meta.part}`;
    lines.push(`| \`${meta.key}\` | ${label} | \`${meta.className}\` | ${count} |`);
  }
  lines.push('');
  lines.push('Order matters: `MasterDocument` assembles them exactly as listed. `key` is what');
  lines.push("you pass to `master.part(key)` and what a `.pspt` file's `type=` or `part` block");
  lines.push('names.');
  lines.push('');
  lines.push(...setterReference(parts));
  lines.push(...trailer());
  return lines.join('\n') + '\n';
}

/** The per-Part setter tables, generated from each SDK's sectionGuide(). */
function setterReference(parts) {
  const lines = [];
  lines.push('');
  lines.push('## Setter reference');
  lines.push('');
  lines.push('Every setter of every Part, generated from `sectionGuide()`. `Data key` is the');
  lines.push('name used in a `.pspt` file — the method name minus `set`, lowercased at the');
  lines.push('front.');

  for (const meta of parts) {
    const registry = setterRegistry(meta.SDK);
    const label = meta.part == null ? 'Front matter' : `Part ${meta.part}`;
    lines.push('');
    lines.push(`### ${label} — ${meta.title} (\`${meta.key}\`)`);
    lines.push('');
    lines.push(`\`${meta.className}\` · ${registry.size} setters`);
    lines.push('');
    if (PART_NOTES[meta.key]) {
      lines.push(PART_NOTES[meta.key]);
      lines.push('');
    }
    lines.push('| Setter | Data key | Shape | Payload | Purpose |');
    lines.push('| --- | --- | --- | --- | --- |');
    for (const entry of meta.SDK.sectionGuide()) {
      const dataKey = entry.method
        .replace(/^set/, '')
        .replace(/^./, (first) => first.toLowerCase());
      const reg = registry.get(dataKey);
      const shape = reg ? reg.shape : 'unknown';
      lines.push(
        `| \`${entry.method}\` | \`${dataKey}\` | ${shape} | \`${escapePipes(shapeOf(entry.example))}\` | ${escapePipes(entry.purpose || '')} |`,
      );
    }
    lines.push('');
    lines.push('`.pspt` forms for this Part:');
    lines.push('');
    const byShape = new Map();
    for (const [key, reg] of registry) {
      if (!byShape.has(reg.shape)) byShape.set(reg.shape, []);
      byShape.get(reg.shape).push(key);
    }
    for (const [shape, keys] of [...byShape].sort()) {
      const names = keys.map((key) => '`' + key + '`').join(', ');
      lines.push(`- **${shape}** — ${DSL_FORM[shape](keys[0])} · ${names}`);
    }
  }

  return lines;
}

/** The closing sections, which need no data from the SDKs. */
function trailer() {
  const lines = [];
  lines.push('');
  lines.push('## Output');
  lines.push('');
  lines.push('```js');
  lines.push("await doc.generate('/path/to/output.docx'); // writes to disk");
  lines.push('const buf = await doc.toBuffer(); // or get a Buffer directly');
  lines.push('```');
  lines.push('');
  lines.push('`MasterDocument` exposes the same two methods. Neither creates the output');
  lines.push('directory — it must already exist.');
  lines.push('');
  lines.push('## Images and diagrams');
  lines.push('');
  lines.push('This suite embeds no images. Fields such as `erdDiagramNote` and');
  lines.push('`gettingStartedScreenshot` are **caption strings**, not file paths — pass the');
  lines.push('caption or link you want printed (`"[See erd.png]"`). Diagrams are authored');
  lines.push('externally and referenced.');
  return lines;
}

const generated = build();
const check = process.argv.includes('--check');

if (check) {
  const current = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';
  if (current !== generated) {
    console.error('llms/docx-sdk.md is out of date with the SDKs. Run `npm run llms:build`.');
    process.exit(1);
  }
  console.log('llms/docx-sdk.md is up to date with the SDKs.');
} else {
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, generated);
  console.log(`wrote ${path.relative(process.cwd(), OUT)} (${generated.split('\n').length} lines)`);
}
