'use strict';
/**
 * =============================================================================
 *  Checks that the agent-facing docs are true of the code.
 * =============================================================================
 * `llms:check` proves the docs are in sync with each other and freshly
 * generated. That is not the same as being *correct*. These docs are served
 * publicly and are the only input some agents get, so this asserts the claims
 * they make actually hold:
 *
 *   - every `type=` the DSL accepts is documented, and no removed one is
 *     advertised outside a deprecation notice
 *   - every setter named in the reference exists on its SDK
 *   - every Part key used in an example is real
 *   - every internal link resolves, with the right prefix for where it is served
 *   - every referenced example file exists
 *   - every `.pspt` snippet in the docs actually compiles
 *   - every SDK identifier used in a JS snippet actually exists
 *
 *   node scripts/audit-llms-docs.js
 * =============================================================================
 */

const fs = require('fs');
const path = require('path');

const { compile } = require('pspt-lang');
const { DOC_TYPES, PART_TYPES, setterRegistry } = require('pspt-lang/src/parts');
const docxPkg = require('pspt-docx');
const ExcelTrackerSDK = require('pspt-xlsx');

const DIR = path.join(__dirname, '..', 'llms');
const files = fs.readdirSync(DIR).filter((name) => name.endsWith('.md') || name.endsWith('.txt'));
const read = (name) => fs.readFileSync(path.join(DIR, name), 'utf8');

const passed = [];
const failed = [];
const check = (cond, label) => (cond ? passed : failed).push(label);

/** Names destructured out of `require('pspt-docx')` in a snippet. */
function destructuredImports(body) {
  const names = [];
  for (const [, group] of body.matchAll(/const \{([^}]+)\} = require\('pspt-docx'\)/g)) {
    for (const ident of group.split(',')) {
      const trimmed = ident.trim();
      if (trimmed) names.push(trimmed);
    }
  }
  return names;
}

/** Fenced code blocks, with the info string. */
function codeBlocks(text) {
  const out = [];
  const re = /```([a-z]*)\n([\s\S]*?)```/g;
  let m;
  while ((m = re.exec(text)) !== null) out.push({ lang: m[1], body: m[2] });
  return out;
}

// --- Doc types -------------------------------------------------------------
const dsl = read('dsl.md');
for (const type of DOC_TYPES)
  check(dsl.includes('`' + type + '`'), `dsl.md documents type=${type}`);

// A removed type may only appear inside a deprecation blockquote.
const unguarded = dsl
  .split('\n')
  .filter((line) => !line.trimStart().startsWith('>') && /`type=docx`|`ProductSpecSDK`/.test(line));
check(unguarded.length === 0, 'dsl.md mentions removed APIs only in deprecation notices');

// --- Counts ----------------------------------------------------------------
const totalSetters = [...PART_TYPES.values()].reduce(
  (sum, meta) => sum + setterRegistry(meta.SDK).size,
  0,
);
for (const name of ['llms.txt', 'docx-sdk.md']) {
  const text = read(name);
  const partsClaimed = [...text.matchAll(/(\d+)\s+Part SDKs/g)].map((match) => Number(match[1]));
  check(
    partsClaimed.every((n) => n === PART_TYPES.size),
    `${name}: "N Part SDKs" matches ${PART_TYPES.size}`,
  );
  const settersClaimed = [...text.matchAll(/\*\*(\d+) setters/g)].map((match) => Number(match[1]));
  check(
    settersClaimed.every((n) => n === totalSetters),
    `${name}: total setter count matches ${totalSetters}`,
  );
}

// --- Every documented setter exists ---------------------------------------
const rows = [...read('docx-sdk.md').matchAll(/^\| `(set[A-Za-z]+)` \| `([a-zA-Z]+)` \|/gm)];
check(rows.length === totalSetters, `docx-sdk.md documents all ${totalSetters} setters`);
const missingSetters = rows.filter(([, method, dataKey]) => {
  return ![...PART_TYPES.values()].some((meta) => {
    const registry = setterRegistry(meta.SDK);
    return (
      registry.has(dataKey) &&
      registry.get(dataKey).method === method &&
      typeof meta.SDK.prototype[method] === 'function'
    );
  });
});
check(
  missingSetters.length === 0,
  `every documented setter exists on its SDK${missingSetters.length ? ' — missing: ' + missingSetters[0][1] : ''}`,
);

// --- Part keys, links, referenced files ------------------------------------
const partKeys = [...PART_TYPES.keys()];
for (const name of files) {
  const text = read(name);

  const named = [...text.matchAll(/master\.part\('([a-zA-Z]+)'\)/g)].map((match) => match[1]);
  const badKeys = named.filter((key) => !partKeys.includes(key));
  check(
    badKeys.length === 0,
    `${name}: Part keys in examples are real${badKeys.length ? ' — ' + badKeys : ''}`,
  );

  const links = [...text.matchAll(/\]\(([^)\s]+)\)/g)]
    .map((match) => match[1])
    .filter((href) => href.startsWith('./'))
    .map((href) => href.slice(2).split('#')[0])
    .map((href) => (href.startsWith('llms/') ? href.slice('llms/'.length) : href))
    .filter((href) => href.endsWith('.md') || href.endsWith('.txt'));
  const dead = [...new Set(links)].filter((link) => !files.includes(link));
  check(
    dead.length === 0,
    `${name}: internal links resolve${dead.length ? ' — dead: ' + dead : ''}`,
  );

  const examples = [
    ...new Set([...text.matchAll(/examples\/[\w/.-]+\.pspt/g)].map((match) => match[0])),
  ];
  for (const rel of examples) {
    check(
      fs.existsSync(path.join(__dirname, '..', rel)),
      `${name}: referenced example exists — ${rel}`,
    );
  }
}

// llms.txt is served at the site root; the rest are served from inside /llms.
check(/\]\(\.\/llms\//.test(read('llms.txt')), 'llms.txt links use the ./llms/ prefix');
check(!/\]\(\.\/llms\//.test(dsl), 'sibling docs link without the llms/ prefix');

// --- Every .pspt snippet in the docs compiles ------------------------------
let snippets = 0;
for (const name of files) {
  for (const { body } of codeBlocks(read(name))) {
    const firstCode = body
      .split('\n')
      .map((line) => line.trim())
      .find((line) => line && !line.startsWith('//') && !line.startsWith('#'));
    if (!firstCode || !firstCode.startsWith('doc "')) continue;
    if (/<[a-z][a-z-]*>/i.test(body)) continue; // schematic, not compilable source
    snippets++;
    try {
      compile(body, name);
      check(true, `${name}: .pspt snippet #${snippets} compiles`);
    } catch (err) {
      check(false, `${name}: .pspt snippet #${snippets} FAILED — ${err.message}`);
    }
  }
}
check(snippets > 0, `found ${snippets} complete .pspt snippets to compile`);

// --- Every SDK identifier used in a JS snippet exists ----------------------
const exported = new Set(Object.keys(docxPkg));
for (const name of files) {
  for (const { lang, body } of codeBlocks(read(name))) {
    if (lang !== 'js') continue;

    for (const ident of destructuredImports(body)) {
      check(exported.has(ident), `${name}: pspt-docx exports ${ident}`);
    }

    for (const [, method] of body.matchAll(/\.(set[A-Za-z]+)\(/g)) {
      const onPart = [...PART_TYPES.values()].some(
        (meta) => typeof meta.SDK.prototype[method] === 'function',
      );
      const onTracker = typeof ExcelTrackerSDK.prototype[method] === 'function';
      check(onPart || onTracker, `${name}: ${method}() exists on a Part SDK or ExcelTrackerSDK`);
    }

    for (const [, statik] of body.matchAll(/(?:SDK|MasterDocument)\.(sectionGuide|parts)\(\)/g)) {
      const onMaster = typeof docxPkg.MasterDocument[statik] === 'function';
      const onPart = [...PART_TYPES.values()].some(
        (meta) => typeof meta.SDK[statik] === 'function',
      );
      check(onMaster || onPart, `${name}: static ${statik}() exists`);
    }
  }
}

const summary = `${passed.length} passed, ${failed.length} failed`;
if (failed.length) {
  console.error('Documentation claims that are not true of the code:\n');
  for (const label of failed) console.error('  FAIL  ' + label);
  console.error('\n' + summary);
  process.exit(1);
}
console.log(`Agent-facing docs check out against the code — ${summary}.`);
