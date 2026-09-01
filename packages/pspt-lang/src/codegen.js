'use strict';
/**
 * Codegen: turns the AST from parser.js into plain, readable JS source that
 * constructs the right pspt-docx Part SDK (or `MasterDocument`, or the
 * pspt-xlsx `ExcelTrackerSDK`) and calls its existing setters. Codegen is
 * template-filling — it introduces no rendering logic of its own.
 *
 * docx mapping
 * ------------
 * `doc "Title" type=<partKey>` picks the Part, and therefore which SDK class
 * to instantiate and which setters exist. `type=master` instead builds a
 * `MasterDocument` and routes each `part <key> { ... }` block to
 * `master.part('<key>')`.
 *
 * Inside a Part, the construct you write has to match the payload shape the
 * setter expects — and that shape is read from the SDK's own `sectionGuide()`
 * rather than hardcoded here (see parts.js), so the mapping cannot drift out
 * of sync with the SDK the way the previous hand-written map did:
 *
 *   scalar     `overview: "..."`                        -> setOverview('...')
 *   stringList `list objectives { item "..." }`          -> setObjectives(['...'])
 *   rows       `table x {...}` + `rows x [ {...} ]`      -> setX([{...}])
 *   nestedRows `list x { item "a" { item "b" } }`        -> setX([{title:'a',children:['b']}])
 *   object     `object metadata { writer: "..." }`       -> setMetadata({writer:'...'})
 *
 * A row cell may itself be a [list] or an {object}, which is what makes the
 * setters with nested payloads — an endpoint's request/response body, an
 * entity's column list, a user story's Given/When/Then bullets — reachable.
 *
 * A construct that doesn't match, or a name with no setter behind it, becomes a
 * `// WARNING:` comment in the generated file rather than failing the compile —
 * so a half-written spec still produces a runnable script.
 *
 * xlsx mapping
 * ------------
 * `calendar <from> .. <to>` -> setCalendarRange(from, to).
 * `group "<title>" color=<c> { task "<name>" start=... end=... ... }` ->
 * addSection({title, bannerColor, rowColor, tasks: [...]});
 * `callout <date> "<text>" color=<c>` -> addCallout({dateStr, text, color}).
 */

const { PART_TYPES, setterRegistry, suggest } = require('./parts');

function jsLiteral(v) {
  return JSON.stringify(v, null, 2);
}

function valueNodeToJs(node, depth = 0) {
  const pad = '  '.repeat(depth + 1);
  const closePad = '  '.repeat(depth);
  switch (node.kind) {
    case 'string':
    case 'number':
    case 'date':
    case 'bool':
    case 'ident':
      return jsLiteral(node.value);
    case 'null':
      return 'null';
    case 'array': {
      if (!node.items.length) return '[]';
      const items = node.items.map((item) => `${pad}${valueNodeToJs(item, depth + 1)}`);
      return `[\n${items.join(',\n')}\n${closePad}]`;
    }
    case 'object': {
      if (!node.fields.length) return '{}';
      const fields = node.fields.map((f) => `${pad}${f.key}: ${valueNodeToJs(f.value, depth + 1)}`);
      return `{\n${fields.join(',\n')}\n${closePad}}`;
    }
    default:
      return 'null';
  }
}

/** Indent every line of a multi-line literal to sit under `pad`. */
function reindent(text, pad) {
  return text
    .split('\n')
    .map((line, i) => (i === 0 ? line : pad + line))
    .join('\n');
}

function rowsArrayLiteral(rowsNode) {
  if (!rowsNode || rowsNode.rows.length === 0) return '[]';
  const rows = rowsNode.rows.map((row) => {
    const fields = row.fields.map((f) => `    ${f.key}: ${valueNodeToJs(f.value, 2)}`);
    return `  {\n${fields.join(',\n')}\n  }`;
  });
  return `[\n${rows.join(',\n')}\n]`;
}

function objectLiteral(fields) {
  if (!fields.length) return '{}';
  return `{\n${fields.map((f) => `  ${f.key}: ${valueNodeToJs(f.value, 1)}`).join(',\n')}\n}`;
}

function stringListLiteral(items) {
  if (!items.length) return '[]';
  return `[\n${items.map((i) => `  ${jsLiteral(i.title)}`).join(',\n')}\n]`;
}

function nestedItemsLiteral(items) {
  const one = (item) => {
    if (item.children && item.children.length) {
      const kids = item.children.map((c) => `      ${jsLiteral(c.title)}`).join(',\n');
      return `  { title: ${jsLiteral(item.title)}, children: [\n${kids}\n    ] }`;
    }
    return `  { title: ${jsLiteral(item.title)} }`;
  };
  if (!items.length) return '[]';
  return `[\n${items.map(one).join(',\n')}\n]`;
}

/** The DSL construct that would satisfy a setter expecting `shape`. */
function constructFor(name, shape) {
  switch (shape) {
    case 'scalar':
      return `\`${name}: "..."\``;
    case 'stringList':
      return `\`list ${name} { item "..." }\``;
    case 'object':
      return `\`object ${name} { key: value }\``;
    default:
      return `\`table ${name} { ... }\` with a matching \`rows ${name} [ ... ]\``;
  }
}

/** "an object payload" / "a rows payload" / "an unrecognised payload". */
function describeShape(shape) {
  if (shape === 'unknown') return 'an unrecognised payload';
  return `${/^[aeiou]/i.test(shape) ? 'an' : 'a'} ${shape} payload`;
}

/**
 * Emits the setter calls for one Part's body.
 *
 * @param {Array} body - section-body nodes belonging to this Part
 * @param {Map} registry - dataKey -> {method, shape} for this Part's SDK
 * @param {string} receiver - the JS variable the calls hang off
 * @param {string} pad - leading indentation
 * @returns {string[]} generated lines
 */
function emitPartBody(body, registry, receiver, pad) {
  const lines = [];
  const known = [...registry.keys()];

  const warn = (message) => lines.push(`${pad}// WARNING: ${message}`);

  /** Reports a name that has no setter, with a spelling suggestion when close. */
  const noSetter = (kind, name, line) => {
    const hint = suggest(name, known);
    warn(
      `${kind} '${name}' (line ${line}) has no matching setter on this Part${hint ? ` — did you mean '${hint}'?` : ''}, skipped`,
    );
  };

  /** Reports a construct that doesn't match the setter's payload shape. */
  const wrongShape = (name, entry, used, line) => {
    warn(
      `'${name}' (line ${line}) maps to ${entry.method}, which expects ${describeShape(entry.shape)}, not ${used} — skipped. Use ${constructFor(name, entry.shape)}.`,
    );
  };

  const emit = (entry, literal) => {
    lines.push(`${pad}${receiver}.${entry.method}(${reindent(literal, pad)});`);
  };

  // Flatten nested `section` blocks; they are an authoring convenience only,
  // exactly as they were before.
  const flat = [];
  const walk = (nodes) => {
    for (const node of nodes) {
      if (node.type === 'Section') walk(node.body);
      else flat.push(node);
    }
  };
  walk(body);

  // Pair each `table` with the `rows` block that follows it, stopping at the
  // next `table` of the same name — so two sections can each declare a table
  // called `keyModules` without the second one's rows leaking into the first.
  const rowsForTable = new Map();
  const pairedRows = new Set();
  flat.forEach((node, idx) => {
    if (node.type !== 'Table') return;
    for (let j = idx + 1; j < flat.length; j++) {
      const other = flat[j];
      if (other.type === 'Table' && other.name === node.name) break;
      if (other.type === 'Rows' && other.name === node.name) {
        rowsForTable.set(node, other);
        pairedRows.add(other);
        break;
      }
    }
  });

  for (const node of flat) {
    if (node.type === 'Field') {
      const entry = registry.get(node.name);
      if (!entry) {
        noSetter('field', node.name, node.line);
        continue;
      }
      if (entry.shape !== 'scalar') {
        wrongShape(node.name, entry, 'a plain value', node.line);
        continue;
      }
      emit(entry, valueNodeToJs(node.value));
    } else if (node.type === 'Table') {
      const entry = registry.get(node.name);
      if (!entry) {
        noSetter('table', node.name, node.line);
        continue;
      }
      if (entry.shape !== 'rows' && entry.shape !== 'nestedRows') {
        wrongShape(node.name, entry, 'table rows', node.line);
        continue;
      }
      emit(entry, rowsArrayLiteral(rowsForTable.get(node)));
    } else if (node.type === 'ObjectBlock') {
      const entry = registry.get(node.name);
      if (!entry) {
        noSetter('object', node.name, node.line);
        continue;
      }
      if (entry.shape !== 'object') {
        wrongShape(node.name, entry, 'an object block', node.line);
        continue;
      }
      emit(entry, objectLiteral(node.fields));
    } else if (node.type === 'List') {
      const entry = registry.get(node.name);
      if (!entry) {
        noSetter('list', node.name, node.line);
        continue;
      }
      const nested = node.items.some((i) => i.children && i.children.length);
      if (entry.shape === 'stringList') {
        if (nested) {
          warn(
            `list '${node.name}' (line ${node.line}) maps to ${entry.method}, which takes a flat list of strings — nested items were dropped`,
          );
        }
        emit(entry, stringListLiteral(node.items));
      } else if (entry.shape === 'nestedRows') {
        emit(entry, nestedItemsLiteral(node.items));
      } else {
        wrongShape(node.name, entry, 'a list', node.line);
      }
    }
  }

  // A `rows` block with no matching `table` silently dropped its data before;
  // it is worth a louder note than an unknown-setter warning.
  for (const node of flat) {
    if (node.type === 'Rows' && !pairedRows.has(node)) {
      warn(
        `rows '${node.name}' (line ${node.line}) has no matching 'table ${node.name} { ... }' in this Part, skipped`,
      );
    }
  }

  return lines;
}

function header(lines, sourceFileLabel) {
  lines.push("'use strict';");
  lines.push(
    `// Generated by pspt-lang from ${sourceFileLabel}. Do not hand-edit — re-run \`pspt compile\` instead.`,
  );
}

function footer(lines, defaultOut) {
  lines.push('  return outputPath;');
  lines.push('}');
  lines.push('');
  lines.push('module.exports = { build };');
  lines.push('');
  lines.push('if (require.main === module) {');
  lines.push(`  const out = process.argv[2] || '${defaultOut}';`);
  lines.push(
    "  build(out).then(() => console.log('wrote ' + out)).catch((e) => { console.error(e); process.exit(1); });",
  );
  lines.push('}');
  lines.push('');
}

function hardwareWarning(program, lines, pad) {
  if (!program.hardware || !program.hardware.enabled) return;
  lines.push(
    `${pad}// WARNING: 'hardware: true' (line ${program.hardware.line}) has no effect — the documentation suite has no hardware section. Remove the line.`,
  );
  lines.push('');
}

/** Compiles a single-Part docx Program AST into JS source text. */
function generatePart(program, sourceFileLabel) {
  const meta = PART_TYPES.get(program.doc.docType);
  const registry = setterRegistry(meta.SDK);
  const lines = [];

  header(lines, sourceFileLabel);
  lines.push(`const { ${meta.className} } = require('pspt-docx');`);
  lines.push('');
  lines.push('async function build(outputPath) {');
  lines.push(`  const doc = new ${meta.className}();`);
  lines.push('');
  hardwareWarning(program, lines, '  ');

  for (const stmt of program.statements) {
    if (stmt.type === 'Part') {
      lines.push(
        `  // WARNING: 'part ${stmt.name}' (line ${stmt.line}) is only valid in a \`type=master\` file — this file is \`type=${program.doc.docType}\`, so the block was skipped.`,
      );
      lines.push('');
      continue;
    }
    if (stmt.type === 'Section') lines.push(`  // section ${stmt.name}`);
    const body = stmt.type === 'Section' ? stmt.body : [stmt];
    lines.push(...emitPartBody(body, registry, 'doc', '  '));
    lines.push('');
  }

  lines.push('  await doc.generate(outputPath);');
  footer(lines, `${program.doc.docType}.docx`);
  return lines.join('\n');
}

/** Compiles a `type=master` Program AST into JS source text. */
function generateMaster(program, sourceFileLabel) {
  const lines = [];
  header(lines, sourceFileLabel);
  lines.push("const { MasterDocument } = require('pspt-docx');");
  lines.push('');
  lines.push('async function build(outputPath) {');
  lines.push('  const master = new MasterDocument();');
  lines.push('');
  hardwareWarning(program, lines, '  ');

  const seen = new Set();
  for (const stmt of program.statements) {
    if (stmt.type !== 'Part') {
      const where = stmt.type === 'Section' ? `section '${stmt.name}'` : `a top-level ${stmt.type}`;
      lines.push(
        `  // WARNING: ${where} (line ${stmt.line}) sits outside any 'part <key> { ... }' block. In a \`type=master\` file every setter belongs to a Part — skipped.`,
      );
      lines.push('');
      continue;
    }
    const meta = PART_TYPES.get(stmt.name);
    if (!meta) {
      const hint = suggest(stmt.name, [...PART_TYPES.keys()]);
      lines.push(
        `  // WARNING: unknown part '${stmt.name}' (line ${stmt.line})${hint ? ` — did you mean '${hint}'?` : ''}. Valid parts: ${[...PART_TYPES.keys()].join(', ')} — skipped.`,
      );
      lines.push('');
      continue;
    }
    if (seen.has(stmt.name)) {
      lines.push(
        `  // WARNING: part '${stmt.name}' (line ${stmt.line}) appears more than once — this block's setters are applied on top of the earlier one.`,
      );
    }
    seen.add(stmt.name);

    lines.push(`  // ${meta.part == null ? 'front matter' : `Part ${meta.part}`} — ${meta.title}`);
    lines.push(`  const ${stmt.name} = master.part(${jsLiteral(stmt.name)});`);
    lines.push(...emitPartBody(stmt.body, setterRegistry(meta.SDK), stmt.name, '  '));
    lines.push('');
  }

  lines.push('  await master.generate(outputPath);');
  footer(lines, 'product-documentation-master.docx');
  return lines.join('\n');
}

/** Compiles an xlsx Program AST into JS source text. */
function generateXlsx(program, sourceFileLabel) {
  const lines = [];
  header(lines, sourceFileLabel);
  lines.push("const ExcelTrackerSDK = require('pspt-xlsx');");
  lines.push('');
  lines.push('async function build(outputPath) {');
  lines.push('  const tracker = new ExcelTrackerSDK();');
  lines.push(`  tracker.setTitle(${jsLiteral(program.doc.title)});`);

  if (program.calendar) {
    lines.push(
      `  tracker.setCalendarRange(${jsLiteral(program.calendar.from)}, ${jsLiteral(program.calendar.to)});`,
    );
  }
  lines.push('');

  program.statements.forEach((stmt) => {
    if (stmt.type === 'Group') {
      lines.push(`  tracker.addSection({`);
      lines.push(`    title: ${jsLiteral(stmt.title)},`);
      if (stmt.bannerColor) lines.push(`    bannerColor: ${jsLiteral(stmt.bannerColor)},`);
      if (stmt.rowColor) lines.push(`    rowColor: ${jsLiteral(stmt.rowColor)},`);
      lines.push('    tasks: [');
      stmt.tasks.forEach((task) => {
        const attrParts = [`name: ${jsLiteral(task.name)}`];
        Object.entries(task.attrs).forEach(([k, v]) => {
          let jsKey;
          if (k === 'start') {
            jsKey = 'startDate';
          } else if (k === 'end') {
            jsKey = 'endDate';
          } else {
            jsKey = k;
          }
          attrParts.push(`${jsKey}: ${valueNodeToJs(v)}`);
        });
        lines.push(`      { ${attrParts.join(', ')} },`);
      });
      lines.push('    ],');
      lines.push('  });');
      lines.push('');
    } else if (stmt.type === 'Callout') {
      const parts = [`dateStr: ${jsLiteral(stmt.date)}`, `text: ${jsLiteral(stmt.text)}`];
      if (stmt.color) parts.push(`color: ${jsLiteral(stmt.color)}`);
      lines.push(`  tracker.addCallout({ ${parts.join(', ')} });`);
    }
  });

  lines.push('');
  lines.push('  await tracker.generate(outputPath);');
  footer(lines, 'output.xlsx');
  return lines.join('\n');
}

function generate(program, sourceFileLabel) {
  if (program.doc.docType === 'xlsx') return generateXlsx(program, sourceFileLabel);
  if (program.doc.docType === 'master') return generateMaster(program, sourceFileLabel);
  return generatePart(program, sourceFileLabel);
}

module.exports = { generate, generatePart, generateMaster, generateXlsx };
