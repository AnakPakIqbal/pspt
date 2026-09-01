'use strict';
/**
 * Copies llms/ into a consumer that serves these docs (e.g. project-tracker-api,
 * which exposes them at /llms.txt and /llms/*.md).
 *
 *   node scripts/sync-llms-docs.js [targetDir]          # copy
 *   node scripts/sync-llms-docs.js [targetDir] --check  # fail if out of sync
 *
 * Default target is ../project-tracker-api/src/llms. The point of the check
 * mode is that these docs are served publicly: a stale copy tells an agent to
 * call an API that no longer exists, which is how the previous drift happened.
 */

const fs = require('fs');
const path = require('path');

const SOURCE = path.join(__dirname, '..', 'llms');
const DEFAULT_TARGET = path.join(__dirname, '..', '..', 'project-tracker-api', 'src', 'llms');

const args = process.argv.slice(2);
const check = args.includes('--check');
const target = path.resolve(args.find((arg) => !arg.startsWith('--')) || DEFAULT_TARGET);

if (!fs.existsSync(SOURCE)) {
  console.error(`No llms/ directory at ${SOURCE} — run \`npm run llms:build\` first.`);
  process.exit(1);
}

const files = fs
  .readdirSync(SOURCE)
  .filter((name) => name.endsWith('.md') || name.endsWith('.txt'));
const drift = [];

for (const file of files) {
  const from = path.join(SOURCE, file);
  const dest = path.join(target, file);
  const source = fs.readFileSync(from);
  const current = fs.existsSync(dest) ? fs.readFileSync(dest) : null;
  if (current && current.equals(source)) continue;
  drift.push(file + (current ? '' : ' (missing)'));
  if (!check) {
    fs.mkdirSync(target, { recursive: true });
    fs.writeFileSync(dest, source);
  }
}

// Files the target has that the source no longer does are stale leftovers.
const orphans = fs.existsSync(target)
  ? fs
      .readdirSync(target)
      .filter((name) => (name.endsWith('.md') || name.endsWith('.txt')) && !files.includes(name))
  : [];
for (const file of orphans) {
  drift.push(file + ' (removed upstream)');
  if (!check) fs.unlinkSync(path.join(target, file));
}

const rel = path.relative(process.cwd(), target);

if (check) {
  if (drift.length) {
    console.error(`${rel} is out of sync with llms/:`);
    for (const entry of drift) console.error('  - ' + entry);
    console.error('Run `npm run llms:sync`.');
    process.exit(1);
  }
  console.log(`${rel} is in sync with llms/ (${files.length} files).`);
} else if (drift.length) {
  console.log(`synced ${drift.length} file(s) to ${rel}:`);
  for (const entry of drift) console.log('  - ' + entry);
} else {
  console.log(`${rel} already in sync (${files.length} files).`);
}
