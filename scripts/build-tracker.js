'use strict';
/**
 * =============================================================================
 *  Authors a tracker .pspt from a repository's real git history.
 * =============================================================================
 * Follows llms/self-document-workflow.md's tracker rules exactly:
 *
 *   - one tracker per repository; histories are never merged
 *   - commits read oldest -> newest
 *   - commits touching the same area on the same day merge into ONE task row,
 *     with commitCount and lines summed across the group
 *   - every row is a single day (start === end); rows never span a range
 *   - `notes` carries exactly one link — the group's largest commit by total
 *     lines changed — as a {text, hyperlink} object
 *   - each month is a `group` with its own banner colour; every row uses the
 *     same rowColor throughout
 *
 * The area a commit belongs to is inferred from the files it touched, not from
 * its subject line: subjects here are mostly bare `feat:` with no scope, and
 * the workflow is explicit that the diff, not the message, is the source.
 *
 *   node scripts/build-tracker.js <repo-path> <out.pspt> "<Tracker Title>"
 * =============================================================================
 */

const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const [repoPath, outPath, title] = process.argv.slice(2);
if (!repoPath || !outPath) {
  console.error('usage: node scripts/build-tracker.js <repo-path> <out.pspt> "<Title>"');
  process.exit(1);
}

const git = (args) =>
  execFileSync('git', ['-C', repoPath, ...args], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });

/**
 * Every commit, oldest first, with per-file insertions/deletions.
 *
 * `--numstat` gives the real +/- per file in a single pass, which both avoids
 * one `git show` per commit and — more importantly — lets a commit's area be
 * decided by how many lines landed where, not merely by how many files it
 * touched. A wide-ranging commit is usually 90% one feature plus incidental
 * edits elsewhere; counting files alone loses that.
 */
function readHistory() {
  // A record separator git passes through verbatim and no subject contains.
  const SEP = '@@PSPT-COMMIT@@';
  const raw = git([
    'log',
    '--all',
    '--reverse',
    '--numstat',
    `--pretty=format:${SEP}%h|%ad|%s`,
    '--date=short',
  ]);

  const commits = [];
  for (const chunk of raw.split(SEP)) {
    if (!chunk.trim()) continue;
    const [header, ...rest] = chunk.split('\n');
    const [hash, date, ...subjectParts] = header.split('|');
    const files = [];
    let insertions = 0;
    let deletions = 0;
    for (const line of rest) {
      const parts = line.split('\t');
      if (parts.length < 3) continue;
      const added = parts[0] === '-' ? 0 : Number(parts[0]);
      const removed = parts[1] === '-' ? 0 : Number(parts[1]);
      insertions += added;
      deletions += removed;
      files.push({ path: parts[2], churn: added + removed });
    }
    commits.push({
      hash,
      date,
      subject: subjectParts.join('|'),
      files,
      insertions,
      deletions,
      churn: insertions + deletions,
    });
  }
  return commits;
}

/** Human label for a code area, derived from the paths a commit touched. */
const AREA_LABELS = {
  auth: 'Authentication',
  users: 'Users',
  permissions: 'Permissions',
  projects: 'Projects',
  tasks: 'Tasks',
  tags: 'Tags',
  ai: 'AI Assistant',
  'ai-brainstorm': 'AI Assistant',
  attachments: 'Attachments',
  'activity-logs': 'Activity Log',
  activity: 'Activity Log',
  invitations: 'Invitations',
  departments: 'Departments',
  support: 'Help & Support',
  'help-support': 'Help & Support',
  mcp: 'MCP Server',
  oauth: 'Google Sign-in',
  'llms-docs': 'Agent Documentation',
  dashboard: 'Dashboard',
  people: 'People',
  profile: 'Profile',
  archive: 'Archive',
  'project-summary': 'Project Summary',
  'user-report': 'Reporting',
  reports: 'Reporting',
  prisma: 'Data Model',
  db: 'Data Model',
  docs: 'Documentation',
  llms: 'Agent Documentation',
  tests: 'Testing',
  e2e: 'Testing',
  docker: 'Deployment',
  shared: 'Shared Foundations',
  container: 'Shared Foundations',
  app: 'Application Shell',
  scripts: 'Build Tooling',
  examples: 'Examples & Fixtures',
  'pspt-core': 'Core (tokens & helpers)',
  'pspt-docx': 'Docx SDK',
  'pspt-xlsx': 'Xlsx SDK',
  'pspt-lang': 'DSL Compiler',
  'pspt-cli': 'CLI',
  pages: 'Application Shell',
};

/** How much a file's churn counts toward deciding a commit's area. */
function incidence(filePath) {
  if (/(?:^|\/)(?:package-lock\.json|bun\.lockb|yarn\.lock)$/.test(filePath)) return 0.02;
  if (/(?:^|\/)(?:docs|llms)\//.test(filePath) || /\.(md|ya?ml)$/.test(filePath)) return 0.15;
  if (
    /(?:^|\/)(?:tests?|e2e|__tests__|fixtures)\//.test(filePath) ||
    /\.(spec|test)\./.test(filePath)
  )
    return 0.3;
  return 1;
}

function areaOf(commit) {
  const votes = new Map();
  const bump = (key, weight) => votes.set(key, (votes.get(key) || 0) + weight);

  for (const file of commit.files) {
    const parts = file.path.split('/');
    // Every file carries at least a little weight, so a pure rename still
    // votes; but a 2,000-line OpenAPI or lockfile diff is a side effect of the
    // feature that shipped, not the thing itself, so it counts for less.
    const weight = Math.max(file.churn, 1) * incidence(file.path);

    const featureIdx = parts.indexOf('features');
    const feature = featureIdx >= 0 ? parts[featureIdx + 1] : null;
    if (feature && !feature.includes('.')) {
      bump(feature, weight);
      continue;
    }
    if (parts[0] === 'packages' && parts[1] && !parts[1].includes('.')) bump(parts[1], weight);
    else if (parts[0] === 'prisma' || parts[0] === 'db') bump('prisma', weight);
    else if (parts[0] === 'scripts') bump('scripts', weight);
    else if (parts[0] === 'llms' || parts[1] === 'llms') bump('llms', weight);
    else if (parts[0] === 'examples') bump('examples', weight);
    else if (parts[0] === 'tests' || parts[0] === 'e2e') bump('tests', weight);
    else if (parts[0] === 'docs') bump('docs', weight);
    else if (parts[0] === 'docker' || /^(?:Dockerfile|docker-compose)/.test(parts[0]))
      bump('docker', weight);
    else if (parts[1] === 'shared' || parts[0] === 'shared') bump('shared', weight);
    else if (parts[1] === 'pages' || parts[1] === 'app') bump('pages', weight);
    else if (parts[1] === 'container') bump('container', weight);
  }

  if (!votes.size) return { key: 'Platform', label: 'Platform' };

  const ranked = [...votes.entries()].sort((a, b) => b[1] - a[1]);
  const total = ranked.reduce((sum, [, weight]) => sum + weight, 0);
  const [topKey, topWeight] = ranked[0];

  // When one area clearly dominates, name the row for it. Otherwise name it
  // for the two areas it actually spans: a generic bucket repeated dozens of
  // times tells a reader nothing about what shipped that day, and is what
  // makes a tracker read as bloat.
  const primary = label(topKey);
  const spansTwo = ranked.length > 1 && topWeight / total < DOMINANCE_THRESHOLD;
  return { key: primary, label: spansTwo ? `${primary} & ${label(ranked[1][0])}` : primary };
}

function label(key) {
  return (
    AREA_LABELS[key] ||
    key.replace(/(^|-)([a-z])/g, (_, sep, ch) => (sep ? ' ' : '') + ch.toUpperCase())
  );
}

/**
 * Turns a commit subject into a short business-language phrase.
 *
 * The workflow is explicit that a row is named for what shipped, not for the
 * commit message: "Auth Middleware & Invitation Service" rather than "feat:
 * implement JWT middleware + invite-by-email flow". So drop the conventional
 * prefix, drop the leading verb, and keep the noun phrase up to the first
 * clause boundary.
 */
const LEADING_VERBS = new Set([
  'implement',
  'add',
  'create',
  'introduce',
  'refactor',
  'update',
  'improve',
  'enhance',
  'integrate',
  'overhaul',
  'scaffold',
  'fix',
  'resolve',
  'remove',
  'rename',
  'reorganize',
  'extend',
  'support',
  'complete',
  'finalize',
  'migrate',
  'transition',
]);

/** Clause openers that mark where a subject stops naming the thing shipped. */
const CLAUSE_OPENERS = new Set(['with', 'using', 'including', 'for', 'to', 'and']);

/** Drops a leading verb (in any common inflection) from a subject. */
function stripLeadingVerb(text) {
  const [first, ...rest] = text.split(' ');
  if (!first) return text;
  const stem = first.toLowerCase().replace(/(?:ed|es|s|d)$/, '');
  if (!LEADING_VERBS.has(stem) && !LEADING_VERBS.has(first.toLowerCase())) return text;
  // "migrate to" / "transition to" carry their preposition with them.
  if (rest[0] === 'to' && (stem === 'migrate' || stem === 'transition'))
    return rest.slice(1).join(' ');
  return rest.join(' ');
}

/** Index of the first clause opener, or -1 — where the noun phrase ends. */
function clauseStart(words) {
  for (let i = 1; i < words.length; i++) {
    const word = words[i].toLowerCase();
    if (CLAUSE_OPENERS.has(word)) return i;
    if (/[,;—]$/.test(words[i - 1])) return i;
  }
  return -1;
}

function phrase(subject) {
  let text = subject
    .replace(/^[a-z]{2,12}(?:\([a-z0-9,/ -]{1,40}\))?!?: /i, '')
    .replace(/\s+/g, ' ')
    .trim();
  text = stripLeadingVerb(text);

  const words = text.split(' ').filter(Boolean);
  const boundary = clauseStart(words);
  const kept = boundary > 0 ? words.slice(0, boundary) : words;
  text = kept
    .slice(0, MAX_PHRASE_WORDS)
    .join(' ')
    .replace(/[.,;:]{1,4}$/, '')
    .trim();
  if (!text) return '';

  // Title-case, leaving acronyms and already-capitalised words alone.
  return text
    .split(' ')
    .map((word) =>
      /^[A-Z0-9/-]{2,}$/.test(word) || /[A-Z]/.test(word.slice(1))
        ? word
        : word.charAt(0).toUpperCase() + word.slice(1),
    )
    .join(' ');
}

/** The label covering the most work in a merged row. */
function bestLabel(labels, fallback) {
  const ranked = [...labels.entries()].sort((a, b) => b[1] - a[1]);
  return ranked.length ? ranked[0][0] : fallback;
}

/** Distinct phrases, longest first, so a summary leads with the biggest thing. */
function summarise(commits) {
  const seen = new Map();
  for (const commit of commits) {
    const text = phrase(commit.subject);
    if (!text) continue;
    seen.set(text, Math.max(seen.get(text) || 0, commit.churn));
  }
  return [...seen.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([text]) => text)
    .join('; ');
}

const DAY_MS = 24 * 60 * 60 * 1000;

/**
 * Merges commits into one row per continuous run of work on a feature.
 *
 * A row is a task, and the commits under it are its subtasks — so the row is
 * named for the feature and its detail summarises what those commits did.
 * Splitting the same feature into one row per calendar day is what made the
 * sheet read as noise: a feature built over four days is one piece of work,
 * not four. Consecutive days (and a weekend gap) stay in the same run; a real
 * pause starts a new one.
 */
function toRows(commits) {
  const byArea = new Map();
  for (const commit of commits) {
    const area = areaOf(commit);
    if (!byArea.has(area.key)) byArea.set(area.key, []);
    byArea.get(area.key).push({ ...commit, label: area.label });
  }

  const rows = [];
  for (const [area, areaCommits] of byArea) {
    areaCommits.sort((a, b) => a.date.localeCompare(b.date));

    let run = [];
    const flush = () => {
      if (!run.length) return;
      rows.push(makeRow(area, run));
      run = [];
    };
    for (const commit of areaCommits) {
      if (run.length) {
        const gap = (Date.parse(commit.date) - Date.parse(run[run.length - 1].date)) / DAY_MS;
        if (gap > RUN_GAP_DAYS) flush();
      }
      run.push(commit);
    }
    flush();
  }

  return rows.sort((a, b) => a.start.localeCompare(b.start) || a.area.localeCompare(b.area));
}

function makeRow(area, run) {
  const insertions = run.reduce((sum, item) => sum + item.insertions, 0);
  const deletions = run.reduce((sum, item) => sum + item.deletions, 0);
  // "Most critical" = the largest single commit by total lines changed.
  const lead = run.reduce((best, item) => (item.churn > best.churn ? item : best), run[0]);

  const labels = new Map();
  for (const commit of run)
    labels.set(commit.label, (labels.get(commit.label) || 0) + commit.churn);

  return {
    area,
    start: run[0].date,
    end: run[run.length - 1].date,
    churn: insertions + deletions,
    name: `${bestLabel(labels, area)} — ${phrase(lead.subject)}`,
    detail: summarise(run),
    commitCount: run.length,
    lines: `+${insertions} / -${deletions}`,
    leadHash: lead.hash,
  };
}

// Distinct banner per group; never the same colour twice in a row.
const BANNERS = ['blue', 'green', 'purple', 'yellow', 'pink', 'gray'];
const ROW_COLOR = 'green'; // uniform across the whole tracker, per the workflow
// A commit whose largest area holds less than this share of its changed lines
// has no real home; only then does it fall into the catch-all bucket.
const DOMINANCE_THRESHOLD = 0.45;
const MAX_PHRASE_WORDS = 7;
// Days of quiet that end a run of work on a feature. Three keeps a Friday
// and the following Monday together.
const RUN_GAP_DAYS = 3;

function esc(text) {
  return String(text).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

function render(rows, remote, trackerTitle) {
  // Group by feature, not by calendar month: a reader scanning the sheet wants
  // "everything that happened to Tasks", not "everything that happened in June".
  const byArea = new Map();
  for (const row of rows) {
    if (!byArea.has(row.area)) byArea.set(row.area, []);
    byArea.get(row.area).push(row);
  }

  const dates = rows.flatMap((row) => [row.start, row.end]).sort();
  const out = [];
  out.push('// Generated by scripts/build-tracker.js from real git history.');
  out.push('// One group per feature; one row per continuous run of work on it,');
  out.push('// summing its commits. Re-run the script rather than editing.');
  out.push('');
  out.push(`doc "${esc(trackerTitle)}" type=xlsx`);
  out.push(`calendar ${dates[0]} .. ${dates[dates.length - 1]}`);
  out.push('');

  // Biggest feature first, so the sheet opens on the substance.
  const ordered = [...byArea.entries()].sort(
    (a, b) =>
      b[1].reduce((sum, row) => sum + row.churn, 0) - a[1].reduce((sum, row) => sum + row.churn, 0),
  );

  let banner = 0;
  for (const [area, areaRows] of ordered) {
    const color = BANNERS[banner % BANNERS.length];
    banner++;
    areaRows.sort((a, b) => a.start.localeCompare(b.start));
    const commitTotal = areaRows.reduce((sum, row) => sum + row.commitCount, 0);
    out.push(
      `group "${esc(area)} (${commitTotal} commits)" color=${color} rowColor=${ROW_COLOR} {`,
    );
    for (const row of areaRows) {
      const link = remote ? `${remote.replace(/\.git$/, '')}/commit/${row.leadHash}` : '';
      const notes = link
        ? ` notes={ text: "${esc(row.leadHash)}", hyperlink: "${esc(link)}" }`
        : '';
      out.push(
        `  task "${esc(row.name)}" start=${row.start} end=${row.end}` +
          ` commitCount=${row.commitCount} lines="${row.lines}"` +
          ` detail="${esc(row.detail)}"${notes}`,
      );
    }
    out.push('}');
    out.push('');
  }
  return out.join('\n');
}

const commits = readHistory();
const rows = toRows(commits);
const remote = (() => {
  try {
    return git(['config', '--get', 'remote.origin.url']).trim();
  } catch {
    return '';
  }
})();

const name = path.basename(path.resolve(repoPath));
fs.mkdirSync(path.dirname(outPath), { recursive: true });

const source = render(rows, remote, title || `${name} Tracker`);
fs.writeFileSync(outPath, source);

const areas = new Set(rows.map((row) => row.area));
const spanning = rows.filter((row) => row.start !== row.end).length;
console.log(
  `${name}: ${commits.length} commits -> ${rows.length} rows in ${areas.size} feature groups ` +
    `(${spanning} spanning more than one day)`,
);

console.log(`wrote ${outPath}`);
