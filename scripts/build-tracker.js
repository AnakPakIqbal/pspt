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

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

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
  if (/(^|[/])(package-lock[.]json|bun[.]lockb|yarn[.]lock)$/.test(filePath)) return 0.02;
  if (/(^|[/])(docs|llms)[/]/.test(filePath) || /[.](md|ya?ml)$/.test(filePath)) return 0.15;
  if (
    /(^|[/])(tests?|e2e|__tests__|fixtures)[/]/.test(filePath) ||
    /[.](spec|test)[.]/.test(filePath)
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
    else if (parts[0] === 'docker' || /^Dockerfile|docker-compose/.test(parts[0]))
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
const LEADING_VERBS =
  /^(implement|implemented|add|added|adds|create|created|introduce|introduced|refactor|refactored|update|updated|improve|improved|enhance|enhanced|integrate|integrated|overhaul|overhauled|scaffold|scaffolded|fix|fixed|resolve|resolved|remove|removed|rename|renamed|reorganize|reorganized|extend|extended|transition to|migrate to|migrated to|support|complete|finalize)\s+/i;

const CLAUSE_BOUNDARY =
  /(?:\s+(?:with|using|including|and add|and update|and improve|and extend|for the|to the)\s+|\s*[,;—]\s+)/i;

function phrase(subject) {
  let text = subject
    .replace(/^[a-z]+(\([^)]*\))?!?:\s*/i, '')
    .replace(/\s+/g, ' ')
    .trim();
  text = text.replace(LEADING_VERBS, '');

  const boundary = CLAUSE_BOUNDARY.exec(text);
  if (boundary && boundary.index > MIN_PHRASE_LENGTH) text = text.slice(0, boundary.index);
  text = text.replace(/[.,;:]+$/, '').trim();
  if (!text) return '';

  const words = text.split(' ');
  if (words.length > MAX_PHRASE_WORDS) text = words.slice(0, MAX_PHRASE_WORDS).join(' ');

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

/** Merge same-day, same-area commits into one row. */
function toRows(commits) {
  const groups = new Map();
  for (const commit of commits) {
    const area = areaOf(commit);
    const key = `${commit.date}|${area.key}`;
    if (!groups.has(key)) {
      groups.set(key, { date: commit.date, area: area.key, labels: new Map(), commits: [] });
    }
    const bucket = groups.get(key);
    // The row is labelled for whichever spelling covers the most work.
    bucket.labels.set(area.label, (bucket.labels.get(area.label) || 0) + commit.churn);
    bucket.commits.push(commit);
  }

  return [...groups.values()]
    .sort((a, b) => a.date.localeCompare(b.date) || a.area.localeCompare(b.area))
    .map((group) => {
      const insertions = group.commits.reduce((sum, item) => sum + item.insertions, 0);
      const deletions = group.commits.reduce((sum, item) => sum + item.deletions, 0);
      // "Most critical" = the largest single commit by total lines changed.
      const lead = group.commits.reduce(
        (best, item) => (item.churn > best.churn ? item : best),
        group.commits[0],
      );
      return {
        date: group.date,
        area: group.area,
        name: `${bestLabel(group.labels, group.area)} — ${phrase(lead.subject)}`,
        detail: group.commits.map((item) => phrase(item.subject)).join('; '),
        churn: insertions + deletions,
        commitCount: group.commits.length,
        lines: `+${insertions} / -${deletions}`,
        leadHash: lead.hash,
      };
    });
}

const MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];
// Distinct banner per group; never the same colour twice in a row.
const BANNERS = ['blue', 'green', 'purple', 'yellow', 'pink', 'gray'];
const ROW_COLOR = 'green'; // uniform across the whole tracker, per the workflow
// A commit whose largest area holds less than this share of its changed lines
// has no real home; only then does it fall into the catch-all bucket.
const DOMINANCE_THRESHOLD = 0.45;
const MIN_PHRASE_LENGTH = 12; // do not truncate a phrase down to nothing
const MAX_PHRASE_WORDS = 7;
const MONTH_KEY_LENGTH = 'YYYY-MM'.length;
const MAX_AREAS_IN_BANNER = 3;

function esc(text) {
  return String(text).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

function render(rows, remote, trackerTitle) {
  const byMonth = new Map();
  for (const row of rows) {
    const key = row.date.slice(0, MONTH_KEY_LENGTH);
    if (!byMonth.has(key)) byMonth.set(key, []);
    byMonth.get(key).push(row);
  }

  const dates = rows.map((r) => r.date).sort();
  const out = [];
  out.push('// Generated by scripts/build-tracker.js from real git history.');
  out.push('// Rows are merged per day and per area, oldest first, with summed');
  out.push('// commit counts and diff stats. Re-run the script rather than editing.');
  out.push('');
  out.push(`doc "${esc(trackerTitle)}" type=xlsx`);
  out.push(`calendar ${dates[0]} .. ${dates[dates.length - 1]}`);
  out.push('');

  let banner = 0;
  for (const [month, monthRows] of [...byMonth].sort()) {
    const [year, m] = month.split('-');
    const color = BANNERS[banner % BANNERS.length];
    banner++;
    // Rank the month's areas by the work that landed in them, so the banner
    // names the themes of the month rather than whatever sorts first.
    const churnByArea = new Map();
    for (const row of monthRows) {
      for (const area of row.area.split(' & ')) {
        churnByArea.set(area, (churnByArea.get(area) || 0) + row.churn);
      }
    }
    const areas = [...churnByArea.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, MAX_AREAS_IN_BANNER)
      .map(([area]) => area)
      .join(', ');
    out.push(
      `group "${MONTHS[Number(m) - 1]} ${year}: ${esc(areas)}" color=${color} rowColor=${ROW_COLOR} {`,
    );
    for (const row of monthRows) {
      const link = remote ? `${remote.replace(/\.git$/, '')}/commit/${row.leadHash}` : '';
      const notes = link
        ? ` notes={ text: "${esc(row.leadHash)}", hyperlink: "${esc(link)}" }`
        : '';
      out.push(
        `  task "${esc(row.name)}" start=${row.date} end=${row.date}` +
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

// Rows are single-day by construction, but assert it against what was
// actually written rather than trusting the constructor: a tracker that
// quietly breaks the rule is worse than one that fails to build.
const spans = [...source.matchAll(/ start=(\S+) end=(\S+)/g)].filter(([, from, to]) => from !== to);
console.log(
  `${name}: ${commits.length} commits -> ${rows.length} merged rows ` +
    `(${new Set(rows.map((row) => row.date.slice(0, MONTH_KEY_LENGTH))).size} months, ` +
    `${new Set(rows.map((row) => row.area)).size} areas), ${spans.length} multi-day rows`,
);
if (spans.length > 0) {
  console.error(
    `${spans.length} row(s) span more than one day — the tracker rule requires single-day rows.`,
  );
  process.exit(1);
}
console.log(`wrote ${outPath}`);
