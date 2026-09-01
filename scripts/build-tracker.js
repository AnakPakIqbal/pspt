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

/** Every commit, oldest first, with the files it touched and its diff stat. */
function readHistory() {
  const SEP = '';
  const raw = git([
    'log',
    '--all',
    '--reverse',
    '--name-only',
    `--pretty=format:${SEP}%h|%ad|%s`,
    '--date=short',
  ]);

  const commits = [];
  for (const chunk of raw.split(SEP)) {
    if (!chunk.trim()) continue;
    const [header, ...rest] = chunk.split('\n');
    const [hash, date, ...subjectParts] = header.split('|');
    const files = rest.map((line) => line.trim()).filter(Boolean);
    commits.push({ hash, date, subject: subjectParts.join('|'), files });
  }

  // One `git show --shortstat` per commit is what gives real +/- numbers;
  // parsing the subject line for them would be guesswork.
  for (const commit of commits) {
    const stat = git(['show', '--shortstat', '--format=', commit.hash]);
    const insertions = /(\d+) insertion/.exec(stat);
    const deletions = /(\d+) deletion/.exec(stat);
    commit.insertions = insertions ? Number(insertions[1]) : 0;
    commit.deletions = deletions ? Number(deletions[1]) : 0;
    commit.churn = commit.insertions + commit.deletions;
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

function areaOf(commit) {
  const votes = new Map();
  const bump = (key, weight = 1) => votes.set(key, (votes.get(key) || 0) + weight);

  for (const file of commit.files) {
    const parts = file.split('/');
    const featureIdx = parts.indexOf('features');
    const feature = featureIdx >= 0 ? parts[featureIdx + 1] : null;
    if (feature && !feature.includes('.')) {
      bump(feature, 3); // a feature folder is the strongest signal
      continue;
    }
    if (parts[0] === 'packages' && parts[1] && !parts[1].includes('.')) {
      bump(parts[1], 3);
      continue;
    }
    if (parts[0] === 'prisma' || parts[0] === 'db') bump('prisma', 2);
    else if (parts[0] === 'scripts') bump('scripts', 2);
    else if (parts[0] === 'llms') bump('llms', 2);
    else if (parts[0] === 'examples') bump('examples', 1);
    else if (parts[0] === 'tests' || parts[0] === 'e2e') bump('tests', 1);
    else if (parts[0] === 'docs') bump('docs', 1);
    else if (parts[0] === 'docker' || /^Dockerfile|docker-compose/.test(parts[0]))
      bump('docker', 2);
    else if (parts[1] === 'shared' || parts[0] === 'shared') bump('shared', 1);
    else if (parts[1] === 'pages' || parts[1] === 'app') bump('pages', 1);
    else if (parts[1] === 'llms') bump('llms', 2);
    else if (parts[1] === 'container') bump('container', 1);
  }

  if (!votes.size) return 'Platform';
  if (votes.size >= 6) return 'Platform Foundations';
  const [top] = [...votes.entries()].sort((a, b) => b[1] - a[1]);
  return (
    AREA_LABELS[top[0]] ||
    top[0].replace(/(^|-)([a-z])/g, (_, sep, ch) => (sep ? ' ' : '') + ch.toUpperCase())
  );
}

/** Turns a commit subject into a short business-language phrase. */
function phrase(subject) {
  let text = subject
    .replace(/^[a-z]+(\([^)]*\))?!?:\s*/i, '') // strip the conventional-commit prefix
    .replace(/\s+/g, ' ')
    .trim();
  if (!text) return '';
  text = text.charAt(0).toUpperCase() + text.slice(1);
  // Keep rows readable in the sheet's Task column.
  if (text.length > 78) {
    const cut = text.slice(0, 78);
    text = cut.slice(0, cut.lastIndexOf(' ')) + '…';
  }
  return text;
}

/** Merge same-day, same-area commits into one row. */
function toRows(commits) {
  const groups = new Map();
  for (const commit of commits) {
    const area = areaOf(commit);
    const key = `${commit.date}|${area}`;
    if (!groups.has(key)) groups.set(key, { date: commit.date, area, commits: [] });
    groups.get(key).commits.push(commit);
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
        name: `${group.area} — ${phrase(lead.subject)}`,
        detail: group.commits.map((item) => phrase(item.subject)).join('; '),
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
const MONTH_KEY_LENGTH = 'YYYY-MM'.length;
const MAX_AREAS_IN_BANNER = 4;

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
    const areas = [...new Set(monthRows.map((r) => r.area))]
      .slice(0, MAX_AREAS_IN_BANNER)
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
