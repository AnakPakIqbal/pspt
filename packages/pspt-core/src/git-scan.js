'use strict';
/**
 * Git-scan utilities, absorbing the ad hoc copy-pasted `execSync('git log ...')`
 * scripts (docs/scan-commits.js, docs/scan-agentic-repos.js) into one reusable
 * module. Used by `pspt scan-git`.
 */

const { execSync } = require('child_process');

/**
 * Runs `git show --shortstat` for one commit and parses insertions/deletions.
 * Same convention as the original docs/generate-pt-trackers.js `getCommitStats`.
 * @param {string} repoPath
 * @param {string} hash
 * @returns {{insertions:number, deletions:number}}
 */
function getCommitDiffStat(repoPath, hash) {
  try {
    const stat = execSync(`git show --shortstat --format="" ${hash.trim()}`, {
      cwd: repoPath,
      encoding: 'utf8',
    }).trim();
    const insMatch = stat.match(/(\d+)\s+insertion/);
    const delMatch = stat.match(/(\d+)\s+deletion/);
    return {
      insertions: insMatch ? parseInt(insMatch[1], 10) : 0,
      deletions: delMatch ? parseInt(delMatch[1], 10) : 0,
    };
  } catch (err) {
    return { insertions: 0, deletions: 0 };
  }
}

/**
 * Sums insertions/deletions across a list of commit hashes, formatted as the
 * `lines` string ExcelTrackerSDK tasks expect (e.g. "+120 / -30").
 * @param {string} repoPath
 * @param {string|string[]} hashes
 * @returns {{commitCount:number, insertions:number, deletions:number, lines:string}}
 */
function getCommitStats(repoPath, hashes) {
  const hashList = Array.isArray(hashes) ? hashes : [hashes];
  let totalIns = 0;
  let totalDel = 0;
  hashList.forEach((hash) => {
    const { insertions, deletions } = getCommitDiffStat(repoPath, hash);
    totalIns += insertions;
    totalDel += deletions;
  });
  return {
    commitCount: hashList.length,
    insertions: totalIns,
    deletions: totalDel,
    lines: `+${totalIns} / -${totalDel}`,
  };
}

/**
 * Runs `git log --all` in a repo and returns deduplicated, parsed commits.
 * @param {string} repoPath - filesystem path to the git repo.
 * @param {Object} [opts]
 * @param {string} [opts.since] - passed to `--since` (e.g. '2026-01-01').
 * @param {string} [opts.authorEmail] - if given, only commits whose <email> includes this string are kept.
 * @param {boolean} [opts.withDiffStat] - if true, run `git show --shortstat` per
 *   commit and attach {insertions, deletions} to each entry. Off by default
 *   since it spawns one extra git process per commit and can be slow on large
 *   histories — opt in when you actually need per-commit line counts.
 * @returns {Array<{hash:string,date:string,author:string,email:string,subject:string,insertions?:number,deletions?:number}>}
 */
function scanRepoCommits(repoPath, opts = {}) {
  const sinceArg = opts.since ? ` --since="${opts.since}"` : '';
  const rawLog = execSync(
    `git log --all --pretty=format:"%h|%ad|%an|<%ae>|%s" --date=short${sinceArg}`,
    { cwd: repoPath, encoding: 'utf8' },
  );
  const lines = rawLog.split('\n').filter(Boolean);

  const seen = new Set();
  const commits = [];
  lines.forEach((line) => {
    const [hash, date, author, email, ...subjectParts] = line.split('|');
    if (seen.has(hash)) return;
    seen.add(hash);
    if (opts.authorEmail && !(email && email.includes(opts.authorEmail))) return;
    const commit = { hash, date, author, email, subject: subjectParts.join('|') };
    if (opts.withDiffStat) {
      const { insertions, deletions } = getCommitDiffStat(repoPath, hash);
      commit.insertions = insertions;
      commit.deletions = deletions;
    }
    commits.push(commit);
  });
  return commits;
}

/**
 * Returns { remote, branches, commits } for a repo — remote/branches best-effort.
 * @param {string} repoPath
 * @param {Object} [opts] - same as scanRepoCommits
 */
function scanRepoFull(repoPath, opts = {}) {
  const result = { path: repoPath, remote: null, branches: null, commits: [], error: null };
  try {
    result.remote = execSync('git config --get remote.origin.url', {
      cwd: repoPath,
      encoding: 'utf8',
    }).trim();
  } catch (err) {
    result.remote = null;
  }
  try {
    result.branches = execSync('git branch -a', { cwd: repoPath, encoding: 'utf8' }).trim();
  } catch (err) {
    result.branches = null;
  }
  try {
    result.commits = scanRepoCommits(repoPath, opts);
  } catch (err) {
    result.error = err.message;
  }
  return result;
}

/**
 * Scans multiple repos, each described as {name, path} (url optional).
 * @param {Array<{name:string, path:string, url?:string}>} repos
 * @param {Object} [opts] - {since, authorEmail}
 * @returns {Array<Object>} one result object per repo, in input order.
 */
function scanRepos(repos, opts = {}) {
  return repos.map((repo) => {
    const scanned = scanRepoFull(repo.path, opts);
    return { name: repo.name, url: repo.url || null, ...scanned };
  });
}

module.exports = { scanRepoCommits, scanRepoFull, scanRepos, getCommitStats, getCommitDiffStat };
