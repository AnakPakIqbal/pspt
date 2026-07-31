'use strict';

const fs = require('fs');
const path = require('path');
const { scanRepos } = require('pspt-core');

/**
 * Scans one or more git repo paths and writes the results as JSON.
 * Replaces the ad hoc docs/scan-commits.js / docs/scan-agentic-repos.js scripts.
 * @param {string[]} repoPaths
 * @param {Object} opts - {since, out, withDiffStat} — pass --with-lines on the
 *   CLI to set withDiffStat and attach {insertions, deletions} to every commit
 *   (one extra `git show --shortstat` per commit, so opt-in only).
 */
function scanGit(repoPaths, opts = {}) {
  const repos = repoPaths.map((p) => ({ name: path.basename(p), path: p }));
  const results = scanRepos(repos, { since: opts.since, withDiffStat: opts.withDiffStat });
  const outPath = opts.out || 'scan-git.json';
  fs.writeFileSync(outPath, JSON.stringify(results, null, 2));
  return { outPath, results };
}

module.exports = { scanGit };
