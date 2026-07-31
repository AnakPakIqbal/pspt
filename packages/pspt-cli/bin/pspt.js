#!/usr/bin/env node
'use strict';

const { buildFile } = require('../src/build');
const { compileFile } = require('../src/compile');
const { scanGit } = require('../src/scan-git');

function parseFlags(args) {
  const positional = [];
  const flags = {};
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg.startsWith('--')) {
      const key = arg.slice(2);
      const next = args[i + 1];
      if (next !== undefined && !next.startsWith('--')) {
        flags[key] = next;
        i++;
      } else {
        flags[key] = true;
      }
    } else {
      positional.push(arg);
    }
  }
  return { positional, flags };
}

function printUsage() {
  console.log(`pspt — compile and build .pspt product-spec / tracker files

Usage:
  pspt compile <file>.pspt                 Compile to <file>.gen.js
  pspt build <file>.pspt --out <output>    Compile + run, writing <output> (.docx/.xlsx)
  pspt scan-git <repo...> --since <date> --out <data>.json [--with-lines]
                                            Scan git history across one or more repos.
                                            --with-lines attaches {insertions, deletions}
                                            per commit (one extra \`git show --shortstat\`
                                            per commit — slower on large histories).

Examples:
  pspt compile examples/fixtures/sample-docx.pspt
  pspt build examples/fixtures/sample-docx.pspt --out out/spec.docx
  pspt scan-git ../project-tracker ../project-tracker-api --since 2026-01-01 --out scan.json --with-lines
`);
}

async function main() {
  const [, , cmd, ...rest] = process.argv;
  const { positional, flags } = parseFlags(rest);

  try {
    if (cmd === 'compile') {
      const file = positional[0];
      if (!file) {
        printUsage();
        process.exitCode = 1;
        return;
      }
      const result = compileFile(file, flags.out);
      console.log(`Compiled ${file} -> ${result.outPath} (type=${result.docType})`);
      return;
    }

    if (cmd === 'build') {
      const file = positional[0];
      if (!file) {
        printUsage();
        process.exitCode = 1;
        return;
      }
      const result = await buildFile(file, flags.out);
      console.log(`Built ${file} -> ${result.outputPath} (type=${result.docType})`);
      return;
    }

    if (cmd === 'scan-git') {
      if (positional.length === 0) {
        printUsage();
        process.exitCode = 1;
        return;
      }
      const result = scanGit(positional, {
        since: flags.since,
        out: flags.out,
        withDiffStat: Boolean(flags['with-lines']),
      });
      console.log(`Scanned ${positional.length} repo(s) -> ${result.outPath}`);
      return;
    }

    printUsage();
    process.exitCode = cmd ? 1 : 0;
  } catch (err) {
    console.error(`pspt: ${err.message}`);
    process.exitCode = 1;
  }
}

main();
