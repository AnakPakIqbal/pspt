'use strict';

const js = require('@eslint/js');
const eslintConfigPrettier = require('eslint-config-prettier');
const nPlugin = require('eslint-plugin-n');
const perfectionist = require('eslint-plugin-perfectionist');
const security = require('eslint-plugin-security');
const sonarjs = require('eslint-plugin-sonarjs');
const globals = require('globals');

// pspt is plain JavaScript (no TypeScript, no tsconfig.json) — this config is
// the JS-only subset of the backend linting playbook. Every @typescript-eslint
// rule (no-floating-promises, no-unsafe-*, switch-exhaustiveness-check,
// return-await, etc.) needs a type-checker and is intentionally omitted; if
// this repo ever migrates to TS, reinstate those alongside typescript-eslint.

const MAX_NESTED_DEPTH = 3;
// 2 and 10 show up constantly as JSON.stringify's indent arg, string
// slice/substring offsets, and small loop/format constants across this
// codebase — flagging every occurrence added noise, not signal.
const NO_MAGIC_NUMBERS_IGNORE = [0, 1, -1, 2, 10];

module.exports = [
  {
    ignores: [
      '**/node_modules/**',
      '**/*.gen.js',
      'examples/output/**',
      'examples/fixtures/edge-cases/**/*.gen.js',
    ],
  },

  js.configs.recommended,
  security.configs.recommended,
  nPlugin.configs['flat/recommended'],
  sonarjs.configs.recommended,

  {
    plugins: { perfectionist },
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'commonjs',
      globals: { ...globals.node },
    },
    rules: {
      // --- Import order: builtin -> external -> relative (./) ---
      'perfectionist/sort-imports': [
        'error',
        {
          type: 'natural',
          order: 'asc',
          ignoreCase: true,
          newlinesBetween: 1,
        },
      ],

      // --- Security ---
      'security/detect-object-injection': 'off', // too many false positives on plain object/array indexing in this codebase
      'security/detect-non-literal-fs-filename': 'error',
      'security/detect-eval-with-expression': 'error',
      'security/detect-non-literal-regexp': 'error',
      'security/detect-unsafe-regex': 'error',
      'security/detect-pseudoRandomBytes': 'error',
      'no-eval': 'error',
      'no-new-func': 'error',

      // --- Node correctness ---
      'n/no-missing-require': 'error',
      'n/no-process-exit': 'error',

      // --- Redundant code ---
      'sonarjs/no-identical-functions': 'error',
      'sonarjs/no-duplicate-string': ['error', { threshold: 4 }],
      'sonarjs/no-collapsible-if': 'error',
      'sonarjs/no-redundant-boolean': 'error',

      // --- General hygiene ---
      eqeqeq: ['error', 'always', { null: 'ignore' }], // `x != null` is the standard idiom for "null or undefined" in one check; requiring `x !== null && x !== undefined` everywhere would hurt readability for no safety gain
      'no-console': 'off', // this is a CLI tool; console output is the product, not debug noise
      'no-implicit-coercion': 'error',
      'prefer-const': 'error',
      'no-var': 'error',

      // --- Complexity / readability ---
      'max-depth': ['error', MAX_NESTED_DEPTH],
      'no-nested-ternary': 'error',
      'id-length': [
        'error',
        {
          min: 3,
          exceptions: ['_', 'i', 'j', 'k', 'fs', 'id', 'js', 'ch', 'n', 'r', 'p', 'v', 't'],
          properties: 'never',
        },
      ],
      'no-magic-numbers': [
        'error',
        {
          ignore: NO_MAGIC_NUMBERS_IGNORE,
          ignoreArrayIndexes: true,
          ignoreDefaultValues: true,
          enforceConst: true,
          detectObjects: false,
        },
      ],
    },
  },

  {
    // The compiled-JS-emitting codegen legitimately builds up JS source as
    // string literals containing repeated punctuation/keywords — sonarjs's
    // duplicate-string check produces noise here, not real bugs.
    files: ['packages/pspt-lang/src/codegen.js'],
    rules: {
      'sonarjs/no-duplicate-string': 'off',
    },
  },

  {
    // Repo build scripts: they read/write paths derived from argv and from the
    // workspace layout, and exit non-zero to fail CI. "Non-literal fs filename"
    // and "no process.exit" are the intended behaviour here, same as for the
    // CLI's own entrypoints below.
    files: ['scripts/**/*.js'],
    rules: {
      'security/detect-non-literal-fs-filename': 'off',
      'n/no-process-exit': 'off',
      // The doc generator is a linear string builder: its 'complexity' is a long
      // sequence of lines.push() calls, not branching logic, so splitting it
      // further would scatter one document across several functions for no gain.
      'sonarjs/cognitive-complexity': 'off',
    },
  },

  {
    // CLI entrypoint scripts are allowed to exit the process directly.
    files: ['packages/pspt-cli/bin/**/*.js'],
    rules: {
      'n/no-process-exit': 'off',
    },
  },

  {
    // Example/demo scripts (not library code) match the original ported
    // examples' own control flow, which calls process.exit() on failure —
    // preserved intentionally rather than "fixed" to avoid touching behavior
    // of files whose whole purpose is to mirror the original SDK usage.
    files: ['examples/*.js'],
    rules: {
      'n/no-process-exit': 'off',
      'sonarjs/no-duplicate-string': 'off', // sample task/row data ported verbatim from the original docs/example-*.js — repeated field values here are demo content, not a real DRY violation
    },
  },

  {
    // These files encode literal layout dimensions (DXA units, pixel sizes),
    // fixed spreadsheet column indices/widths, and design-token defaults —
    // the numbers ARE the content, not incidental magic values. Naming every
    // one as a constant (`const THREE = 3`) would add noise, not clarity.
    files: [
      'packages/pspt-docx/src/sdk/*.js',
      'packages/pspt-docx/src/index.js',
      'packages/pspt-xlsx/src/excel-tracker-sdk.js',
      'packages/pspt-core/src/tokens.js',
      'packages/pspt-core/src/helpers.js',
    ],
    rules: {
      'no-magic-numbers': 'off',
    },
  },

  {
    // git-scan.js intentionally invokes `git` via PATH (matching the original
    // ad hoc scan-commits.js/scan-agentic-repos.js scripts) and parses
    // `git show --shortstat` output with small, fixed, non-backtracking
    // regexes (`\d+\s+insertion` / `\d+\s+deletion` — no nested quantifiers,
    // cannot exhibit catastrophic backtracking) against a repo path the CLI
    // user supplies on the command line, by design.
    files: ['packages/pspt-core/src/git-scan.js'],
    rules: {
      'sonarjs/no-os-command-from-path': 'off',
      'sonarjs/super-linear-regex': 'off',
      'sonarjs/no-ignored-exceptions': 'off', // getCommitDiffStat intentionally degrades to {insertions:0, deletions:0} if `git show` fails (e.g. unknown hash) rather than aborting the whole scan
      'no-unused-vars': ['error', { caughtErrors: 'none' }], // the caught error is deliberately unused in these soft-fail branches
    },
  },

  {
    // pspt-cli's entire job is to read/require a file path the CLI user
    // names on the command line — "non-literal fs filename/require" is
    // exactly the intended behavior here, not an injection risk to fix.
    files: ['packages/pspt-cli/src/*.js'],
    rules: {
      'security/detect-non-literal-fs-filename': 'off',
      'security/detect-non-literal-require': 'off',
    },
  },

  {
    // The docx SDKs' generate() methods read/write whatever
    // output path and image path the calling code passes in — same
    // by-design "caller-supplied path" pattern as pspt-cli above.
    files: ['packages/pspt-docx/src/product-spec-sdk.js'],
    rules: {
      'security/detect-non-literal-fs-filename': 'off',
    },
  },

  {
    // The DSL's `w2.5`-style weight suffix regex (/^w\d+(\.\d+)?$/) is
    // anchored, has no nested/overlapping quantifiers, and cannot backtrack
    // catastrophically — flagged as unsafe only because the plugin doesn't
    // statically prove the anchor bounds.
    files: ['packages/pspt-lang/src/parser.js'],
    rules: {
      'security/detect-unsafe-regex': 'off',
    },
  },

  {
    // Hand-written lexer: a single `tokenize()` function that walks the
    // source char-by-char is the standard shape for this kind of code, and
    // splitting it apart purely to satisfy a complexity budget risks
    // introducing a real bug in security-sensitive parsing logic for no
    // behavioral benefit. The nested `while`/`if` depth here is inherent to
    // character-by-character scanning, not accidental nesting.
    files: ['packages/pspt-lang/src/lexer.js'],
    rules: {
      'sonarjs/cognitive-complexity': 'off',
      'max-depth': 'off',
    },
  },

  eslintConfigPrettier,
];
