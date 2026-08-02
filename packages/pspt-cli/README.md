# pspt-cli

The `pspt` command-line tool: compiles/runs [`.pspt` DSL files](../pspt-lang/README.md)
and scans git history for use in trackers.

## Running it

`pspt-cli` is published on npm, so the `pspt` command works globally once
installed:

```bash
npm install -g pspt-cli
pspt <command> [args]
# or, without installing:
npx pspt-cli <command> [args]
```

Working from inside this monorepo without installing the published package,
invoke the bin directly with `node` instead:

```bash
node packages/pspt-cli/bin/pspt.js <command> [args]
```

Both invocation styles are identical — the commands below apply either way.

## Commands

### `pspt compile`

```bash
pspt compile <file>.pspt
```

Compiles a `.pspt` source file to `<file>.gen.js` next to it (plain JS calling
`pspt-docx`/`pspt-xlsx` setters — see [pspt-lang](../pspt-lang/README.md) for
what the generated code looks like). Does **not** run it.

Prints `Compiled <file> -> <file>.gen.js (type=docx|xlsx)` on success.

On a lex/parse error, prints `pspt: <file>: Line N, col M: <message>` and exits
with a non-zero status — no partial `.gen.js` is written.

**Always check the generated `.gen.js` for `// WARNING:` comments** after
compiling — the DSL compiles successfully even when a field/table has no
matching setter, or a `rows` block doesn't match any declared `table`; those
cases are flagged with a warning comment rather than failing the build.

### `pspt build`

```bash
pspt build <file>.pspt --out <output-path>
```

Compiles the `.pspt` file (same as `compile`) and immediately executes its
generated `build()` function, writing the final `.docx`/`.xlsx` to `--out`.

If `--out` is omitted, defaults to `output.docx` or `output.xlsx` depending on
the file's declared `type=`.

```bash
pspt build examples/fixtures/sample-docx.pspt --out out/spec.docx
pspt build examples/fixtures/sample-xlsx.pspt --out out/tracker.xlsx
```

### `pspt scan-git`

```bash
pspt scan-git <repo-path...> [--since <date>] [--out <file>.json] [--with-lines]
```

Scans one or more git repositories and writes the results as a single JSON
file. Replaces the old ad hoc `docs/scan-commits.js` / `docs/scan-agentic-repos.js`
scripts from the original codebase.

| Flag             | Default         | Meaning                                                                                                                                                                                                                                                                           |
| ---------------- | --------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `--since <date>` | (all history)   | Passed straight to `git log --since=<date>` (e.g. `2026-01-01`).                                                                                                                                                                                                                  |
| `--out <file>`   | `scan-git.json` | Output JSON path.                                                                                                                                                                                                                                                                 |
| `--with-lines`   | off             | Also run `git show --shortstat` **per commit** to attach `{insertions, deletions}` to each entry. Costs one extra git process per commit — fine for tens/hundreds of commits, slow for very large histories. Omit it if you only need commit metadata (hash/date/author/subject). |

Multiple repo paths can be passed at once; each becomes its own entry in the
output array, named after the repo's directory basename:

```bash
pspt scan-git ../project-tracker ../project-tracker-api --since 2026-01-01 \
  --out scan.json --with-lines
```

**Output shape** (one object per repo):

```jsonc
{
  "name": "project-tracker",
  "url": "https://github.com/you/project-tracker", // from `git config remote.origin.url`, null if unavailable
  "path": "../project-tracker",
  "remote": "https://github.com/you/project-tracker",
  "branches": "  main\n  remotes/origin/main\n...",
  "commits": [
    {
      "hash": "5b2e629",
      "date": "2026-07-27",
      "author": "Jane Doe",
      "email": "<jane@example.com>",
      "subject": "refactor: remove unused assignees handling functions",
      "insertions": 3, // only present with --with-lines
      "deletions": 63, // only present with --with-lines
    },
  ],
  "error": null, // set instead of `commits` populating, if `git log` itself failed
}
```

**Using the results in a tracker**: copy the relevant `commitCount`/`lines`
values into your `.pspt` file's `task` attributes, or compute a summed total
across several commit hashes for one task with `pspt-core`'s `getCommitStats`:

```js
const { getCommitStats } = require('pspt-core');
const stats = getCommitStats('../project-tracker', ['5b2e629', '3915fe1']);
// -> { commitCount: 2, insertions: 27, deletions: 71, lines: '+27 / -71' }
```

See [pspt-xlsx's task-field docs](../pspt-xlsx/README.md#task-fields) for where
`commitCount`/`lines` land in the generated spreadsheet.

## Usage / help

Running `pspt` with no command, or an unrecognized one, prints full usage:

```bash
node packages/pspt-cli/bin/pspt.js
```

## Errors

All commands report failures as `pspt: <message>` on stderr and set a non-zero
exit code — this includes DSL lex/parse errors (with line/column), missing
required positional arguments (falls back to printing usage), and any runtime
error thrown while generating the final document (e.g. a filesystem error
writing the output path).
