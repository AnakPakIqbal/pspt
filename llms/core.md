# pspt-core

Shared design tokens, table/column-schema abstraction, placeholder convention,
render helpers, and git-scan utilities used by [pspt-docx](./docx-sdk.md),
[pspt-xlsx](./xlsx-sdk.md), and [pspt-cli](./cli.md). Not usually consumed
directly except for `getCommitStats` (see below) — everything else is
internal plumbing shared by the two document SDKs.

## Install

```js
const {
  getCommitStats,
  getCommitDiffStat,
  scanRepoCommits,
  scanRepoFull,
  scanRepos,
  distributeColumnWidths,
  resolveRows,
  docxBorderSet,
  applyXlsxBorder,
  parseDate,
  resetTime,
  formatDateDisplay,
  resolveColor,
  PLACEHOLDER_ROW_TEXT,
  DOCX_FONT,
  DOCX_COLOR,
  DOCX_SIZE,
  DOCX_PAGE,
  DOCX_CONTENT_WIDTH,
  XLSX_COLORS,
  XLSX_ROW_PRESETS,
} = require('pspt-core');
```

## Git-scan utilities (most likely to be used directly)

### `getCommitStats(repoPath, hashes)`

Sums insertions/deletions across one or more commit hashes and formats the
result as the `lines` string [pspt-xlsx](./xlsx-sdk.md#task-fields) tasks
expect.

```js
const { getCommitStats } = require('pspt-core');

getCommitStats('/path/to/repo', 'abc123');
// -> { commitCount: 1, insertions: 12, deletions: 3, lines: '+12 / -3' }

getCommitStats('/path/to/repo', ['abc123', 'def456']);
// -> { commitCount: 2, insertions: 145, deletions: 38, lines: '+145 / -38' }
```

`hashes` may be a single string or an array. Internally runs
`git show --shortstat --format="" <hash>` per hash via `getCommitDiffStat`,
which returns `{ insertions: 0, deletions: 0 }` (not a thrown error) if the
underlying git command fails for that hash.

### `scanRepoCommits(repoPath, opts?)`

Runs `git log --all --pretty=format:"%h|%ad|%an|<%ae>|%s" --date=short` in
`repoPath` and returns deduplicated, parsed commits.

```js
scanRepoCommits('/path/to/repo', { since: '2026-01-01', withDiffStat: true });
// -> [{ hash, date, author, email, subject, insertions?, deletions? }, ...]
```

| Option         | Effect                                                                                                            |
| -------------- | ----------------------------------------------------------------------------------------------------------------- |
| `since`        | Passed to `--since=<value>`.                                                                                      |
| `authorEmail`  | Keeps only commits whose `<email>` contains this substring.                                                       |
| `withDiffStat` | Also runs `git show --shortstat` per commit (one extra process each — slower on large histories, off by default). |

### `scanRepoFull(repoPath, opts?)`

Wraps `scanRepoCommits` with best-effort `remote` (from
`git config --get remote.origin.url`) and `branches` (from `git branch -a`)
lookups. Returns `{ path, remote, branches, commits, error }` — `remote`/
`branches` are `null` (not thrown) if those commands fail; `error` is set
(and `commits` left `[]`) only if `git log` itself fails.

### `scanRepos(repos, opts?)`

Batch form — `repos` is `Array<{name, path, url?}>`; returns one
`scanRepoFull` result per repo, each merged with `{ name, url }`. This is what
backs `pspt scan-git` (see [cli.md](./cli.md#pspt-scan-git)) when given
multiple repo paths.

## Design tokens

Extracted directly from the original `product-spec-sdk.js` /
`excel-tracker-sdk.js` templates so both SDKs render pixel-identical output to
the pre-port originals. **Not meant to be changed** unless the underlying
visual templates themselves change — there's no per-document override for
these beyond what each SDK's own setters expose (e.g. `bannerColor`/`rowColor`
on `addSection`, which resolve against `XLSX_ROW_PRESETS`).

- `PLACEHOLDER_ROW_TEXT` — `'[...]'`, the italicized placeholder shown for any
  unset docx field/table row.
- `DOCX_FONT` — `'Arial'`, used throughout the generated Word doc.
- `DOCX_COLOR` — hex colors (no `#`) for headings, table header bg/text,
  borders, cover-page placeholder text, confidential-footer red, etc.
- `DOCX_SIZE` — font sizes in half-points (Word's unit — 22 half-points =
  11pt) for each heading level, body text, header/footer.
- `DOCX_PAGE` — US-Letter page dimensions and margins in DXA (twentieths of a
  point).
- `DOCX_CONTENT_WIDTH` — usable table width in DXA, pre-computed as
  `DOCX_PAGE.width - left margin - right margin` (9360 dxa ≈ 6.5in).
- `XLSX_COLORS` — ARGB hex (with `FF` alpha prefix) for the navy header bar,
  dark-green section banner, red weekend column, blue Gantt bar, grid border.
- `XLSX_ROW_PRESETS` — the named color presets (`green`, `yellow`, `pink`,
  `blue`, `purple`, `gray`) accepted by `addSection({bannerColor, rowColor})`
  and `addCallout({color})` in [pspt-xlsx](./xlsx-sdk.md#colors), and by the
  DSL's `color=`/`rowColor=` attributes in [dsl.md](./dsl.md).

## Table/column-schema helpers

### `distributeColumnWidths(columns, totalWidth)`

Given `columns: Array<{key, header, weight?}>` and a `totalWidth`, returns an
array of per-column widths that sum exactly to `totalWidth` (weight defaults
to `1` if omitted; any rounding drift is absorbed into the last column). Used
by the docx table builder to turn `.pspt` `table { col: "Header" w2 }`
declarations into actual DXA column widths.

### `resolveRows(rows, placeholderRows)`

Resolves which row set to render for a table: returns
`{ hasData: boolean, rows }`, where `rows` is the real `rows` array if it's a
non-empty array, otherwise falls back to `placeholderRows`. Backs the
"unset table renders its placeholder row automatically" behavior described in
[dsl.md](./dsl.md#tables) and [docx-sdk.md](./docx-sdk.md).

## Render helpers

- `docxBorderSet(docx)` — given the `docx` package's exports, returns the
  standard 0.5pt black border set (`top`/`bottom`/`left`/`right`/
  `insideHorizontal`/`insideVertical`) applied to every table in
  [pspt-docx](./docx-sdk.md).
- `applyXlsxBorder(cell, gridBorderArgb?)` — applies the standard thin grid
  border to an ExcelJS cell; defaults to `XLSX_COLORS.gridBorder` if no color
  is given.
- `parseDate(dateVal)` — lenient date parser: passes through `Date` instances,
  parses `'YYYY-MM-DD'` strings as local-time (not UTC, avoiding off-by-one
  timezone shifts), falls back to `new Date(dateVal)` for anything else, and
  returns `new Date()` (today) if `dateVal` is falsy.
- `resetTime(date)` — returns a copy of `date` with the time set to
  `00:00:00.000`.
- `formatDateDisplay(date)` — formats a `Date` as `DD/MM/YYYY`; returns `''`
  for a falsy input.
- `resolveColor(colorInput, defaultHex, rowPresets?)` — resolves a color input
  (preset name or raw hex, `#` optional) into an ARGB hex string. Checks
  `rowPresets` (defaults to `XLSX_ROW_PRESETS`) first, then treats the input
  as a raw hex and prefixes `FF` if it's exactly 6 hex digits. Falls back to
  `defaultHex` (prefixing `FF` if not already present) when `colorInput` is
  falsy.

## Where this fits

You will rarely need to `require('pspt-core')` directly unless you're
computing git stats to feed into a tracker by hand (`getCommitStats`) or
scanning a repo outside of the CLI (`scanRepoFull`/`scanRepos`). Everything
else in this package exists so that [pspt-docx](./docx-sdk.md) and
[pspt-xlsx](./xlsx-sdk.md) share one visual-style source of truth rather than
each hard-coding their own copy of the original templates' colors/fonts/sizes.
