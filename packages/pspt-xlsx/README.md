# pspt-xlsx

`ExcelTrackerSDK` generates an Excel (.xlsx) "Project Tracker & Gantt Chart"
spreadsheet: a left-side task table plus a right-side day-by-day Gantt timeline
with month banners, weekend highlighting, colored task bars, and commit-count
badges. It's a lift-and-adapt port of the original `excel-tracker-sdk.js` —
every method and rendering behavior is preserved; only the shared color
tokens/helpers now come from [`pspt-core`](../pspt-core).

## Install

```js
const ExcelTrackerSDK = require('pspt-xlsx');
```

## Quick start

```js
const ExcelTrackerSDK = require('pspt-xlsx');

async function main() {
  const tracker = new ExcelTrackerSDK();

  tracker.setTitle('Project Tracker & Gantt Chart');
  tracker.setCalendarRange('2026-06-01', '2026-08-31');

  tracker.addSection({
    title: 'Backend Core Tasks',
    bannerColor: '2D7D46', // dark green
    rowColor: 'yellow',
    tasks: [
      {
        no: 1, name: 'Auth Middleware', detail: 'JWT + refresh tokens',
        startDate: '2026-06-11', endDate: '2026-06-12',
        deliverables: 'Auth Middleware', commitCount: 3, lines: '+120 / -30',
      },
    ],
  });

  await tracker.generate('./project-tracker-gantt.xlsx');
}

main();
```

All setters/`addSection`/`addCallout` return `this`, so they can be chained.

## Setter reference

| Method | Purpose |
|---|---|
| `setTitle(title)` | Sets the workbook title / worksheet name. |
| `setCalendarRange(startDate, endDate)` | Explicit calendar bounds (e.g. `'2026-06-01'`, `'2026-08-31'`). If never called, the range is computed automatically from the min/max task dates. |
| `setMinMonthsSpan(count = 3)` | Minimum number of months the calendar should span, even if task dates cover less. |
| `setExpandFullMonths(enabled = true)` | Whether to auto-expand the calendar to the 1st/last day of its start/end months. |
| `addSection({title, bannerColor?, rowColor?, tasks})` | Adds one banner-row group of tasks. See below for the task shape and color options. |
| `addCallout({dateStr, text, color?})` | Records a callout on a specific date. **Currently accepted but not rendered anywhere** — see [Known gap](#known-gap-addcallout-is-not-rendered) below. |
| `generate(filePath)` | Renders and writes the `.xlsx` file to disk (async). |

## Colors

Both `bannerColor` and `rowColor` (on `addSection`) and `color` (on
`addCallout`) accept either:
- a named preset: `green`, `yellow`, `pink`, `blue`, `purple`, `gray` (these are
  the light row-background presets), or
- a raw hex string, with or without a leading `#` (e.g. `'2D7D46'`, `'#2D7D46'`).

Design tokens for the fixed chrome (navy header bar, weekend red, Gantt bar
blue, grid border) live in `pspt-core` and aren't user-configurable — they
match the original template exactly.

## Task fields

Passed as objects inside `addSection({tasks: [...]})`. Every field has a
default if omitted:

| Field | Default | Notes |
|---|---|---|
| `no` | — | Row number, any value you want to display. |
| `name` | `''` | Task name (left table, "FEATURE NAME" column). |
| `detail` | `''` | "FEATURE DETAIL" column. |
| `checklist` | `'Completed'` | Free-text status column. |
| `mandays` | `1` | Numeric. |
| `total` | `mandays` | Numeric, defaults to `mandays` if omitted. |
| `startDate` / `endDate` | — | Parsed via a lenient date parser — accepts `'YYYY-MM-DD'` strings or `Date` objects. Drives the Gantt bar's date range. |
| `deliverables` | `''` | Free text. |
| `notes` | `''` | Free text, **or** `{ text, hyperlink }` for a clickable link (rendered blue/underlined). |
| `commitCount` | `1` | Shown as a badge on the first day of the task's Gantt bar. |
| `lines` | `'+0 / -0'` | Shown in the "Lines (+/-)" column and the Gantt bar's tooltip note. Can instead be derived automatically by passing `linesAdded`/`linesRemoved` (see below). |
| `check` | `true` | Rendered as ☑/☐ in the "Check" column. |

`lines` can also be supplied indirectly:

```js
{ ..., linesAdded: 120, linesRemoved: 30 }  // -> lines: '+120 / -30'
```

## Getting real commit/line-change data

`commitCount`/`lines` don't have to be typed by hand — `pspt-cli`'s
`scan-git --with-lines` command can compute them from actual git history (via
`git show --shortstat` per commit) and `pspt-core` exposes the same logic
directly:

```js
const { getCommitStats } = require('pspt-core');

const stats = getCommitStats('/path/to/repo', ['abc123', 'def456']);
// -> { commitCount: 2, insertions: 145, deletions: 38, lines: '+145 / -38' }

tracker.addSection({
  title: 'Backend',
  tasks: [{ name: 'Auth API', startDate: '2026-01-05', endDate: '2026-01-10', ...stats }],
});
```

See [pspt-cli's scan-git docs](../pspt-cli/README.md#scan-git) for the CLI form.

## Known gap: `addCallout` is not rendered

`addCallout({dateStr, text, color})` stores the callout on `this.callouts`, but
nothing in `generate()` actually draws it onto the worksheet. This is a
pre-existing gap carried over unchanged from the original `excel-tracker-sdk.js`
— not something introduced by this port. If you need callouts to actually
appear, this is open follow-up work (would need to overlay text/highlighting on
the relevant date column during `generate()`).

## Output

```js
await tracker.generate('/path/to/output.xlsx');
```

There is no in-memory buffer accessor (unlike `pspt-docx`'s `toBuffer()`) —
`generate()` always writes directly to disk via ExcelJS's `workbook.xlsx.writeFile()`.

## Using it via the `.pspt` DSL instead

If you'd rather author trackers in a terser, non-JS format, see
[`pspt-lang`](../pspt-lang/README.md) — its `calendar`/`group`/`task`/`callout`
syntax compiles directly into `setCalendarRange`/`addSection`/`addCallout` calls
against this SDK.
