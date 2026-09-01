# Workflow: use pspt to document your own codebase

This is the step-by-step recipe for an agent that wants to use `pspt-cli` to
generate a product's own documentation (`.docx`) and its own project
tracker(s) (`.xlsx`) — i.e. self-documentation, not documenting `pspt` itself.

A product is often more than one repository (e.g. a frontend and a backend
repo side by side). This workflow always treats the _product_ as the unit for
the spec, and each _repository_ as the unit for the tracker.

**The deliverable is always a `.pspt` script, never hand-written JS calling
the SDK setters directly.** You author `.pspt` files under a shared `docs/`
folder (see "Where files live" below), then execute them with `pspt-cli`
(`pspt build`). The DSL and its grammar are documented in [dsl.md](./dsl.md)
— that's what you're writing.

Two independent outputs, two independent data-gathering rules:

1. **Documentation (.docx)** — exactly **one** combined document for the whole
   product, even when it spans multiple repos. It is a `type=master` `.pspt`
   file assembling the PIC Matrix front matter and all 14 Parts. Content comes from **reading
   each repo's actual source**, verified against real code, not guessed and
   not copied from `pspt`'s own docs (those only tell you which fields
   exist).
2. **Tracker (.xlsx)** — **one tracker per repository** — git history can't
   be meaningfully merged across repos, so a frontend repo and a backend repo
   each get their own tracker. Content comes from **real git history**,
   scanned oldest → newest, with real diff stats per commit — never inferred
   from commit message text. Commits are grouped **by feature, never by
   calendar month**, and each row covers one continuous run of work on that
   feature (summed `commitCount=`/`lines=`, one link to the most critical
   commit in `notes=`), listed oldest → newest.

## Step 0 — ask the user which folders make up the product

**Do not guess or assume from whatever happens to be in the current working
directory.** Before doing anything else, ask the user which folder(s)
correspond to which part of the product — e.g. "which folder is the
frontend, which is the backend, are there any other components (mobile,
infra, a shared package)?" A product may be one repo or several; get the
explicit list and each folder's role before proceeding.

Also confirm the **workspace root** — the parent directory containing all of
the folders the user just gave you. `docs/` (Step below) is created as a
sibling of those folders, under that shared parent, not inside any one of
them.

## Step 1 — install/confirm pspt-cli is available

```bash
npm install -g pspt-cli   # or: npx pspt-cli <command>
pspt                       # prints usage if installed correctly
```

See [cli.md](./cli.md) for full command reference.

## Where files live

All `pspt` artifacts live under `docs/pspt/` at the **workspace root** — the
shared parent of every repo/folder the user gave you in Step 0 — never inside
one of the individual repos:

```
<workspace-root>/
  frontend/               # one of the repos the user pointed you at
  backend/                # another one
  docs/pspt/
    documentation.pspt              # authored DSL source for the ONE combined product spec
    documentation.docx               # pspt build output
    trackers/
      frontend.pspt         # authored DSL source for frontend's tracker
      frontend.xlsx          # pspt build output
      backend.pspt           # authored DSL source for backend's tracker
      backend.xlsx            # pspt build output
```

For a single-repo product, this collapses to just `docs/pspt/documentation.pspt` +
`docs/pspt/trackers/<repo-name>.pspt`, still at the workspace root rather than
inside the repo.

Note what's _not_ in this tree: `pspt compile`/`pspt build` also produce a
`<file>.gen.js` next to each `.pspt` source (see Steps below) — that's a
disposable build intermediate, not a deliverable, and should be deleted after
its `// WARNING:` check rather than left sitting in `docs/pspt/`.

Create the directory if it doesn't exist yet (`pspt build` will not create
`--out`'s parent directory for you — see [cli.md](./cli.md#pspt-build)).

## Step 2 — Documentation (.docx): read every repo's code, not the docs, for content

One combined `documentation.pspt` covers the **whole product**, so gather content by
reading through **each** folder the user gave you in Step 0, not just one —
e.g. the technology-stack section should list both the frontend's and the
backend's stacks, the API section should describe backend endpoints alongside
how the frontend consumes them, distinguished clearly by repo/component
within each field's content.

**Write for a business stakeholder, not a fellow engineer.** This document is
read by people deciding on the product, not reviewing the code — say what a
feature _does for the user_ or _why a technical choice matters to the
business_, not just the implementation detail. "Users can reset a forgotten
password by email without contacting support" reads better than "implements
a POST /auth/reset-password endpoint with a JWT-signed reset token." Keep
concrete facts (real endpoint names, real table names, real dependency
versions) — accuracy still matters — but frame _why it matters_, don't just
transcribe the code.

### Go through every Part — do not stop at a handful

[docx-sdk.md](./docx-sdk.md) lists all 15 Parts and all 184 setters, each with
its data key and payload shape. Walk the Parts as a checklist, filling each one
from what the code actually says:

- **PIC Matrix** (front matter) — product name, type, status and the people
  named on the cover. Its PIC and storage-guide tables have sensible defaults;
  leave them unless your organisation genuinely differs.
- **Part 1 Style Guide** — leave untouched unless the product deliberately uses
  a different visual language from the template.
- **Part 2 Project Brief** — the elevator pitch, the pains that justify the work,
  measurable objectives, the module list (derive it from the real feature
  directories across every repo), a rough phase timeline, deliverables, early
  risks.
- **Part 3 BRD** — business-language requirements (BR-xxx), each traced to an
  objective; the actors and use cases the system actually supports; the
  traceability matrix. Derive requirements from what the code _does_, not from a
  README's marketing description — a documented feature with no matching code
  is not a shipped feature.
- **Part 4 PRD** — personas, MoSCoW-prioritised features, user stories, the
  single success-metrics table, the single Risk Register.
- **Part 5 SRS** — atomic, testable FR-xxx requirements traced back to BR-xxx;
  non-functional targets; the screen inventory (read the router/pages); the
  operating environment (read `package.json`, `Dockerfile`, engines).
- **Part 6 Technical Documentation + Data Model + API Spec** — the densest Part,
  and the one with the most real material available:
  - _Architecture and components_ → read the actual folder structure and entry
    points, not an architecture doc that may have drifted.
  - _Technology stack_ → each repo's `package.json` dependencies, framework
    files, `Dockerfile`, `docker-compose.yml`. Do not let one repo's stack
    stand in for another's.
  - _Codebase tree, naming conventions, data flow, integration points_ → read
    them; every one of these is verifiable.
  - _Security_ → real auth middleware, validation, rate-limiting config, secrets
    handling. Report what is implemented, not a generic checklist. State plainly
    where something is absent.
  - _Data Model_ → the real schema (Prisma schema, SQL migrations, ORM models):
    entity list, per-entity columns with types and constraints, relationships,
    and the indexes that actually exist.
  - _API Specification_ → the real routers/controllers and any OpenAPI file:
    enumerate real endpoints with method, path, request and response shape, and
    error conditions. `setEndpoints` takes nested list cells for the bodies —
    use them; an API spec with an empty endpoint table is the single most common
    way this document comes out useless.
- **Part 7 UI/UX** — design tokens if a design system exists in the codebase
  (Tailwind config, a tokens file, Storybook, Figma links); otherwise say
  plainly that no formal design system was found.
- **Part 8 UAT** — real test directories and CI config: what is actually run
  (unit/integration/e2e, coverage), who owns each level, real environments.
- **Part 9 Deployment Guide** — the real `Dockerfile`/compose files, the real
  environment variables (read the config schema, not a sample `.env` that may
  have drifted), the real deploy and rollback commands.
- **Part 10 User Manual** — one walkthrough per user-facing feature, in plain
  language.
- **Parts 11–12 Changelog / Change Request Log** — from real release history if
  any exists; otherwise leave them as the blank template.
- **Part 13 Glossary** — the domain terms a reader would otherwise guess at.
- **Part 14 Appendix** — links to the living resources you read while writing.

Where you genuinely find nothing for a section after checking (e.g. no
pricing model exists because the product isn't monetized yet, or there's no
UI because it's a backend-only service), **say so explicitly in that
section** rather than silently skipping it — a stakeholder reading "no formal
pricing model has been defined yet" learns something; a blank/missing section
just looks incomplete.

**If a repo already has its own documentation** (a README,
`docs/architecture.md`, an OpenAPI file, etc.), treat it as a _lead_, not a
_source of truth_ — open the referenced code and confirm the doc still
matches reality before using it to fill a setter. Docs drift; code doesn't
lie about what it currently does.

Once you've gathered real values from every repo in scope, author
`docs/pspt/documentation.pspt` at the workspace root using the docx grammar (see
[dsl.md](./dsl.md#docx-documents)) — one `section` block per area above,
`field:`/`table`/`rows`/`list` entries filled with what you actually found in
the source, written in business language, noting which repo/component each
finding came from wherever the product spans more than one. Reference
[docx-sdk.md](./docx-sdk.md) only to know which field names exist and what
shape each expects (e.g. `table`+`rows` for `setFeatures`/`setApis`-style
array data, a `list` block for `setSecurity`'s nested measures).

Then compile and build it:

```bash
pspt compile docs/pspt/documentation.pspt          # writes docs/pspt/documentation.gen.js — check for // WARNING: comments first
pspt build docs/pspt/documentation.pspt --out docs/pspt/documentation.docx
```

**Always check the generated `.gen.js` for `// WARNING:` comments** — they
mean a field/table had no matching setter, or a `rows` block didn't match a
declared `table`, so that content silently didn't make it into the doc. Fix
the `.pspt` source and re-run, don't hand-edit the generated `.gen.js`.

`pspt compile`/`pspt build` always write `<file>.gen.js` next to the `.pspt`
source — there's no flag to put it elsewhere. Once you've checked it for
`// WARNING:` comments (and re-run after fixing any), **delete
`docs/pspt/documentation.gen.js`** — it's a disposable build intermediate, not a
deliverable. Only `documentation.pspt` (source) and `documentation.docx` (output) should remain
in `docs/pspt/`:

```bash
rm docs/pspt/documentation.gen.js
```

### Every setter is reachable from the DSL — there is no escape hatch

Older copies of this workflow described a "sanctioned exception" permitting
hand-written JS for setters the grammar could not express: object-payload
setters, and setters whose row shape used a column named after a DSL keyword.
**Both limits are gone.**

- Object payloads are written with an `object <name> { key: value }` block.
- Nested payloads — an endpoint's request/response body, an entity's column
  list, a user story's bullets — are written as `[list]` or `{object}` values
  inside a row cell.
- Keywords (`item`, `type`, `color`, `table`, …) are valid data-key names
  wherever a name is expected.

So the rule is now unconditional: **the deliverable is a `.pspt` script, built
with `pspt build`. Never hand-write JS against the SDKs.** If you believe a
setter is unreachable, you have the wrong data key or the wrong construct —
check [docx-sdk.md](./docx-sdk.md), or call `SDK.sectionGuide()` and read the
`example` payload, before reaching for JS.

And never paraphrase content into some other free-text field as a workaround.
Dumping API endpoints as prose into a notes field produces a document with an
empty endpoint table and a footnote pointing elsewhere — worse than calling the
real setter.

## Step 3 — Tracker(s) (.xlsx): one per repo, scan git history oldest → newest, real diff stats only

Every repo the user gave you in Step 0 gets **its own** `tracker.pspt` /
`tracker.xlsx` under `docs/pspt/trackers/<repo-name>.*` — git history from a
frontend repo and a backend repo is unrelated and must never be merged into
one tracker. Repeat this whole step once per repo in scope.

The tracker's `commitCount` and `lines` (+/-) fields must come from actual
`git diff` output, not from parsing/counting commit message text. Two ways to
get this, both produce the same underlying numbers:

### Option A — `pspt scan-git` (preferred, one command)

```bash
pspt scan-git <path-to-repo> --out /tmp/scan-git.json --with-lines
```

Run this once per repo (pass each repo's path in turn, or all of them at once
— `scan-git` accepts multiple `<repo-path...>` and produces one array entry
per repo, see [cli.md](./cli.md#pspt-scan-git)). `--with-lines` is required —
without it you only get commit metadata (hash/date/author/subject), no
insertions/deletions. The JSON output is scratch data for computing the
numbers below — no need to keep it around after you've pulled the stats you
need into that repo's `docs/pspt/trackers/<repo-name>.pspt`.

The output's `commits` array is already in the order `git log` produces it —
**reverse it (or sort by `date` ascending) before grouping into tasks**, since
`git log`'s default order is newest-first and this workflow requires
oldest → newest when assigning commits to tasks/timeline positions:

```js
const scan = require('/tmp/scan-git.json');
const chronological = [...scan[0].commits].reverse(); // oldest first
```

### Option B — plain git on the host system

If you're not scripting against the JSON output, the equivalent raw commands
are:

```bash
# oldest -> newest, one line per commit
git log --all --reverse --pretty=format:"%h|%ad|%an|<%ae>|%s" --date=short

# per-commit real diff stat (this is the +/- source of truth, not the subject line)
git show --shortstat --format="" <hash>
```

`--reverse` is what gives you oldest-first order directly from git itself.
Parse `git show --shortstat`'s `N insertion(s)`/`N deletion(s)` text for the
real `+`/`-` counts — this is exactly what `pspt-core`'s `getCommitDiffStat`
does internally (see [core.md](./core.md#git-scan-utilities-most-likely-to-be-used-directly)).

### Turning commits into task fields: group by feature, one row per run of work

Read every commit’s diff — not just its subject line. A commit titled "fix
bug" may touch the feature you are tracking, and a commit titled "feat: X" may
only be a stub. Which feature a commit belongs to is decided by **where its
lines landed**, weighting source over documentation, tests and lockfiles: a
commit that adds 2,000 lines of OpenAPI alongside 300 lines of feature code is
the feature, not the documentation.

**Group by feature, never by calendar month.** Each `group` is one feature or
area of the product — "Tasks", "Authentication", "Data Model" — so a reader
scanning the sheet sees everything that happened to Tasks in one block. A
tracker grouped into "June 2026", "July 2026" answers a question nobody asks.
Order the groups by how much work landed in each, biggest first.

**One row per continuous run of work on a feature, not one row per commit or
per day.** A row is a task; the commits inside it are its subtasks. A feature
built across four days is one piece of work, not four rows — splitting it by
calendar day is what makes a tracker read as noise. Let a run absorb a weekend
gap (roughly three quiet days) before starting a new row; a genuine pause in
the work starts the next one. `start=` is the run’s first commit date and
`end=` its last, so the Gantt bar shows the real shape of the work.

**Name the row for what shipped, and let `detail=` summarise the subtasks.**
The name is plain business language — "Auth Middleware & Invitation Service",
never the raw commit subject "feat: implement JWT middleware + invite-by-email
flow". Strip the conventional-commit prefix and the leading verb, and keep the
noun phrase. `detail=` then lists what each merged commit contributed, largest
first, which is the row’s summary of its own subtasks.

**List rows oldest → newest** within each group, and order the whole sheet by
the date work started — reuse the oldest-first ordering from `pspt scan-git`’s
reversed output (or `git log --reverse`); never leave rows in git’s native
newest-first order.

**Vary the banner `color=` per group, but keep every row the same `rowColor=`**
throughout the whole tracker. Each group’s banner (e.g. "Tasks (13 commits)")
gets its own distinct colour so features read as visually separate blocks,
while the task rows underneath every group share one consistent background
(e.g. `rowColor=green`). Do not do the opposite, and do not let banner colours
repeat back-to-back between adjacent groups.

For a merged row's `commitCount=`/`lines=`, **sum across every commit in the
group** — `pspt scan-git --with-lines`'s output (Option A above) gives you
each commit's `insertions`/`deletions` individually; add them up for the
group. With `pspt-core` directly, the equivalent is:

```js
const { getCommitStats } = require('pspt-core');

const stats = getCommitStats('<path-to-repo>', ['<hash-1>', '<hash-2>', '<hash-3>']);
// -> { commitCount: 3, insertions: X, deletions: Y, lines: '+X / -Y' } — summed across the group
```

**Put exactly one commit link in `notes=` — the most critical commit in the
group, not all of them.** "Most critical" means the commit with the largest
total line change (`insertions + deletions`) among the group — it's the best
single representative of what the group actually delivered. Build the link
from `scan-git`'s per-repo `url` (from `git config remote.origin.url`) plus
that one commit's hash: `<url>/commit/<hash>`. The DSL's `value` grammar now includes inline objects, so `notes=` can carry
the `{ text, hyperlink }` shape [xlsx-sdk.md](./xlsx-sdk.md#task-fields)
documents for clickable links — prefer it over a bare URL string:

```
notes={ text: "see commit", hyperlink: "https://github.com/you/repo/commit/861f55c" }
```

The summed numbers and the single most-critical link land directly as
`commitCount=`/`lines=`/`notes=` attributes on a `task` line in that repo's
`docs/pspt/trackers/<repo-name>.pspt` (see
[dsl.md](./dsl.md#manually-supplying-commitline-change-data)):

```
doc "Backend Tracker" type=xlsx
calendar 2026-06-01 .. 2026-08-31

group "June 2026: Auth & Invitations" color=green rowColor=green {
  task "Auth Middleware & Invitation Service" start=2026-06-11 end=2026-06-11 commitCount=3 lines="+11118 / -105" notes="https://github.com/you/backend/commit/861f55c"
}

group "July 2026: Billing" color=blue rowColor=green {
  task "Invoice PDF Export" start=2026-07-13 end=2026-07-13 commitCount=1 lines="+64 / -0" notes="https://github.com/you/backend/commit/7a1c9de"
}
```

(The first group merges three same-day auth/invitation commits into one row
— `commitCount=3` and `lines=` summed across all three, `notes=` pointing
only at whichever of the three had the largest total line change. The second
group is a single commit, so it's already one row. Note the two groups'
banners differ — `color=green` vs `color=blue` — while both use the same
`rowColor=green` for their task rows.)

Build it — repeat for every repo in scope, one `.pspt`/`.xlsx` pair each —
then delete each tracker's `.gen.js` the same way as the spec's, so only the
`.pspt` source and `.xlsx` output remain:

```bash
pspt compile docs/pspt/trackers/backend.pspt
pspt build docs/pspt/trackers/backend.pspt --out docs/pspt/trackers/backend.xlsx
rm docs/pspt/trackers/backend.gen.js

pspt compile docs/pspt/trackers/frontend.pspt
pspt build docs/pspt/trackers/frontend.pspt --out docs/pspt/trackers/frontend.xlsx
rm docs/pspt/trackers/frontend.gen.js
```

## Summary of the rule

| Output          | Scope                                                                                                                       | Authored as                           | Content source                                                                  | Never                                                                                                                                                                                                                                                                                                                       |
| --------------- | --------------------------------------------------------------------------------------------------------------------------- | ------------------------------------- | ------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `.docx` spec    | Whole product (all repos combined), every available section from [docx-sdk.md](./docx-sdk.md), written in business language | `docs/pspt/documentation.pspt`        | Real source code from every repo in scope (routes, schema, deps, config, tests) | Copying from READMEs/docs without verifying against code, skipping sections instead of stating "not found", paraphrasing a DSL-unreachable setter's content into an unrelated field, hand-written JS setter calls outside the narrow sanctioned exception, or leaving `documentation.gen.js` behind after the WARNING check |
| `.xlsx` tracker | One per repo, one `group` per feature, one `task` row per continuous run of work on it, oldest → newest                     | `docs/pspt/trackers/<repo-name>.pspt` | That repo's real per-file `git log --numstat` output, summed per merged run     | Merging git history across repos, grouping by calendar month instead of by feature, one row per raw commit, more than one commit link in `notes=`, inferring line counts or scope from commit message text, hand-written JS setter calls, or leaving `<repo-name>.gen.js` behind after the WARNING check                    |

Both are always run through `pspt-cli` (`pspt compile` / `pspt build`) — never
executed by hand-writing JS against `pspt-docx`/`pspt-xlsx` directly, except
for the docx spec's narrow set of DSL-unreachable setters (see "Setters the
DSL can't express" above), which is the only sanctioned exception. `docs/pspt/`
itself always lives at the workspace root (Step 0), never inside one of the
individual repos.

For the tracker specifically: one `group` per feature, ordered by how much
work landed in each; one row per continuous run of work on that feature, named
for what shipped with `detail=` summarising the commits under it; vary
`color=` per group while keeping `rowColor=` the same across the whole
tracker; put only the single most-critical commit's real
`<remote-url>/commit/<hash>` link in that row's `notes=`; and list every row
oldest → newest.
