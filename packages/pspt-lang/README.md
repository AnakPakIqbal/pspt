# pspt-lang — the `.pspt` DSL

A small, hand-written language for authoring product documentation (docx) and
project trackers (xlsx) without writing JavaScript. It mirrors the shape the
underlying SDKs already use (sections, tables, key/value fields) rather than
inventing new concepts — there are no loops, conditionals, or imports by design.

`.pspt` source compiles to plain, readable JavaScript that calls
[`pspt-docx`](../pspt-docx/README.md)'s or [`pspt-xlsx`](../pspt-xlsx/README.md)'s
existing setters. The generated file is never meant to be hand-edited — re-run
`pspt compile` instead (see [pspt-cli](../pspt-cli/README.md)).

Every parse/lex error includes a 1-based line and column number, since `.pspt`
authors include non-developers and AI agents, not just engineers.

## Document types

The first line of every `.pspt` file declares the document title and its type:

```
doc "My Document Title" type=projectBrief
```

`type=` is **required** — there is no default. Valid values:

| `type=`                                                                                                                                                  | Compiles to                                               |
| -------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------- |
| `master`                                                                                                                                                 | A `MasterDocument` assembling all 15 Parts into one .docx |
| `picMatrix`                                                                                                                                              | The PIC Matrix / Documentation SOP front matter           |
| `styleGuide` `projectBrief` `brd` `prd` `srs` `techDoc` `uiux` `uat` `deploymentGuide` `userManual` `changelog` `changeRequestLog` `glossary` `appendix` | That single Part's SDK (Parts 1–14)                       |
| `xlsx`                                                                                                                                                   | A `pspt-xlsx` `ExcelTrackerSDK` script                    |

> **`type=docx` was removed.** It referred to the single-document
> `ProductSpecSDK`, which no longer exists — the compiler rejects it with a
> message naming the replacement. See
> [pspt-docx's migration table](../pspt-docx/README.md#migrating-from-productspecsdk).

The rest of the grammar depends on which type you picked — docx Parts use
`section`/`table`/`rows`/`list`/`object` (and `part` in a `master` file); xlsx
uses `calendar`/`group`/`task`/`callout`.

## Comments

`#` and `//` both start a line comment (rest of the line is ignored):

```
# a full-line comment
doc "My Doc" type=projectBrief  // trailing comment
```

## Values

Every field/attribute value is one of:

| Kind       | Example                          | Notes                                                                                                |
| ---------- | -------------------------------- | ---------------------------------------------------------------------------------------------------- |
| string     | `"Hello, \"world\"\nline two"`   | `\"`, `\\`, `\n` escapes supported. Must stay on one line (unterminated at end-of-line is an error). |
| number     | `42`, `-1`, `2.5`                |                                                                                                      |
| date       | `2026-07-22`                     | Must match `YYYY-MM-DD` exactly and not be followed by more identifier characters.                   |
| bool       | `true` / `false`                 |                                                                                                      |
| null       | `null`                           |                                                                                                      |
| bare ident | `bareword`, `1B365D` (hex color) | Treated as a string value. Used for color names/hex codes and unquoted words.                        |

## Docx documents

### Structure

```
doc "Project Brief" type=projectBrief

section <name> {
  <field> : <value>
  object <name> { <key>: <value>, ... }
  list   <name> { item "<text>" { item "<text>" ... } }
  table  <name> { <col>: "<Header>" [w<weight>]  ... }
  rows   <name> [ { <field>: <value>, ... }, ... ]
}
```

You can have as many `section { ... }` blocks as you like; they're purely an
authoring convenience — the codegen looks at each field/object/list/table inside
every section for a matching setter, not the section name itself. Grouping
related content into sections just keeps the source readable.

### Which construct to use

Every name — a field name, an `object`/`list`/`table` name — is a **data key**:
the setter name minus its `set` prefix, lowercased at the front. `overview:`
means `setOverview`, `table keyModules` means `setKeyModules`.

Which of the four constructs you write is decided by what the setter expects,
and the compiler knows that because it reads the Part SDK's own
`sectionGuide()` at compile time — the mapping cannot drift out of sync with
the SDK:

| Setter takes         | You write                            | Compiles to                       |
| -------------------- | ------------------------------------ | --------------------------------- |
| a string/number/bool | `overview: "..."`                    | `doc.setOverview('...')`          |
| an array of strings  | `list objectives { item "..." }`     | `doc.setObjectives(['...'])`      |
| an array of objects  | `table x { ... }` + `rows x [ ... ]` | `doc.setX([{...}])`               |
| a single object      | `object metadata { writer: "..." }`  | `doc.setMetadata({writer:'...'})` |

Use the wrong one and you get a `// WARNING:` comment naming the right one —
never a silently corrupted document:

```
// WARNING: 'metadata' (line 4) maps to setMetadata, which expects an object
// payload, not a plain value — skipped. Use `object metadata { key: value }`.
```

A name with no setter behind it is also skipped with a warning, and gets a
spelling suggestion when it's close to a real one (`did you mean 'overview'?`).
**Always check generated `.gen.js` files for `WARNING` comments after
compiling.**

To see every data key a Part accepts, ask its SDK directly:

```js
const { ProjectBriefSDK } = require('pspt-docx');
console.log(ProjectBriefSDK.sectionGuide());
```

### Reserved words as data keys

Data keys come from the SDKs, which know nothing about this grammar's reserved
words — `setNamingConventions` really does take an `item` key, and a table
column can legitimately be called `type` or `status`. So a keyword is treated
as a plain name wherever a name is expected (field names, table and column
names, row keys, object keys), and a `<keyword>: <value>` line is always read as
a field rather than the start of a block:

```
table namingConventions {
  item:       "Item"       w1
  convention: "Convention" w1
}
```

### Fields and object blocks

```
section brief {
  overview: "Acme Widget lets small teams submit expenses from their phones."

  object headerFooterLabels { productNameLabel: "Acme Widget" }
  object metadata { writer: "Jane Doe", status: "Draft", version: "V1", lastUpdate: "Sept 1, 2026" }
}
```

compiles to:

```js
doc.setOverview('Acme Widget lets small teams submit expenses from their phones.');
doc.setHeaderFooterLabels({ productNameLabel: 'Acme Widget' });
doc.setMetadata({ writer: 'Jane Doe', status: 'Draft', version: 'V1', lastUpdate: 'Sept 1, 2026' });
```

Commas between `object` entries are optional. Object values are scalars only —
there is no nested-object-inside-object syntax; the handful of setters wanting
deeper nesting (`BrdSDK.setRaci`, `TechnicalDocumentationSDK.setEndpoints`,
`UserManualSDK.setFeatureWalkthroughs`) still need the direct SDK API.

**Every Part you generate wants `object headerFooterLabels { productNameLabel: "..." }`** —
it's what fills that Part's running header. Skip it and the header reads
"Product name/logo".

### Lists

```
section brief {
  list objectives {
    item "Cut expense processing time by 70%"
    item "Achieve 90% team adoption within Q4"
  }
}
```

becomes `doc.setObjectives([...])`. These setters take flat lists of strings, so
a nested `{ item "..." }` block parses fine but its children are dropped with a
warning saying so.

### Tables

```
section brief {
  table keyModules {
    module:   "Module"        w1
    features: "Core Features" w2.5
  }
  rows keyModules [
    { module: "Receipts",  features: "Photo capture + OCR" },
    { module: "Approvals", features: "One-tap approve/reject" },
  ]
}
```

- `table <name> { key: "Header" [w<weight>] ... }` declares the columns. The
  weight suffix (`w1`, `w2.5`, `w0.5`) sets the relative column width — omit it
  to default to `1`. It must immediately follow the header string with no comma.
- `rows <name> [ {...}, {...} ]` supplies the row data — a trailing comma after
  the last row is allowed.
- A `table` is paired with the **next** `rows` block of the same name, so two
  sections can each declare a `table keyModules` without their rows crossing.
- **A `rows` block with no matching `table` is skipped with a `// WARNING:`** —
  check for this if your table data doesn't show up; it usually means a typo.
- Omitting `rows` entirely for a declared table is fine — the SDK renders its
  own placeholder row when no real data is supplied.

### Nested values inside a row

A row cell can itself be a `[list]` or an `{object}`. Several setters take
payloads that need it — an endpoint's request and response body, an entity's
column list, a user story's Given/When/Then bullets:

```
rows endpoints [
  {
    method: "POST",
    path: "/tasks",
    description: "Creates a task.",
    requestBody: [ "{", "  \"name\": \"string\"", "}" ],
    responseBody: [ "{ \"id\": \"uuid\" }" ],
    errors: "Errors: 422 on validation failure."
  }
]

rows entityDetails [
  {
    tableName: "tasks",
    columns: [
      { column: "id", type: "CHAR(36)", constraints: "PK", description: "Identifier" }
    ]
  }
]
```

Both forms nest arbitrarily and allow a trailing comma.

### Authoring the whole master document

In a `type=master` file, wrap each Part's content in a `part <key> { ... }`
block. The body is exactly the same grammar as a section body, so a Part can be
lifted into its own `type=<key>` file (or folded back in) without rewriting it:

```
doc "Acme Widget — Product Documentation" type=master

part projectBrief {
  object headerFooterLabels { productNameLabel: "Acme Widget" }
  overview: "Acme Widget lets small teams submit and approve expenses from their phones."
}

part glossary {
  table terms {
    term:       "Term"       w1
    definition: "Definition" w3
  }
  rows terms [
    { term: "OCR", definition: "Optical Character Recognition" }
  ]
}
```

Parts you leave out still render, as their own blank template, so this produces
a complete document rather than a document with holes in it. A `part` block in a
single-Part file — or a `section` sitting outside every `part` in a master file —
is skipped with a `// WARNING:`.

### Legacy `hardware:` flag

`hardware: true` still parses (it toggled the old `ProductSpecSDK`'s hardware
section) but has no effect — the documentation suite has no hardware section.
The compiler emits a `// WARNING:` telling you to delete the line.

## Xlsx trackers

### Structure

```
doc "Sprint Tracker" type=xlsx
calendar <start-date> .. <end-date>   # optional — auto-computed from tasks if omitted

group "<Section Title>" [color=<name-or-hex>] [rowColor=<name-or-hex>] {
  task "<Task Name>" start=<date> end=<date> [<key>=<value> ...]
}

callout <date> "<text>" [color=<name-or-hex>]
```

- `calendar 2026-01-01 .. 2026-06-30` maps to `setCalendarRange(...)`. If
  omitted, the SDK computes the range from the min/max task dates itself.
- `group "<title>"` maps to one `addSection({title, bannerColor, rowColor, tasks})`
  call. `color=` sets the banner color; `rowColor=` sets the row background —
  both accept either a named preset (`green`, `yellow`, `pink`, `blue`, `purple`,
  `gray`) or a raw hex code (`1B365D`, `00FF00` — the `#` is optional and may be
  omitted).
- `task "<name>" start=... end=...` requires at least `start`/`end`; any other
  `key=value` pairs pass straight through into the task object with the same
  keys the SDK expects, with two renames handled automatically: `start` →
  `startDate`, `end` → `endDate`. Everything else (`mandays`, `detail`,
  `deliverables`, `notes`, `commitCount`, `lines`, `check`, ...) is passed
  through as-is — see [pspt-xlsx](../pspt-xlsx/README.md#task-fields) for the
  full field list and what each one does.
- `callout <date> "<text>" [color=...]` maps to `addCallout({dateStr, text, color})`.
  **Note:** as of writing, callouts are accepted by the underlying SDK but never
  actually rendered onto the worksheet — this is a pre-existing gap in the
  original `ExcelTrackerSDK`, not something this DSL introduces or fixes.

### Manually supplying commit/line-change data

There's no separate grammar for git stats — you write them as ordinary task
attributes, in the same place as `mandays=`:

```
task "Auth API" start=2026-01-05 end=2026-01-10 mandays=5 commitCount=7 lines="+120 / -30"
```

- `commitCount=<number>` shows as the badge on the first day of that task's
  Gantt bar.
- `lines="+N / -M"` (a plain string, in exactly that format) shows in the
  tooltip note and the "Lines (+/-)" column.

To generate these values from real git history instead of typing them by hand,
run `pspt scan-git --with-lines` first and copy the relevant numbers in — see
[pspt-cli's scan-git docs](../pspt-cli/README.md#scan-git).

## Full worked examples

See [`examples/fixtures/sample-docx.pspt`](../../examples/fixtures/sample-docx.pspt)
and [`examples/fixtures/sample-xlsx.pspt`](../../examples/fixtures/sample-xlsx.pspt)
in the repo root for complete, runnable files exercising most of the grammar
above (a `type=master` file filling four Parts with object blocks, lists and
tables+rows; xlsx calendar + multiple groups + callout).

[`examples/fixtures/edge-cases/`](../../examples/fixtures/edge-cases/) has ~50
additional fixtures covering edge cases (empty sections, decimal weights, hex
colors, CRLF line endings, unicode/escapes, every warning path, and deliberately
invalid syntax to see the exact error messages) if you want to see the grammar's exact
boundaries.

## Compiling and running

```bash
pspt compile my-file.pspt              # emits my-file.gen.js next to it
pspt build my-file.pspt --out out.docx # compile + run in one step
```

See [pspt-cli's README](../pspt-cli/README.md) for the full command reference.

## Grammar reference (informal EBNF)

```
program       := doc-decl (calendar-decl)? (hardware-decl)? statement*
doc-decl      := 'doc' string ('type' '=' ident)?
calendar-decl := 'calendar' date '..' date
hardware-decl := 'hardware' ':' bool

statement     := section | group | callout | table-decl | rows-decl | field

section       := 'section' ident '{' section-body* '}'
section-body  := field | table-decl | rows-decl | list-decl
field         := ident ':' value
list-decl     := 'list' ident '{' item* '}'
item          := 'item' string ('{' item* '}')?

table-decl    := 'table' ident '{' column* '}'
column        := ident ':' string weight?
weight        := ident            (e.g. 'w1', 'w1.5' — parsed from ident text)
rows-decl     := 'rows' ident '[' row (',' row)* ','? ']'
row           := '{' kv (',' kv)* ','? '}'
kv            := ident ':' value

group         := 'group' string ('color' '=' ident)? '{' (task|callout)* '}'
task          := 'task' string kvpair*
callout       := 'callout' date string ('color' '=' ident)?
kvpair        := ident '=' value

value         := string | number | date | bool | null
```
