# pspt-lang — the `.pspt` DSL

A small, hand-written language for authoring product specs (docx) and project
trackers (xlsx) without writing JavaScript. It mirrors the shape the underlying
SDKs already use (sections, tables, key/value fields) rather than inventing new
concepts — there are no loops, conditionals, or imports by design.

`.pspt` source compiles to plain, readable JavaScript that calls
[`pspt-docx`](../pspt-docx/README.md)'s or [`pspt-xlsx`](../pspt-xlsx/README.md)'s
existing setters. The generated file is never meant to be hand-edited — re-run
`pspt compile` instead (see [pspt-cli](../pspt-cli/README.md)).

Every parse/lex error includes a 1-based line and column number, since `.pspt`
authors include non-developers and AI agents, not just engineers.

## Two document types, one grammar

The first line of every `.pspt` file declares the document title and, optionally,
its type:

```
doc "My Document Title" type=docx
```

- `type=docx` (default if omitted) compiles to a `pspt-docx` `ProductSpecSDK` script.
- `type=xlsx` compiles to a `pspt-xlsx` `ExcelTrackerSDK` script.

The rest of the grammar differs depending on which type you picked (docx uses
`section`/`table`/`rows`/`list`; xlsx uses `calendar`/`group`/`task`/`callout`) —
see the two sections below.

## Comments

`#` and `//` both start a line comment (rest of the line is ignored):

```
# a full-line comment
doc "My Doc" type=docx  // trailing comment
```

## Values

Every field/attribute value is one of:

| Kind | Example | Notes |
|---|---|---|
| string | `"Hello, \"world\"\nline two"` | `\"`, `\\`, `\n` escapes supported. Must stay on one line (unterminated at end-of-line is an error). |
| number | `42`, `-1`, `2.5` | |
| date | `2026-07-22` | Must match `YYYY-MM-DD` exactly and not be followed by more identifier characters. |
| bool | `true` / `false` | |
| null | `null` | |
| bare ident | `bareword`, `1B365D` (hex color) | Treated as a string value. Used for color names/hex codes and unquoted words. |

## Docx documents

### Structure

```
doc "Product Spec" type=docx

hardware: true   # optional — turns on the Hardware Specification section

section <name> {
  <field> : <value>
  table <name> { <col>: "<Header>" [w<weight>]  ... }
  rows <name> [ { <field>: <value>, ... }, ... ]
  list <name> { item "<title>" { item "<title>" ... } }
}
```

You can have as many `section { ... }` blocks as you like; they're purely an
authoring convenience — the codegen looks at each `field`/`table`/`rows`/`list`
inside every section for a matching setter, not the section name itself. Grouping
related fields into sections just keeps the source readable.

### Free-text and scalar fields

A bare `field: value` line maps directly onto the matching `ProductSpecSDK`
setter, using the exact same data-key names as `ProductSpecSDK`'s internal
`this.data` object (which mirror the setter names — see
[pspt-docx's README](../pspt-docx/README.md) for the full field-name list):

```
section overview {
  executiveSummary: "Acme Widget is a widget-connecting widget."
}
```

compiles to:

```js
doc.setExecutiveSummary("Acme Widget is a widget-connecting widget.");
```

If a field name has no matching setter, it's skipped with a `// WARNING:`
comment in the generated JS rather than failing the whole compile — check for
these after every `pspt compile`.

**Image/diagram path fields** (`logoImagePath`, `diagramImagePath`, etc.) are
plain string fields — pass a real file path and the SDK embeds it, omit it (or
pass a path that doesn't exist) and the SDK draws a placeholder box. No special
grammar is needed for these.

**Known limitation:** fields whose setter expects an *object* payload (e.g.
`setCoverPage({productName, shortDescription, lastUpdated, status})`,
`setApis({rows, docsLink})`, `setPricingModel({modelDescription, tiers})`) are
not yet expressible as a nested object literal in this grammar version — a bare
`cover: "..."` field maps to one of these, and the compiler will emit a
`// WARNING:` and skip it rather than silently passing a broken shape. Use the
direct SDK API (see [pspt-docx](../pspt-docx/README.md)) for these fields until
nested-field DSL syntax is added.

### Tables

```
section business {
  table features {
    name:        "Feature Name" w1
    description: "Description"  w2
    priority:    "Priority"     w1
  }
  rows features [
    { name: "Auth", description: "OAuth2 login", priority: "Must" },
    { name: "Dashboard", description: "Usage overview", priority: "Should" },
  ]
}
```

- `table <name> { key: "Header" [w<weight>] ... }` declares the columns. The
  weight suffix (`w1`, `w2.5`, `w0.5`) sets the relative column width — omit it
  to default to `1`. It must immediately follow the header string with no comma.
- `rows <name> [ {...}, {...} ]` supplies the row data — a trailing comma after
  the last row is allowed.
- The table `name` must match a docx data key with a known array-typed setter
  (`features`, `productRoadmap`, `targetMarket`, `technologyStack`, etc. — see
  [pspt-docx](../pspt-docx/README.md) for the complete list). An unrecognized
  table name produces a `// WARNING:` and is skipped.
- **A `rows` block with no matching `table` of the same name in the same
  section is also skipped with a `// WARNING:`** — always check for this if
  your table data doesn't seem to show up; it usually means a typo in the name.
- Omitting `rows` entirely for a declared table is fine — the underlying SDK
  renders its placeholder row automatically when no real data is supplied.

### Nested numbered lists (Security section)

The `setSecurity({measures, notes})` payload is expressed with a `list` block
named `measures`, plus an optional sibling `notes:` field:

```
section technical {
  list measures {
    item "Cloudflare" {
      item "WAF"
      item "Turnstile"
    }
    item "Cloud Armor" {
      item "Rate limiting"
    }
  }
  notes: "Data encrypted at rest and in transit; secrets stored in a managed vault."
}
```

- Each top-level `item "..."` becomes a numbered entry (`1.`, `2.`, ...).
- A nested `{ item "..." ... }` block becomes the lettered sub-list (`a.`, `b.`,
  ...) under that entry.
- Nesting is only rendered one level deep by the underlying `numberedBlock()`
  renderer — a third nesting level parses fine but is flattened into the same
  lettered sub-list rather than gaining its own further indent.
- `notes:` is consumed as part of the `setSecurity` call and is *not* also
  emitted as a separate field — this is intentional, not a bug.
- A `list` block with any name other than `measures` has no known consumer yet
  and is skipped with a `// WARNING:`.

### Hardware section toggle

```
doc "My Device" type=docx
hardware: true
```

Must appear directly after the `doc` line (before any `section` blocks). Only
`true`/`false` are valid — any other value type is a parse error. Omitting the
line entirely leaves hardware off (the `ProductSpecSDK` default).

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
above (hardware toggle, tables+rows, nested security list, xlsx calendar +
multiple groups + callout).

[`examples/fixtures/edge-cases/`](../../examples/fixtures/edge-cases/) has ~40
additional fixtures covering edge cases (empty sections, decimal weights, hex
colors, CRLF line endings, unicode/escapes, deliberately invalid syntax to see
the exact error messages, etc.) if you want to see the grammar's exact
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
