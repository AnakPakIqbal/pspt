# The `.pspt` DSL

A small language for authoring product documentation (`.docx`) and project
trackers (`.xlsx`) without writing JavaScript. It mirrors the shape the SDKs
already use — sections, tables, key/value fields — rather than inventing new
concepts. There are no loops, conditionals or imports, by design.

`.pspt` source compiles to plain, readable JavaScript that calls
[`pspt-docx`](./docx-sdk.md)'s or [`pspt-xlsx`](./xlsx-sdk.md)'s existing
setters. The generated `.gen.js` is never meant to be hand-edited — re-run
`pspt compile` (see [cli.md](./cli.md)).

Every lex/parse error carries a 1-based line and column.

## Document types

The first line declares the title and the type. **`type=` is required — there
is no default.**

```
doc "My Document Title" type=projectBrief
```

| `type=`                                                                                                                                                  | Compiles to                                                 |
| -------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------- |
| `master`                                                                                                                                                 | A `MasterDocument` assembling all 15 Parts into one `.docx` |
| `picMatrix`                                                                                                                                              | The PIC Matrix / documentation-SOP front matter             |
| `styleGuide` `projectBrief` `brd` `prd` `srs` `techDoc` `uiux` `uat` `deploymentGuide` `userManual` `changelog` `changeRequestLog` `glossary` `appendix` | That single Part's SDK (Parts 1–14)                         |
| `xlsx`                                                                                                                                                   | A `pspt-xlsx` `ExcelTrackerSDK` script                      |

> **`type=docx` was removed.** It named the single-document `ProductSpecSDK`,
> which no longer exists. The compiler rejects it with a message listing the
> replacements. If you have older instructions telling you to write
> `type=docx`, `hardware: true`, `executiveSummary:` or `cover:`, they are
> stale — see [docx-sdk.md](./docx-sdk.md).

## Comments

`#` and `//` both start a line comment.

```
# a full-line comment
doc "My Doc" type=projectBrief  // trailing comment
```

## Values

| Kind       | Example                        | Notes                                                             |
| ---------- | ------------------------------ | ----------------------------------------------------------------- |
| string     | `"Hello, \"world\"\nline two"` | `\"`, `\\`, `\n` escapes. Must stay on one line.                  |
| number     | `42`, `-1`, `2.5`              |                                                                   |
| date       | `2026-07-22`                   | Exactly `YYYY-MM-DD`, not followed by more identifier characters. |
| bool       | `true` / `false`               |                                                                   |
| null       | `null`                         |                                                                   |
| bare ident | `bareword`, `1B365D`           | Treated as a string. Used for colour names and hex codes.         |
| list       | `[ "a", "b" ]`                 | Trailing comma allowed; nests.                                    |
| object     | `{ key: "value" }`             | Trailing comma allowed; nests.                                    |

## Reserved words are still valid names

Data keys come from the SDKs, which know nothing about this grammar's keywords
— `setNamingConventions` genuinely takes an `item` key, and a table column can
legitimately be called `type`, `status` or `color`. So a keyword is treated as
a plain name wherever a name is expected, and a `<keyword>: <value>` line is
always read as a field rather than the start of a block:

```
table namingConventions {
  item:       "Item"       w1
  convention: "Convention" w1
}
```

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

`section` blocks are an authoring convenience only — the compiler looks at each
field/object/list/table inside every section for a matching setter, not at the
section name. Group things however reads best.

### Choosing the right construct

Every name — a field name, an `object`/`list`/`table` name — is a **data key**:
the setter name minus `set`, lowercased at the front. `overview:` means
`setOverview`; `table keyModules` means `setKeyModules`.

Which construct you write is decided by what the setter expects, and the
compiler knows that because it reads the Part SDK's own `sectionGuide()` at
compile time. The mapping cannot drift from the SDK.

| Setter takes         | Write                                | Compiles to                       |
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

A name with no setter behind it is skipped with a warning too, and gets a
spelling suggestion when it is close to a real one. **Always check the
generated `.gen.js` for `WARNING` after compiling.**

[docx-sdk.md](./docx-sdk.md) lists every data key of every Part with its shape.
At runtime, `SDK.sectionGuide()` is the authority.

### Fields and object blocks

```
section brief {
  overview: "Acme Widget lets small teams submit expenses from their phones."

  object headerFooterLabels { productNameLabel: "Acme Widget" }
  object metadata { writer: "Jane Doe", status: "Draft", version: "V1", lastUpdate: "Sept 1, 2026" }
}
```

Commas between `object` entries are optional.

**Every Part you generate wants
`object headerFooterLabels { productNameLabel: "..." }`** — it is what fills
that Part's running header. Skip it and the header reads "Product name/logo".

### Lists

```
list objectives {
  item "Cut expense processing time by 70%"
  item "Achieve 90% team adoption within Q4"
}
```

These setters take flat lists of strings, so a nested `{ item "..." }` block
parses but its children are dropped with a warning saying so.

### Tables

```
table keyModules {
  module:   "Module"        w1
  features: "Core Features" w2.5
}
rows keyModules [
  { module: "Receipts",  features: "Photo capture + OCR" },
  { module: "Approvals", features: "One-tap approve/reject" },
]
```

- The weight suffix (`w1`, `w2.5`, `w0.5`) sets relative column width; omit it
  for `1`. It must follow the header string with no comma.
- A trailing comma after the last row is allowed.
- A `table` pairs with the **next** `rows` block of the same name, so two
  sections can each declare a `table keyModules` without their rows crossing.
- A `rows` block with no matching `table` is skipped with a warning — check for
  this if data doesn't show up; it usually means a typo.
- Omitting `rows` for a declared table is fine — the SDK renders its own
  placeholder row.

### Nested values inside a row

A row cell can itself be a `[list]` or an `{object}`. This is what makes the
richer setters reachable — an endpoint's request/response body, an entity's
column list, a user story's Given/When/Then bullets:

```
rows endpoints [
  {
    method: "POST",
    path: "/tasks",
    description: "Creates a task.",
    requestBody: [ "{", "  \"projectId\": \"uuid, required\"", "}" ],
    responseBody: [ "{ \"id\": \"uuid\" }" ],
    errors: "Errors: 422 on validation failure."
  }
]

rows entityDetails [
  {
    tableName: "tasks",
    columns: [
      { column: "id",         type: "CHAR(36)", constraints: "PK",      description: "Identifier" },
      { column: "project_id", type: "CHAR(36)", constraints: "FK",      description: "Owning project" }
    ]
  }
]
```

Both forms nest arbitrarily.

**The compiler checks these nested values, not just whether the field/table
name matches a setter.** Matching a name only proves the _top-level_ construct
is right — `table`+`rows` for a rows setter, `object` for an object setter.
What's _inside_ a row or object field can still be shaped wrong one or more
levels down: `setEntityDetails`' `columns` field wants a list of
`{column, type, constraints, description}` objects, and giving it a list of
plain strings used to compile clean and render blank cells with no warning
anywhere. It doesn't any more:

```
rows entityDetails [
  { tableName: "users", columns: [ "id: uuid, PK", "email: string" ] }
]
```

```
// WARNING: 'entityDetails[0].columns[0]' (line 2) expected an object with
// keys { column, type, constraints, description } but found a string —
// setEntityDetails kept the value as written, but it will render as a blank
// or placeholder cell. Check setEntityDetails's `sectionGuide()` example for
// the exact nested shape.
```

The value is still emitted exactly as written — a shape mismatch warns, it
never silently drops content — so the row is there to fix once you've read
the warning. The check only flags a list/object given where a plain value was
expected, or vice versa; it never complains about a number where the example
happened to be a string, since that still renders fine. If you'd rather see
the expected shape before you write the row, `docx-sdk.md`'s setter tables
show it inline (`columns: { column, type, constraints, description }[]`), or
call `SDK.sectionGuide()` directly for the setter's own example payload.

### Authoring the whole master document

In a `type=master` file, wrap each Part's content in a `part <key> { ... }`
block. The body is the same grammar as a section body, so a Part can be lifted
into its own `type=<key>` file — or folded back in — without rewriting it:

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

Parts you leave out still render as their own blank template, so this produces
a complete document rather than one with holes. A `part` block in a
single-Part file — or a `section` outside every `part` in a master file — is
skipped with a warning.

### Legacy `hardware:` flag

`hardware: true` still parses but has no effect; the suite has no hardware
section. The compiler emits a warning telling you to delete the line.

## Xlsx trackers

### Structure

```
doc "Sprint Tracker" type=xlsx
calendar <start-date> .. <end-date>   # optional — computed from tasks if omitted

group "<Section Title>" [color=<name-or-hex>] [rowColor=<name-or-hex>] {
  task "<Task Name>" start=<date> end=<date> [<key>=<value> ...]
}

callout <date> "<text>" [color=<name-or-hex>]
```

- `calendar 2026-01-01 .. 2026-06-30` maps to `setCalendarRange(...)`. Omitted,
  the SDK computes the range from the min/max task dates.
- `group "<title>"` maps to one `addSection({title, bannerColor, rowColor, tasks})`.
  `color=` sets the banner; `rowColor=` the row background. Both accept a named
  preset (`green`, `yellow`, `pink`, `blue`, `purple`, `gray`) or a raw hex code
  (the `#` is optional).
- `task "<name>" start=... end=...` requires at least `start`/`end`. Other
  `key=value` pairs pass through with two automatic renames: `start` →
  `startDate`, `end` → `endDate`. Everything else (`mandays`, `detail`,
  `deliverables`, `notes`, `commitCount`, `lines`, `check`, …) passes as-is —
  see [xlsx-sdk.md](./xlsx-sdk.md#task-fields).
- A task attribute value may be an object, so a clickable note works:
  `notes={ text: "see commit", hyperlink: "https://github.com/you/repo/commit/abc" }`.
- `callout <date> "<text>" [color=...]` maps to `addCallout({dateStr, text, color})`.
  **Note:** callouts are accepted by the SDK but never rendered onto the
  worksheet — a pre-existing gap in `ExcelTrackerSDK`, not something the DSL
  introduces.

### Supplying commit/line-change data

There is no separate grammar for git stats — write them as ordinary task
attributes:

```
task "Auth API" start=2026-01-05 end=2026-01-10 mandays=5 commitCount=7 lines="+120 / -30"
```

- `commitCount=<number>` shows as the badge on the first day of the Gantt bar.
- `lines="+N / -M"` (a plain string, exactly that format) shows in the tooltip
  note and the "Lines (+/-)" column.

To generate these from real history rather than typing them, run
`pspt scan-git --with-lines` first — see [cli.md](./cli.md#pspt-scan-git).

## Grammar summary

```
program       := doc-decl (calendar-decl)? (hardware-decl)? statement*
doc-decl      := 'doc' string 'type' '=' ident
calendar-decl := 'calendar' date '..' date
hardware-decl := 'hardware' ':' bool          (legacy; warns at codegen)

statement     := section | part | group | callout | table-decl | rows-decl | field

part          := 'part' name '{' section-body* '}'      (type=master only)
section       := 'section' name '{' section-body* '}'
section-body  := field | table-decl | rows-decl | list-decl | object-decl | section
field         := name ':' value
list-decl     := 'list' name '{' item* '}'
item          := 'item' string ('{' item* '}')?
object-decl   := 'object' name '{' (name ':' value ','?)* '}'

table-decl    := 'table' name '{' column* '}'
column        := name ':' string weight?
weight        := ident            (e.g. 'w1', 'w1.5')
rows-decl     := 'rows' name '[' row (',' row)* ','? ']'
row           := '{' kv (',' kv)* ','? '}'
kv            := name ':' value

name          := ident | keyword

group         := 'group' string ('color' '=' ident)? '{' task* '}'
task          := 'task' string kvpair*
callout       := 'callout' date string ('color' '=' ident)?
kvpair        := ident '=' value

value         := string | number | date | bool | null | ident
               | '[' (value (',' value)* ','?)? ']'
               | '{' (name ':' value ','?)* '}'
```

Keywords: `doc`, `type`, `calendar`, `group`, `task`, `callout`, `section`,
`table`, `rows`, `hardware`, `list`, `item`, `object`, `part`, `color`, `true`,
`false`, `null` — all of which are still usable as data-key names, as above.

## Worked examples

The upstream repo carries `examples/fixtures/sample-docx.pspt` (a `type=master`
file filling four Parts), `examples/project-tracker.pspt` (a large
real-world master document), `examples/fixtures/sample-xlsx.pspt`, and
`examples/fixtures/edge-cases/` — ~50 fixtures probing exact grammar boundaries,
every warning path, and deliberately invalid syntax showing the error messages.
