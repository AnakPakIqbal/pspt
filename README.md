# pspt

`pspt` unifies two document generators — a Word (.docx) product-spec builder and an
Excel (.xlsx) Gantt tracker builder — behind one shared core, one CLI, and one small
DSL (`.pspt` files) so you can author specs/trackers without writing JS by hand.

## Packages

| Package | Purpose | Docs |
|---|---|---|
| `pspt-core` | Shared design tokens, table/column-schema, placeholder rule, date/border/color helpers, git-scan utilities | (internal — no public docs, see source) |
| `pspt-docx` | `ProductSpecSDK` — generates Word product-spec documents | [packages/pspt-docx/README.md](packages/pspt-docx/README.md) |
| `pspt-xlsx` | `ExcelTrackerSDK` — generates Excel Gantt project trackers | [packages/pspt-xlsx/README.md](packages/pspt-xlsx/README.md) |
| `pspt-lang` | The `.pspt` DSL — lexer/parser/codegen, compiles to plain JS | [packages/pspt-lang/README.md](packages/pspt-lang/README.md) |
| `pspt-cli` | The `pspt` command-line tool: `compile`, `build`, `scan-git` | [packages/pspt-cli/README.md](packages/pspt-cli/README.md) |

## Install

```bash
cd pspt
npm install
```

This is an npm workspaces monorepo — `npm install` at the root links all `packages/*`
together, so `require('pspt-core')` etc. resolve without any manual linking.

## Quick start — three ways to generate a document

### 1. Write a `.pspt` file and use the CLI (recommended for most authors)

```
doc "Acme Widget Spec" type=docx

section overview {
  executiveSummary: "Acme Widget lets teams connect other widgets together."
}
```

```bash
node packages/pspt-cli/bin/pspt.js build my-spec.pspt --out my-spec.docx
```

See [packages/pspt-lang/README.md](packages/pspt-lang/README.md) for the full `.pspt`
language reference and [packages/pspt-cli/README.md](packages/pspt-cli/README.md) for
every CLI command.

### 2. Call the SDKs directly from JS

```js
const ProductSpecSDK = require('pspt-docx');
const doc = new ProductSpecSDK();
doc.setExecutiveSummary('Acme Widget lets teams connect other widgets together.');
await doc.generate('./my-spec.docx');
```

See [packages/pspt-docx/README.md](packages/pspt-docx/README.md) and
[packages/pspt-xlsx/README.md](packages/pspt-xlsx/README.md) for the full setter
reference of each SDK.

### 3. Pull real commit history into a tracker

```bash
node packages/pspt-cli/bin/pspt.js scan-git ../my-repo --since 2026-01-01 \
  --out scan.json --with-lines
```

Then reference the commit counts / line-change totals from `scan.json` in your
`.pspt` tracker file's `task` attributes (`commitCount=`, `lines=`) or in a
direct `ExcelTrackerSDK.addSection()` call. See
[packages/pspt-cli/README.md](packages/pspt-cli/README.md#scan-git) for details.

## Directory layout

```
pspt/
  packages/
    pspt-core/    shared tokens, table schema, placeholder rule, helpers, git-scan
    pspt-docx/    ProductSpecSDK (Word/.docx)
    pspt-xlsx/    ExcelTrackerSDK (Excel/.xlsx Gantt tracker)
    pspt-lang/    .pspt DSL: lexer, parser, codegen
    pspt-cli/     `pspt` bin: compile / build / scan-git
  examples/
    example-usage.js            direct ProductSpecSDK usage (ported from product-spec-sdk)
    example-tracker-usage.js    direct ExcelTrackerSDK usage (ported from product-spec-sdk)
    fixtures/*.pspt              sample .pspt files (docx + xlsx)
    fixtures/edge-cases/*.pspt   DSL edge-case test fixtures (valid + invalid syntax)
```

## Design notes worth knowing before you dig in

- **Every method from the original SDKs is preserved.** `pspt-docx` still has all
  ~50 `set*()` methods across the 6 document sections (cover, business, functional,
  technical, optional hardware, UI/UX, QA), and `pspt-xlsx` still has every method
  of the original `ExcelTrackerSDK` (`addSection`, `addCallout`, calendar/Gantt
  rendering, etc.) — nothing was cut during the refactor into `pspt-core` +
  adapters.
- **`addCallout`/`this.callouts` is pre-existing dead code** in `pspt-xlsx`: it's
  stored but never actually rendered onto the worksheet. This was true in the
  original `excel-tracker-sdk.js` too — it was not introduced by this refactor.
- **The DSL warns instead of failing silently.** An unrecognized field/table name,
  an orphaned `rows` block with no matching `table`, or a scalar value passed to a
  setter that expects an object shape all produce a `// WARNING:` comment in the
  generated JS rather than corrupting the output or failing the whole build.
  Always check generated `.gen.js` files for `WARNING` comments after compiling.
- **`pspt scan-git` diff stats are opt-in** (`--with-lines`) since they cost one
  extra `git show --shortstat` process per commit — fine for a few dozen commits,
  slow for thousands.
