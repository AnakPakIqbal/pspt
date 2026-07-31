# pspt-docx

`ProductSpecSDK` generates a Word (.docx) "Product Specification" document with a
fixed, consistent visual style (Arial throughout, navy/blue heading colors,
navy-header/white-text tables, US-Letter page size, header/footer, cover page).
It's a lift-and-adapt port of the original `product-spec-sdk.js` — every setter,
section, and rendering behavior is preserved; only the shared tokens/helpers now
come from [`pspt-core`](../pspt-core).

You call whichever `set*()` methods you have real content for; anything you skip
renders as italicized `[...]` placeholder text, so a human reviewer can see
exactly what's still missing. Calling `generate()` with zero setter calls
reproduces the blank template exactly.

## Install

```js
const ProductSpecSDK = require('pspt-docx');
```

## Quick start

```js
const ProductSpecSDK = require('pspt-docx');

async function main() {
  const doc = new ProductSpecSDK();

  doc.setCoverPage({
    productName: 'Acme Widget',
    shortDescription: 'A widget that connects other widgets.',
    lastUpdated: '2026-07-22',
    status: 'Draft',
  });
  doc.setExecutiveSummary('Acme Widget lets teams save time by connecting widgets automatically.');
  doc.setFeatures([
    { name: 'Push notifications', description: 'Real-time alerts', priority: 'Must' },
  ]);

  await doc.generate('./acme-widget-spec.docx');
}

main();
```

Every setter returns `this`, so you can chain them:

```js
doc.setCoverPage({ productName: 'Acme Widget' })
   .setExecutiveSummary('...')
   .setFeatures([...]);
```

## Discovering sections programmatically

```js
const guide = ProductSpecSDK.sectionGuide(); // static, no instance needed
```

Returns a machine-readable array of every section: `{ method, purpose, example }`.
Useful for an AI agent (or a script) that wants to enumerate what content is
expected without reading the source/JSDoc directly.

## Enabling the Hardware Specification section

Pure-software products should leave this off (the default). Hardware/IoT
products should call:

```js
doc.setEnableHardwareSection(true);
```

This turns on the "2.4 Hardware Specification" block and the "Hardware
Validation" table in Quality Control — both are skipped entirely otherwise.

## Setter reference

All setters accept plain JS values/arrays/objects; see the JSDoc directly above
each method in [`src/product-spec-sdk.js`](src/product-spec-sdk.js) for the
exact expected shape of every parameter. Table-shaped rows are always an array
of plain objects keyed exactly as shown.

### Cover

| Setter | Shape |
|---|---|
| `setCoverPage(p)` | `{ productName, shortDescription, lastUpdated, status, logoImagePath? }` |

### 1. Product Overview

| Setter | Shape |
|---|---|
| `setProductInfo(p)` | `{ productName, version, status }` |
| `setExecutiveSummary(text)` | string |
| `setProductRoadmap(rows)` | `{ phase, theme, timeframe, status }[]` |

### 2.1 Business Perspective

| Setter | Shape |
|---|---|
| `setTargetMarket(rows)` | `{ segment, description }[]` |
| `setCustomerPersonas(rows)` | `{ persona, role, goals, painPoints, buyingBehavior }[]` |
| `setUserJourney(p)` | `{ imagePath?, caption?, widthPx?, heightPx? }` |
| `setUseCases(rows)` | `{ useCase, actor, trigger, outcome }[]` |
| `setCompetitiveAnalysis(rows)` | `{ competitor, strengths, weaknesses, pricing, differentiator }[]` |
| `setPricingModel(p)` | `{ modelDescription, tiers: { tier, price, includes, targetSegment }[] }` |
| `setRevenueModel(text)` | string |

### 2.2 Functional Specification

| Setter | Shape |
|---|---|
| `setFeatures(rows)` | `{ name, description, priority }[]` (priority: Must/Should/Could) |

### 2.3 Technical Specification

| Setter | Shape |
|---|---|
| `setSystemArchitecture(p)` | `{ description, diagramImagePath? }` |
| `setTechnologyStack(rows)` | `{ layer, technology, justification }[]` |
| `setApis(p)` | `{ rows: { endpoint, method, description, authRequired, role }[], docsLink? }` |
| `setDatabaseDesign(p)` | `{ rows: { table, purpose, keyFields, url }[], diagramImagePath? }` |
| `setAuthentication(text)` | string |
| `setSecurity(p)` | `{ measures: { title, children? }[], notes? }` — renders as a nested numbered list |
| `setMonitoring(text)` | string |
| `setLogging(text)` | string |
| `setBackupStrategy(text)` | string |

### 2.4 Hardware Specification (only rendered if `setEnableHardwareSection(true)`)

| Setter | Shape |
|---|---|
| `setHardwareOverview(text)` | string |
| `setHardwareComponents(rows)` | `{ component, function, vendor, altVendor }[]` |
| `setBillOfMaterials(rows)` | `{ partNo, description, qty, unitCost, vendor, leadTime }[]` |
| `setMechanicalDesign(p)` | `{ description, diagramImagePath? }` |
| `setElectricalSpecification(p)` | `{ description, diagramImagePath? }` |
| `setSensors(rows)` | `{ sensor, type, range, accuracy, purpose }[]` |
| `setConnectivity(text)` | string |
| `setPowerRequirements(text)` | string |
| `setFirmware(text)` | string |
| `setCertifications(rows)` | `{ certification, region, status, targetDate }[]` |
| `setEnvironmentalRequirements(text)` | string |
| `setManufacturingNotes(text)` | string |
| `setMaintenance(text)` | string |

### 3. UI/UX Specification

| Setter | Shape |
|---|---|
| `setDesignTools(rows)` | `{ category, tool, link }[]` |
| `setDesignPrinciples(text)` | string |
| `setLayoutGrid(rows)` | `{ property, value }[]` |
| `setTypography(rows)` | `{ weight, sizes }[]` (first row typically `{weight:'Font family', sizes:'<name>'}`) |
| `setColorPalette(rows)` | `{ role, color, hex }[]` |
| `setComponentsStates(rows)` | `{ component, state, behavior }[]` |
| `setResponsiveBehavior(rows)` | `{ breakpoint, device, notes }[]` |
| `setInteractionAnimation(rows)` | `{ aspect, notes }[]` |
| `setRevisionHistory(rows)` | `{ version, date, changes }[]` |

### 4. Quality Control

| Setter | Shape |
|---|---|
| `setTestStrategy(text)` | string |
| `setTestPlan(rows)` | `{ phase, scope, entryCriteria, exitCriteria }[]` |
| `setTestCases(rows)` | `{ id, description, steps, expectedResult, status }[]` |
| `setBugTracking(rows)` | `{ severity, definition, sla }[]` |
| `setSecurityTesting(text)` | string |
| `setHardwareValidation(rows)` | `{ type, description, standard, result }[]` — only rendered if hardware section is enabled |

## Images and diagrams

Any field ending in `ImagePath` (`logoImagePath`, `diagramImagePath`, ...)
follows the same rule: pass a real, existing file path and it's embedded;
omit it (or point at a path that doesn't exist) and a bordered placeholder box
is drawn instead, in the same visual style as the rest of the template.

Architecture/ER/CAD/circuit diagrams are expected to be authored externally
(e.g. as Mermaid `.mmd` source, rendered to PNG via `mermaid-cli`) and passed
in as a `diagramImagePath` — this SDK has no Mermaid-rendering capability of
its own.

## Output

```js
await doc.generate('/path/to/output.docx');       // writes to disk
const buf = await doc.toBuffer();                  // or get a Buffer directly
```

## Using it via the `.pspt` DSL instead

If you'd rather author specs in a terser, non-JS format, see
[`pspt-lang`](../pspt-lang/README.md) — it compiles `.pspt` files into scripts
that call this same SDK. Note the DSL currently only supports scalar/table
fields; a handful of setters that expect a nested *object* payload (`setCoverPage`,
`setApis`, `setPricingModel`, `setSystemArchitecture`, `setDatabaseDesign`,
`setUserJourney`, `setMechanicalDesign`, `setElectricalSpecification`) aren't yet
expressible in the DSL grammar and must be set via direct SDK calls for now.
