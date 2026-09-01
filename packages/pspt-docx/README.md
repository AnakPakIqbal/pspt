# pspt-docx

Generates the **Product Documentation Master** — a Word (.docx) document suite
made of a PIC Matrix front-matter section followed by 14 numbered Parts (Style
Guide, Project Brief, BRD, PRD, SRS, Technical Documentation, UI/UX, UAT,
Deployment Guide, User Manual, Changelog, Change Request Log, Glossary,
Appendix).

Each Part is a self-contained mini-document with its own title page and running
header, driven by `set...()` methods. You never touch styling — fonts, colors,
table styles, page setup and header/footer chrome all come from the shared
tokens in [`src/sdk/pspt-core.js`](src/sdk/pspt-core.js).

> Replaces the earlier single-document `ProductSpecSDK`. That class and its
> `setCoverPage`/`setExecutiveSummary`/`setEnableHardwareSection` API are gone —
> see [Migrating from ProductSpecSDK](#migrating-from-productspecsdk) below.

## Install

```js
const { ProjectBriefSDK, MasterDocument } = require('pspt-docx');
```

## Generate one Part

```js
const { ProjectBriefSDK } = require('pspt-docx');

const doc = new ProjectBriefSDK();
doc.setHeaderFooterLabels({ productNameLabel: 'Acme Widget' });
doc.setOverview('Acme Widget lets small teams submit and approve expenses from their phones.');
doc.setObjectives(['Cut expense processing time by 70%']);
doc.setKeyModules([{ module: 'Receipts', features: 'Photo capture + OCR amount extraction' }]);

await doc.generate('./02-project-brief.docx');
```

Every setter returns `this`, so calls chain. **Skipping a setter is fine** — the
section still renders, using the template's own `[bracketed placeholder]` text,
so a reviewer can see exactly what is still missing. `new AnySDK().generate()`
with zero calls reproduces that Part's blank template exactly.

## Generate the whole master document

`MasterDocument` assembles every Part into one `.docx` in a single run. It does
not merge packed files — because every Part is built by the same library with
the same style ids, it concatenates their _sections_ into one `Document`, so
there is no OOXML surgery and Word's "Page X of Y" fields resolve against the
real total.

```js
const { MasterDocument } = require('pspt-docx');

const master = new MasterDocument();
master.setHeaderFooterLabels({ productNameLabel: 'Acme Widget' }); // applies to all 15

master.part('projectBrief').setOverview('...').setObjectives(['...']);
master.part('brd').setPurpose('...');

await master.generate('./product-documentation-master.docx');
```

Parts you never touch still render as their blank template, so an unfinished
suite still produces a clean, complete document.

| Option                                                  | Effect                                                                                       |
| ------------------------------------------------------- | -------------------------------------------------------------------------------------------- |
| `new MasterDocument({ include: ['picMatrix', 'brd'] })` | Assemble only those Parts, still in canonical order                                          |
| `MasterDocument.parts()`                                | The ordered Part list — `{ key, part, title, heading }`, enough to build a table of contents |
| `master.part(key)`                                      | The SDK instance for one Part, created on first access                                       |

## The Parts

| Key                |              | Class                       | Module                                                           |
| ------------------ | ------------ | --------------------------- | ---------------------------------------------------------------- |
| `picMatrix`        | front matter | `PicMatrixSDK`              | [pic-matrix-docx.js](src/sdk/pic-matrix-docx.js)                 |
| `styleGuide`       | Part 1       | `StyleGuideSDK`             | [style-guide-docx.js](src/sdk/style-guide-docx.js)               |
| `projectBrief`     | Part 2       | `ProjectBriefSDK`           | [project-brief-docx.js](src/sdk/project-brief-docx.js)           |
| `brd`              | Part 3       | `BrdSDK`                    | [brd-docx.js](src/sdk/brd-docx.js)                               |
| `prd`              | Part 4       | `PrdSDK`                    | [prd-docx.js](src/sdk/prd-docx.js)                               |
| `srs`              | Part 5       | `SrsSDK`                    | [srs-docx.js](src/sdk/srs-docx.js)                               |
| `techDoc`          | Part 6       | `TechnicalDocumentationSDK` | [tech-doc-docx.js](src/sdk/tech-doc-docx.js)                     |
| `uiux`             | Part 7       | `UiUxSDK`                   | [uiux-docx.js](src/sdk/uiux-docx.js)                             |
| `uat`              | Part 8       | `UatSDK`                    | [uat-docx.js](src/sdk/uat-docx.js)                               |
| `deploymentGuide`  | Part 9       | `DeploymentGuideSDK`        | [deployment-guide-docx.js](src/sdk/deployment-guide-docx.js)     |
| `userManual`       | Part 10      | `UserManualSDK`             | [user-manual-docx.js](src/sdk/user-manual-docx.js)               |
| `changelog`        | Part 11      | `ChangelogSDK`              | [changelog-docx.js](src/sdk/changelog-docx.js)                   |
| `changeRequestLog` | Part 12      | `ChangeRequestLogSDK`       | [change-request-log-docx.js](src/sdk/change-request-log-docx.js) |
| `glossary`         | Part 13      | `GlossarySDK`               | [glossary-docx.js](src/sdk/glossary-docx.js)                     |
| `appendix`         | Part 14      | `AppendixSDK`               | [appendix-docx.js](src/sdk/appendix-docx.js)                     |

Part 6 is three mini-documents in one module: Technical Documentation, then Data
Model (ERD), then API Specification.

## Discovering an SDK's API

Every class exposes a machine-readable section guide — the fastest way for a
script or an agent to learn what a Part expects without reading the source:

```js
console.log(ProjectBriefSDK.sectionGuide());
// [{ method: 'setOverview', purpose: '...', example: '...' }, ...]
```

**[`src/README.md`](src/README.md) is the full authoring reference** — every
setter of every Part, the exact shape of its payload, and what to prepare before
you start writing.

## Layout

```
src/
├── index.js       package entry point (this is what `require('pspt-docx')` gives you)
├── README.md      full authoring reference for all 15 Parts
├── sdk/           the SDK modules — keep them together, they require each other by relative path
└── sample-docs/   one example .docx per Part, generated with placeholder content
```

`src/sdk/` is deliberately self-contained: it depends only on `docx` and its own
`pspt-core.js`, so the folder can be copied out and used standalone.

## Shared tokens and primitives

[`src/sdk/pspt-core.js`](src/sdk/pspt-core.js) is the single source of truth for
the suite's look — font, sizes, colors, page geometry — plus the paragraph and
table primitives every Part shares (`run`, `h1`/`h2`/`h3`, `guidanceNote`,
`bullet`, `codeBlock`, `checklistItem`, `buildZebraTable`, …). Change a value
there and it ripples through all 15 Parts, which is the point; hand-tuning a
value inside a single Part module is how the suite drifts out of sync.

Note this is a _different_ file from the `pspt-core` workspace package, which
carries the Arial-based token set used by `pspt-xlsx`. This suite's tokens are
Calibri-based and were extracted from the master template's own XML.

## Migrating from ProductSpecSDK

The old single-document `ProductSpecSDK` is gone. Its content maps onto the new
Parts roughly as follows:

| Old setter                                                                                                                                       | Now lives in                                                            |
| ------------------------------------------------------------------------------------------------------------------------------------------------ | ----------------------------------------------------------------------- |
| `setCoverPage`, `setProductInfo`                                                                                                                 | `PicMatrixSDK.setCoverInfo`                                             |
| `setExecutiveSummary`, `setProductRoadmap`                                                                                                       | `ProjectBriefSDK` / `PrdSDK.setRoadmap`                                 |
| `setTargetMarket`, `setCustomerPersonas`, `setUseCases`, `setCompetitiveAnalysis`, `setPricingModel`                                             | `BrdSDK` (use cases, actors, RACI) and `PrdSDK.setTargetAudience`       |
| `setFeatures`                                                                                                                                    | `PrdSDK.setRequirements` → `SrsSDK.setFunctionalRequirements`           |
| `setSystemArchitecture`, `setTechnologyStack`, `setApis`, `setDatabaseDesign`, `setSecurity`, `setMonitoring`, `setLogging`, `setBackupStrategy` | `TechnicalDocumentationSDK` (all three sub-documents)                   |
| `setDesignTools`, `setTypography`, `setColorPalette`, `setComponentsStates`, `setResponsiveBehavior`                                             | `UiUxSDK` and `StyleGuideSDK`                                           |
| `setTestStrategy`, `setTestPlan`, `setTestCases`, `setBugTracking`                                                                               | `UatSDK`                                                                |
| `setEnableHardwareSection` and the hardware setters                                                                                              | **no equivalent** — the new suite is written for pure-software products |

## Authoring via the `.pspt` DSL

[`pspt-lang`](../pspt-lang/README.md) compiles `.pspt` files into scripts that
call these SDKs — see its README for the per-Part `type=` declarations.
