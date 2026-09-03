<!-- GENERATED FILE — do not edit by hand.
     Produced by `npm run llms:build` from each Part SDK's own sectionGuide().
     Edit scripts/generate-llms-docs.js (or the SDK) and re-run instead. -->

# pspt-docx

**This reference describes pspt-docx 2.0.0.** An agent that installed
`pspt-docx@^1` has the removed `ProductSpecSDK` instead — check with
`require('pspt-docx/package.json').version` before relying on anything below.

`pspt-docx` generates the **Product Documentation Master** — a Word (.docx) suite
made of PIC Matrix front matter followed by 14 numbered Parts. Each Part is a
self-contained mini-document with its own title page and running header, and
can be generated on its own or assembled with every other Part into one file.

There are **184 setters across 15 Part SDKs**. You never touch styling — fonts,
colours, table styles, page setup and header/footer chrome are fixed by the
template.

> **Replaces `ProductSpecSDK`.** The earlier single-document class, and its
> `setCoverPage` / `setExecutiveSummary` / `setEnableHardwareSection` API, no
> longer exist. There is no hardware section — this suite targets software
> products. If you were told otherwise by an older copy of these docs, that
> copy is stale.

## The two ways in

```js
const { ProjectBriefSDK, MasterDocument } = require('pspt-docx');

// One Part on its own
const doc = new ProjectBriefSDK();
doc.setHeaderFooterLabels({ productNameLabel: 'Acme Widget' });
doc.setOverview('...');
await doc.generate('./02-project-brief.docx');

// Or every Part assembled into one document
const master = new MasterDocument();
master.setHeaderFooterLabels({ productNameLabel: 'Acme Widget' }); // all 15 at once
master.part('projectBrief').setOverview('...');
master.part('brd').setPurpose('...');
await master.generate('./product-documentation-master.docx');
```

Every setter returns `this`, so calls chain.

## The placeholder rule

**Skipping a setter is fine and is the intended behaviour.** The section still
renders, using the template's own `[bracketed placeholder]` text, so a reviewer
sees exactly what is still missing. `new AnySDK().generate()` with zero calls
reproduces that Part's blank template exactly, and a `MasterDocument` with only
two Parts filled in still produces a complete, clean document.

Corollary for an agent: **never invent content to avoid an empty section.** If
the codebase gives no signal for a field, leave the setter uncalled, or pass an
explicit `[define …]` note. A placeholder is information; a fabricated value is
a defect.

## Payload shapes

Every setter takes one of five shapes. The third column is how it is written in
a `.pspt` file (see [dsl.md](./dsl.md)); calling the SDK directly, pass the
JavaScript value shown in the second column.

| Shape | JavaScript | `.pspt` form |
| --- | --- | --- |
| `scalar` | a string (or number/boolean) | `overview: "..."` |
| `stringList` | `string[]` | `list objectives { item "..." }` |
| `rows` | `Object[]` — one object per table row | `table x { ... }` + `rows x [ ... ]` |
| `nestedRows` | `{ title, children? }[]` | `list x { item "a" { item "b" } }` |
| `object` | a single plain object | `object metadata { writer: "..." }` |

A row cell may itself be a list or an object — that is what makes
`setEndpoints`' request/response bodies and `setEntityDetails`' column lists
reachable. Extra keys in a row are ignored; a missing key renders as a blank
cell rather than throwing.

## Discovering the API at runtime

Every class carries its own machine-readable guide — prefer it over trusting
any documentation, including this file:

```js
ProjectBriefSDK.sectionGuide();
// [{ method: 'setOverview', purpose: '...', example: '...' }, ...]

const { MasterDocument } = require('pspt-docx');
MasterDocument.parts(); // [{ key, part, title, heading }, ...] in document order
```

## The Parts

| Key | | Class | Setters |
| --- | --- | --- | --- |
| `picMatrix` | front matter | `PicMatrixSDK` | 6 |
| `styleGuide` | Part 1 | `StyleGuideSDK` | 9 |
| `projectBrief` | Part 2 | `ProjectBriefSDK` | 9 |
| `brd` | Part 3 | `BrdSDK` | 19 |
| `prd` | Part 4 | `PrdSDK` | 14 |
| `srs` | Part 5 | `SrsSDK` | 21 |
| `techDoc` | Part 6 | `TechnicalDocumentationSDK` | 41 |
| `uiux` | Part 7 | `UiUxSDK` | 16 |
| `uat` | Part 8 | `UatSDK` | 13 |
| `deploymentGuide` | Part 9 | `DeploymentGuideSDK` | 12 |
| `userManual` | Part 10 | `UserManualSDK` | 11 |
| `changelog` | Part 11 | `ChangelogSDK` | 4 |
| `changeRequestLog` | Part 12 | `ChangeRequestLogSDK` | 3 |
| `glossary` | Part 13 | `GlossarySDK` | 3 |
| `appendix` | Part 14 | `AppendixSDK` | 3 |

Order matters: `MasterDocument` assembles them exactly as listed. `key` is what
you pass to `master.part(key)` and what a `.pspt` file's `type=` or `part` block
names.


## Setter reference

Every setter of every Part, generated from `sectionGuide()`. `Data key` is the
name used in a `.pspt` file — the method name minus `set`, lowercased at the
front.

### Front matter — PIC Matrix & Documentation SOP (`picMatrix`)

`PicMatrixSDK` · 6 setters

Front matter for the whole document: the cover identity table, the team and their RACI-style roles, the PIC matrix saying who makes/checks/approves each Part, and the storage guide. Its `picMatrix` and `storageGuide` tables ship with sensible defaults — leave them alone unless your organisation genuinely differs.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setCoverInfo` | `coverInfo` | object | `{ productName, productType, applicationEngineer, productLead, status, writer, checker: { name, email }, approver: { name, email }, lastUpdate, latestHistory }` | Cover-page identity table: application name, product type, application engineer, product lead/manager, status, writer, checker, approver, last-update date, latest history note. |
| `setDocumentTitle` | `documentTitle` | object | `{ title, tagline }` | Overrides the big cover title and italic tagline underneath it. |
| `setTeam` | `team` | rows | `{ role, people: { name, email }[], roleType, responsibilities: { title, highlight, children: […] }[] }[]` | "1. Team" table: who holds each role, their name(s) (optionally linked to email), the PIC role (Maker/Checker/Approver), and a numbered list of what they do in this process. |
| `setPicMatrix` | `picMatrix` | rows | `{ docType, maker, checker, approver }[]` | "2. PIC Matrix" table: for each document type, who is the Maker/Checker/Approver. If not called, the template's default 13-row matrix (Style Guide through Appendices) is used verbatim. |
| `setStorageGuide` | `storageGuide` | rows | `{ category, process }[]` | "3. Where and How to Store the Documentation" table: category + standard/process pairs. If not called, the template's default 6-row guide is used verbatim. |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel, docTypeLabel, companyLabel, teamLabel }` | Overrides the running header ("Product Name" / "Documentation SOP \| Confidential") and footer ("PointStar / Product Team"). |

`.pspt` forms for this Part:

- **object** — `object coverInfo { key: value }` · `coverInfo`, `documentTitle`, `headerFooterLabels`
- **rows** — `table team { ... }` + `rows team [ ... ]` · `team`, `picMatrix`, `storageGuide`

### Part 1 — Documentation Style Guide (`styleGuide`)

`StyleGuideSDK` · 9 setters

Documents the template's own design system. Usually correct as-is; override only if your product deliberately uses a different visual language.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides the running header/footer product name label. |
| `setFontFamilies` | `fontFamilies` | rows | `{ role, font }[]` | Overrides the "1.1 Font Families" table. |
| `setTypeScale` | `typeScale` | rows | `{ element, size, weight, color, wordStyle }[]` | Overrides the "1.2 Type Scale" table. |
| `setFormattingRules` | `formattingRules` | rows | `{ category, rule }[]` | Overrides the "1.3 Text Formatting Rules" table. |
| `setColorPalette` | `colorPalette` | rows | `{ name, hex, rgb, use }[]` | Overrides the "2. Colour Palette" table. |
| `setStatusColors` | `statusColors` | rows | `{ status, hex, use }[]` | Overrides the "2.1 Semantic Status Colours" table. |
| `setHeaderFooterGuidelines` | `headerFooterGuidelines` | rows | `{ element, guideline }[]` | Overrides the "3. Header & Footer Guidelines" table. |
| `setPageSetup` | `pageSetup` | rows | `{ setting, value }[]` | Overrides the "3.1 Page Setup" table. |
| `setDocumentSuiteMap` | `documentSuiteMap` | rows | `{ part, document, audience }[]` | Overrides the "4. Document Suite Map" table — the real list of parts for this product. |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`
- **rows** — `table fontFamilies { ... }` + `rows fontFamilies [ ... ]` · `fontFamilies`, `typeScale`, `formattingRules`, `colorPalette`, `statusColors`, `headerFooterGuidelines`, `pageSetup`, `documentSuiteMap`

### Part 2 — Project Brief (`projectBrief`)

`ProjectBriefSDK` · 9 setters

The elevator pitch: what is being built, the pains that justify it, measurable objectives, the module list that seeds the BRD and PRD, a rough phase timeline, deliverables and early risks.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides the running header product name label. |
| `setMetadata` | `metadata` | object | `{ writer, status, version, lastUpdate }` | Writer/Status/Version/Last Update metadata table. |
| `setOverview` | `overview` | scalar | `string` | 2-4 sentence elevator pitch (Section 1). |
| `setBackgroundPains` | `backgroundPains` | stringList | `string[]` | Bulleted pain points (Section 2). |
| `setObjectives` | `objectives` | stringList | `string[]` | Bulleted measurable objectives (Section 3). |
| `setKeyModules` | `keyModules` | rows | `{ module, features }[]` | Module/Core Features table (Section 4). |
| `setTimeline` | `timeline` | rows | `{ phase, duration }[]` | Phase/duration bullet list (Section 5). |
| `setDeliverables` | `deliverables` | stringList | `string[]` | Bulleted deliverables list (Section 6). |
| `setPreliminaryRisks` | `preliminaryRisks` | rows | `{ risk, mitigation }[]` | Risk/Mitigation table (Section 7). |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`, `metadata`
- **rows** — `table keyModules { ... }` + `rows keyModules [ ... ]` · `keyModules`, `timeline`, `preliminaryRisks`
- **scalar** — `overview: "..."` · `overview`
- **stringList** — `list backgroundPains { item "..." }` · `backgroundPains`, `objectives`, `deliverables`

### Part 3 — Business Requirements Document (`brd`)

`BrdSDK` · 19 setters

Business Requirements. The largest Part. Business-language requirements (BR-xxx) each traced to an objective, the process and its RACI, actors, use cases and their detailed specs, and the traceability matrix the SRS reads back.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides the running header product name label. |
| `setMetadata` | `metadata` | object | `{ writer, status, version, lastUpdate }` | Writer/Status/Version/Last Update metadata table. |
| `setPurpose` | `purpose` | scalar | `string` | Section 1 purpose paragraph. |
| `setBusinessObjectives` | `businessObjectives` | stringList | `string[]` | Section 2 bulleted objectives. |
| `setScope` | `scope` | object | `{ inScope, outOfScope }` | Section 3 in/out of scope statement. |
| `setProjectRoles` | `projectRoles` | rows | `{ name, role, responsibility }[]` | Section 4 Name/Role/Responsibility table. |
| `setBusinessRequirements` | `businessRequirements` | rows | `{ id, requirement, objective, priority }[]` | Section 5 ID/Requirement/Objective/Priority table. |
| `setProcessOverview` | `processOverview` | scalar | `string` | Section 6.1 process description. |
| `setProcessScope` | `processScope` | scalar | `string` | Section 6.2 process boundaries. |
| `setRaci` | `raci` | object | `{ roles: string[], steps: { step, values: string[] }[] }` | Section 6.3 RACI table. |
| `setExceptions` | `exceptions` | stringList | `string[]` | Section 6.5 bulleted exceptions/edge cases. |
| `setActors` | `actors` | rows | `{ actor, description }[]` | Section 7.1 Actor/Description table. |
| `setUseCaseList` | `useCaseList` | rows | `{ id, name, actor, description }[]` | Section 7.3 ID/Name/Actor/Description table. |
| `setUseCaseSpecs` | `useCaseSpecs` | rows | `{ id, name, primaryActor, preconditions, postconditions, mainFlow, alternateFlow, exceptionFlow, businessRules, relatedRequirements }[]` | Section 7.4 detailed use case blocks. |
| `setTraceabilityMatrix` | `traceabilityMatrix` | rows | `{ useCase, requirements }[]` | Section 7.5 Use Case/Requirement(s) table. |
| `setAssumptionsDependencies` | `assumptionsDependencies` | stringList | `string[]` | Section 8 bulleted list. |
| `setRisks` | `risks` | stringList | `string[]` | Section 9 bulleted list. |
| `setApprovals` | `approvals` | rows | `{ name, role, signature, date }[]` | Section 10 Name/Role/Signature/Date table. |
| `setRevisionHistory` | `revisionHistory` | rows | `{ version, date, author, changes }[]` | Section 11 Version/Date/Author/Changes table. |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`, `metadata`, `scope`, `raci`
- **rows** — `table projectRoles { ... }` + `rows projectRoles [ ... ]` · `projectRoles`, `businessRequirements`, `actors`, `useCaseList`, `useCaseSpecs`, `traceabilityMatrix`, `approvals`, `revisionHistory`
- **scalar** — `purpose: "..."` · `purpose`, `processOverview`, `processScope`
- **stringList** — `list businessObjectives { item "..." }` · `businessObjectives`, `exceptions`, `assumptionsDependencies`, `risks`

### Part 4 — Product Requirements Document (`prd`)

`PrdSDK` · 14 setters

Product Requirements. Personas, MoSCoW-prioritised features (which the SRS turns into testable FR-xxx), user stories, the single canonical success-metrics table and the single Risk Register.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides the running header product name label. |
| `setMetadata` | `metadata` | object | `{ writer, status, version, lastUpdate }` | Writer/Status/Version/Last Update metadata table. |
| `setOverview` | `overview` | scalar | `string` | Section 1: 2-3 sentence overview. |
| `setTargetAudience` | `targetAudience` | rows | `{ persona, needs, painPoints }[]` | Section 2 Persona/Needs/Pain Points table. |
| `setRequirements` | `requirements` | rows | `{ id, feature, description, priority, status }[]` | Section 3 ID/Feature/Description/Priority/Status table. |
| `setNfrNotes` | `nfrNotes` | stringList | `string[]` | Section 3.1 bulleted NFR category notes. |
| `setUserStories` | `userStories` | rows | `{ story, bullets: string[] }[]` | Section 4 story blocks. |
| `setDesignUx` | `designUx` | stringList | `string[]` | Section 5 bulleted design links. |
| `setTechnicalConsiderations` | `technicalConsiderations` | stringList | `string[]` | Section 6 bulleted technical notes. |
| `setSuccessMetrics` | `successMetrics` | rows | `{ metric, baseline, target, owner }[]` | Section 7 Metric/Baseline/Target/Owner table. |
| `setRiskRegister` | `riskRegister` | rows | `{ risk, likelihood, impact, mitigation }[]` | Section 8.1 Risk/Likelihood/Impact/Mitigation table. |
| `setOpenQuestions` | `openQuestions` | stringList | `string[]` | Section 8.2 bulleted questions. |
| `setRoadmap` | `roadmap` | rows | `{ horizon, focus }[]` | Section 9 Horizon/Focus table. |
| `setRevisionHistory` | `revisionHistory` | rows | `{ version, date, author, changes }[]` | Section 10 Version/Date/Author/Changes table. |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`, `metadata`
- **rows** — `table targetAudience { ... }` + `rows targetAudience [ ... ]` · `targetAudience`, `requirements`, `userStories`, `successMetrics`, `riskRegister`, `roadmap`, `revisionHistory`
- **scalar** — `overview: "..."` · `overview`
- **stringList** — `list nfrNotes { item "..." }` · `nfrNotes`, `designUx`, `technicalConsiderations`, `openQuestions`

### Part 5 — Software Requirements Specification (`srs`)

`SrsSDK` · 21 setters

Software Requirements. Atomic, testable functional requirements traced back to BRD ids, non-functional targets, interfaces, and the screen inventory.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides the running header product name label. |
| `setMetadata` | `metadata` | object | `{ writer, status, version, lastUpdate }` | Writer/Status/Version/Last Update metadata table. |
| `setPurposeAudience` | `purposeAudience` | scalar | `string` | Section 1.1 purpose & audience. |
| `setScope` | `scope` | scalar | `string` | Section 1.2 scope. |
| `setReferences` | `references` | stringList | `string[]` | Section 1.3 bulleted references. |
| `setProductPerspective` | `productPerspective` | scalar | `string` | Section 2.1 product perspective. |
| `setProductFunctions` | `productFunctions` | stringList | `string[]` | Section 2.2 bulleted functions. |
| `setUserClasses` | `userClasses` | rows | `{ userClass, description, technicalLevel }[]` | Section 2.3 User Class/Description/Technical Level table. |
| `setOperatingEnvironment` | `operatingEnvironment` | stringList | `string[]` | Section 2.4 bulleted environment notes. |
| `setConstraints` | `constraints` | stringList | `string[]` | Section 2.5 bulleted constraints. |
| `setAssumptionsDependencies` | `assumptionsDependencies` | stringList | `string[]` | Section 2.6 bulleted assumptions. |
| `setFunctionalRequirements` | `functionalRequirements` | rows | `{ id, requirement, tracesTo }[]` | Section 3 ID/Requirement/Traces to table. |
| `setNonFunctionalRequirements` | `nonFunctionalRequirements` | rows | `{ category, requirement }[]` | Section 4 Category/Requirement table. |
| `setUserInterfaces` | `userInterfaces` | rows | `{ screenId, screenName, relatedRequirements }[]` | Section 5.1 Screen ID/Name/Related Requirements table. |
| `setHardwareInterfaces` | `hardwareInterfaces` | scalar | `string` | Section 5.2 hardware interfaces note. |
| `setSoftwareInterfaces` | `softwareInterfaces` | scalar | `string` | Section 5.3 software interfaces note. |
| `setCommunicationInterfaces` | `communicationInterfaces` | scalar | `string` | Section 5.4 communication interfaces note. |
| `setSystemFeature` | `systemFeature` | scalar | `string` | Section 6 detailed feature spec paragraph. |
| `setOtherRequirements` | `otherRequirements` | stringList | `string[]` | Section 7 bulleted list. |
| `setAppendix` | `appendix` | scalar | `string` | Section 8 open issues paragraph. |
| `setRevisionHistory` | `revisionHistory` | rows | `{ version, date, author, changes }[]` | Section 9 Version/Date/Author/Changes table. |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`, `metadata`
- **rows** — `table userClasses { ... }` + `rows userClasses [ ... ]` · `userClasses`, `functionalRequirements`, `nonFunctionalRequirements`, `userInterfaces`, `revisionHistory`
- **scalar** — `purposeAudience: "..."` · `purposeAudience`, `scope`, `productPerspective`, `hardwareInterfaces`, `softwareInterfaces`, `communicationInterfaces`, `systemFeature`, `appendix`
- **stringList** — `list references { item "..." }` · `references`, `productFunctions`, `operatingEnvironment`, `constraints`, `assumptionsDependencies`, `otherRequirements`

### Part 6 — Technical Documentation, Data Model & API Spec (`techDoc`)

`TechnicalDocumentationSDK` · 41 setters

Three mini-documents in one Part, generated back to back: **Technical Documentation** (architecture, components, stack, repo tree, data flow, integrations, security), **Data Model** (entity list, per-entity columns, relationships, indexing) and **API Specification** (auth, response envelope, error codes, endpoints, rate limiting, changelog).

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides running header product name label. |
| `setMetadata` | `metadata` | object | `{ writer, status, version, lastUpdate }` | Writer/Status/Version/Last Update metadata table (applies to 6a). |
| `setArchitectureOverview` | `architectureOverview` | scalar | `string` | 6a §1.1 architecture style summary. |
| `setComponents` | `components` | rows | `{ component, responsibility }[]` | 6a §1.3 Component/Responsibility table. |
| `setTechStack` | `techStack` | rows | `{ layer, technology, version, notes }[]` | 6a §1.4 Layer/Technology/Version/Notes table. |
| `setCodebaseTree` | `codebaseTree` | stringList | `string[]` | 6a §1.5 repo tree lines (rendered as a shaded code block). |
| `setNamingConventions` | `namingConventions` | rows | `{ item, convention, example }[]` | 6a §1.5 Item/Convention/Example table. |
| `setDataFlow` | `dataFlow` | scalar | `string` | 6a §1.6 data flow walkthrough. |
| `setIntegrationPoints` | `integrationPoints` | rows | `{ system, purpose, protocol, criticality }[]` | 6a §1.7 External System/Purpose/Protocol/Criticality table. |
| `setScalabilityNotes` | `scalabilityNotes` | stringList | `string[]` | 6a §1.8 bulleted scalability notes. |
| `setDeploymentTopologyNote` | `deploymentTopologyNote` | scalar | `string` | 6a §1.9 deployment topology note. |
| `setSecurityOverview` | `securityOverview` | scalar | `string` | 6a §2.1 security scope/compliance. |
| `setAuthRequirements` | `authRequirements` | rows | `{ id, requirement }[]` | 6a §2.2 ID/Requirement table. |
| `setDataProtection` | `dataProtection` | rows | `{ id, requirement }[]` | 6a §2.3 ID/Requirement table. |
| `setInputValidation` | `inputValidation` | rows | `{ id, requirement }[]` | 6a §2.4 ID/Requirement table. |
| `setInputValidationCode` | `inputValidationCode` | stringList | `string[]` | 6a §2.4 code sample lines. |
| `setSessionManagement` | `sessionManagement` | stringList | `string[]` | 6a §2.5 bulleted notes. |
| `setAuditLogging` | `auditLogging` | rows | `{ event, fields }[]` | 6a §2.6 Event/Logged Fields table. |
| `setComplianceRequirements` | `complianceRequirements` | stringList | `string[]` | 6a §2.7 bulleted notes. |
| `setVulnerabilityManagement` | `vulnerabilityManagement` | rows | `{ activity, frequency }[]` | 6a §2.8 Activity/Frequency table. |
| `setSecurityChecklist` | `securityChecklist` | stringList | `string[]` | 6a §2.9 checklist items. |
| `setRevisionHistory` | `revisionHistory` | rows | `{ version, date, author, changes }[]` | 6a §3 Version/Date/Author/Changes table. |
| `setErdOverview` | `erdOverview` | scalar | `string` | ERD §1 data domain/engine. |
| `setEntityList` | `entityList` | rows | `{ entity, description }[]` | ERD §2 Entity/Description table. |
| `setErdDiagramNote` | `erdDiagramNote` | scalar | `string` | ERD §3 diagram placeholder note. |
| `setEntityDetails` | `entityDetails` | rows | `{ tableName, columns: { column, type, constraints, description }[] }[]` | ERD §4 per-entity column tables. |
| `setRelationships` | `relationships` | rows | `{ from, to, cardinality, rule }[]` | ERD §5 From/To/Cardinality/Rule table. |
| `setNormalisationNotes` | `normalisationNotes` | scalar | `string` | ERD §6 note. |
| `setIndexingStrategy` | `indexingStrategy` | rows | `{ table, index, reason }[]` | ERD §7 Table/Index/Reason table. |
| `setApiOverview` | `apiOverview` | scalar | `string` | API §1 purpose/style/consumers. |
| `setApiBaseUrls` | `apiBaseUrls` | stringList | `string[]` | API §2 base URL lines. |
| `setApiVersioningNote` | `apiVersioningNote` | scalar | `string` | API §2 versioning note. |
| `setApiAuthNote` | `apiAuthNote` | scalar | `string` | API §3 auth scheme note. |
| `setApiAuthHeader` | `apiAuthHeader` | stringList | `string[]` | API §3 auth header code line(s). |
| `setApiRolesNote` | `apiRolesNote` | scalar | `string` | API §3 roles note. |
| `setApiResponseFormat` | `apiResponseFormat` | stringList | `string[]` | API §4 response format JSON lines. |
| `setApiErrorCodes` | `apiErrorCodes` | rows | `{ status, meaning }[]` | API §4 Status/Meaning table. |
| `setEndpoints` | `endpoints` | rows | `{ method, path, description, requestBody: [], responseBody: string[], errors }[]` | API §5 endpoint blocks. |
| `setPaginationNote` | `paginationNote` | scalar | `string` | API §6 pagination note. |
| `setRateLimiting` | `rateLimiting` | rows | `{ limit, window, header }[]` | API §7 Limit/Window/Header table. |
| `setApiChangelog` | `apiChangelog` | rows | `{ version, date, change }[]` | API §8 Version/Date/Change table. |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`, `metadata`
- **rows** — `table components { ... }` + `rows components [ ... ]` · `components`, `techStack`, `namingConventions`, `integrationPoints`, `authRequirements`, `dataProtection`, `inputValidation`, `auditLogging`, `vulnerabilityManagement`, `revisionHistory`, `entityList`, `entityDetails`, `relationships`, `indexingStrategy`, `apiErrorCodes`, `endpoints`, `rateLimiting`, `apiChangelog`
- **scalar** — `architectureOverview: "..."` · `architectureOverview`, `dataFlow`, `deploymentTopologyNote`, `securityOverview`, `erdOverview`, `erdDiagramNote`, `normalisationNotes`, `apiOverview`, `apiVersioningNote`, `apiAuthNote`, `apiRolesNote`, `paginationNote`
- **stringList** — `list codebaseTree { item "..." }` · `codebaseTree`, `scalabilityNotes`, `inputValidationCode`, `sessionManagement`, `complianceRequirements`, `securityChecklist`, `apiBaseUrls`, `apiAuthHeader`, `apiResponseFormat`

### Part 7 — UI/UX Documentation (`uiux`)

`UiUxSDK` · 16 setters

Colour palette, typography, components and states, spacing, interaction, accessibility and responsive behaviour. Reference a component library rather than redefining one.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides the running header product name label. |
| `setMetadata` | `metadata` | object | `{ writer, status, version, lastUpdate }` | Writer/Status/Version/Last Update metadata table. |
| `setOverviewBullets` | `overviewBullets` | stringList | `string[]` | Section 1 bulleted overview (goal/entry points/success state). |
| `setUserFlowNote` | `userFlowNote` | scalar | `string` | Section 2 flow/prototype narration. |
| `setColorPalette` | `colorPalette` | rows | `{ role, hex, usage }[]` | Section 3.1 Role/Hex/Usage table. |
| `setTypography` | `typography` | rows | `{ style, spec, usage }[]` | Section 3.2 Style/Font-Size-Weight/Usage table. |
| `setComponents` | `components` | rows | `{ component, states, notes }[]` | Section 3.3 Component/States/Notes table. |
| `setSpacingGridNote` | `spacingGridNote` | scalar | `string` | Section 3.4 spacing/grid note. |
| `setComponentBehaviour` | `componentBehaviour` | stringList | `string[]` | Section 4.1 bulleted interaction notes. |
| `setEdgeCases` | `edgeCases` | stringList | `string[]` | Section 4.2 bulleted edge cases. |
| `setAnimationNote` | `animationNote` | scalar | `string` | Section 4.3 animation/motion note. |
| `setContentCopy` | `contentCopy` | rows | `{ element, copy, notes }[]` | Section 5 Element/Copy/Notes table. |
| `setAccessibilityHeading` | `accessibilityHeading` | scalar | `string` | Overrides Section 6's heading text (defaults to the source template's exact wording, including its glitch). |
| `setAccessibilityBullets` | `accessibilityBullets` | stringList | `string[]` | Section 6 bulleted accessibility notes. |
| `setResponsiveNotes` | `responsiveNotes` | rows | `{ platform, notes }[]` | Section 7 Platform/Notes table. |
| `setRevisionHistory` | `revisionHistory` | rows | `{ version, date, author, changes }[]` | Section 8 Version/Date/Author/Changes table. |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`, `metadata`
- **rows** — `table colorPalette { ... }` + `rows colorPalette [ ... ]` · `colorPalette`, `typography`, `components`, `contentCopy`, `responsiveNotes`, `revisionHistory`
- **scalar** — `userFlowNote: "..."` · `userFlowNote`, `spacingGridNote`, `animationNote`, `accessibilityHeading`
- **stringList** — `list overviewBullets { item "..." }` · `overviewBullets`, `componentBehaviour`, `edgeCases`, `accessibilityBullets`

### Part 8 — User Acceptance Testing (`uat`)

`UatSDK` · 13 setters

Objectives, features in scope, test levels and owners, environments, the test-case table traced to BRD use cases, non-functional criteria, severity definitions and sign-off.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides the running header product name label. |
| `setMetadata` | `metadata` | object | `{ writer, status, version, lastUpdate }` | Writer/Status/Version/Last Update metadata table. |
| `setObjectivesScope` | `objectivesScope` | scalar | `string` | Section 1 objectives/scope paragraph. |
| `setFeaturesToTest` | `featuresToTest` | stringList | `string[]` | Section 1.1 bulleted features. |
| `setTestStrategy` | `testStrategy` | rows | `{ level, description, owner }[]` | Section 2 Level/Description/Owner table. |
| `setTestEnvironments` | `testEnvironments` | rows | `{ environment, url, database, notes }[]` | Section 3 Environment/URL/Database/Notes table. |
| `setTestAccounts` | `testAccounts` | stringList | `string[]` | Section 3.1 bulleted accounts/devices notes. |
| `setTestSchedule` | `testSchedule` | rows | `{ activity, start, end }[]` | Section 4 Activity/Start/End table. |
| `setRolesResponsibilities` | `rolesResponsibilities` | rows | `{ role, responsibility }[]` | Section 5 Role/Responsibility table. |
| `setTestCases` | `testCases` | rows | `{ id, relatedUseCase, steps, expectedResult, status }[]` | Section 6 ID/Related Use Case/Steps/Expected Result/Status table. |
| `setNonFunctionalTesting` | `nonFunctionalTesting` | rows | `{ type, criteria, required, status }[]` | Section 7 Type/Criteria/Required/Status table. |
| `setBugSeverityDefinitions` | `bugSeverityDefinitions` | rows | `{ severity, definition }[]` | Section 8 Severity/Definition table. |
| `setSignoff` | `signoff` | rows | `{ role, name, approved, date }[]` | Section 9 Role/Name/Approved/Date table. |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`, `metadata`
- **rows** — `table testStrategy { ... }` + `rows testStrategy [ ... ]` · `testStrategy`, `testEnvironments`, `testSchedule`, `rolesResponsibilities`, `testCases`, `nonFunctionalTesting`, `bugSeverityDefinitions`, `signoff`
- **scalar** — `objectivesScope: "..."` · `objectivesScope`
- **stringList** — `list featuresToTest { item "..." }` · `featuresToTest`, `testAccounts`

### Part 9 — Deployment Guide (`deploymentGuide`)

`DeploymentGuideSDK` · 12 setters

What is deployed where, prerequisites, environments, the `.env` example, the deploy and rollback procedures, every configuration variable, the post-deploy checklist and monitoring.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides the running header product name label. |
| `setMetadata` | `metadata` | object | `{ writer, status, version, lastUpdate }` | Writer/Status/Version/Last Update metadata table. |
| `setOverview` | `overview` | scalar | `string` | Section 1 overview paragraph. |
| `setPrerequisites` | `prerequisites` | rows | `{ requirement, details }[]` | Section 2 Requirement/Details table. |
| `setEnvironments` | `environments` | rows | `{ environment, purpose, branch }[]` | Section 3 Environment/Purpose/Branch table. |
| `setEnvExample` | `envExample` | stringList | `string[]` | Section 3 .env.example code block lines. |
| `setDeploymentSteps` | `deploymentSteps` | stringList | `string[]` | Section 4 shell command code block lines (use empty strings for blank separator lines). |
| `setConfiguration` | `configuration` | rows | `{ variable, description, required }[]` | Section 5 Variable/Description/Required table. |
| `setRollbackSteps` | `rollbackSteps` | stringList | `string[]` | Section 6 rollback shell command code block lines. |
| `setRollbackNote` | `rollbackNote` | scalar | `string` | Section 6 note after the code block. |
| `setPostDeploymentChecklist` | `postDeploymentChecklist` | stringList | `string[]` | Section 7 checklist items. |
| `setMonitoring` | `monitoring` | rows | `{ what, tool, alertRecipient }[]` | Section 8 What/Tool/Alert Recipient table. |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`, `metadata`
- **rows** — `table prerequisites { ... }` + `rows prerequisites [ ... ]` · `prerequisites`, `environments`, `configuration`, `monitoring`
- **scalar** — `overview: "..."` · `overview`, `rollbackNote`
- **stringList** — `list envExample { item "..." }` · `envExample`, `deploymentSteps`, `rollbackSteps`, `postDeploymentChecklist`

### Part 10 — User Manual (`userManual`)

`UserManualSDK` · 11 setters

Written for the end user in plain language: getting started, navigation, one walkthrough per feature, FAQ, troubleshooting and support contact.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides the running header product name label. |
| `setMetadata` | `metadata` | object | `{ writer, status, version, lastUpdate }` | Writer/Status/Version/Last Update metadata table. |
| `setIntroduction` | `introduction` | scalar | `string` | Section 1 introduction paragraph. |
| `setGettingStarted` | `gettingStarted` | scalar | `string` | Section 2 access/login instructions. |
| `setGettingStartedScreenshot` | `gettingStartedScreenshot` | scalar | `string` | Section 2 screenshot caption. |
| `setNavigating` | `navigating` | scalar | `string` | Section 3 layout orientation paragraph. |
| `setNavigatingScreenshot` | `navigatingScreenshot` | scalar | `string` | Section 3 screenshot caption. |
| `setFeatureWalkthroughs` | `featureWalkthroughs` | rows | `{ taskTitle, steps: string[], screenshot }[]` | Section 4 repeatable "How to [Task]" blocks with numbered steps and an optional screenshot line. |
| `setFaq` | `faq` | scalar | `string` | Section 5 FAQ paragraph (Q:/A: format). |
| `setTroubleshooting` | `troubleshooting` | rows | `{ problem, solution }[]` | Section 6 Problem/Solution table. |
| `setSupportContact` | `supportContact` | scalar | `string` | Section 7 support contact paragraph. |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`, `metadata`
- **rows** — `table featureWalkthroughs { ... }` + `rows featureWalkthroughs [ ... ]` · `featureWalkthroughs`, `troubleshooting`
- **scalar** — `introduction: "..."` · `introduction`, `gettingStarted`, `gettingStartedScreenshot`, `navigating`, `navigatingScreenshot`, `faq`, `supportContact`

### Part 11 — Changelog (`changelog`)

`ChangelogSDK` · 4 setters

Semantic-versioned release history with tagged entries (ADDED / CHANGED / FIXED / DEPRECATED / REMOVED / SECURITY), newest release first, plus an Unreleased section.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides the running header product name label. |
| `setMetadata` | `metadata` | object | `{ writer, status, version, lastUpdate }` | Writer/Status/Version/Last Update metadata table. |
| `setUnreleased` | `unreleased` | stringList | `string[]` | Section 3 bulleted unreleased changes. |
| `setReleases` | `releases` | rows | `{ heading, entries: string[] }[]` | Repeatable version blocks (newest first). Each has a heading and TAG — description bullets. |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`, `metadata`
- **rows** — `table releases { ... }` + `rows releases [ ... ]` · `releases`
- **stringList** — `list unreleased { item "..." }` · `unreleased`

### Part 12 — Change Request Log (`changeRequestLog`)

`ChangeRequestLogSDK` · 3 setters

One row per change request as they arrive: id, date, requester, description, affected flow and status.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides the running header product name label. |
| `setMetadata` | `metadata` | object | `{ writer, status, version, lastUpdate }` | Writer/Status/Version/Last Update metadata table. |
| `setChangeRequests` | `changeRequests` | rows | `{ crId, dateRequested, requestedBy, description, affectedFlow, status }[]` | CR ID/Date Requested/Requested By/Description of Change/Affected flow-feature/Status table rows. |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`, `metadata`
- **rows** — `table changeRequests { ... }` + `rows changeRequests [ ... ]` · `changeRequests`

### Part 13 — Glossary (`glossary`)

`GlossarySDK` · 3 setters

One consolidated glossary for the whole document — business terms, acronyms and technical jargon.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides the running header product name label. |
| `setMetadata` | `metadata` | object | `{ writer, status, version, lastUpdate }` | Writer/Status/Version/Last Update metadata table. |
| `setTerms` | `terms` | rows | `{ term, definition }[]` | Term/Definition table rows. |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`, `metadata`
- **rows** — `table terms { ... }` + `rows terms [ ... ]` · `terms`

### Part 14 — Appendix (`appendix`)

`AppendixSDK` · 3 setters

Links out to living resources (tracker, design files, OpenAPI, runbooks, dashboards) instead of duplicating them here.

| Setter | Data key | Shape | Payload | Purpose |
| --- | --- | --- | --- | --- |
| `setHeaderFooterLabels` | `headerFooterLabels` | object | `{ productNameLabel }` | Overrides the running header product name label. |
| `setMetadata` | `metadata` | object | `{ writer, status, version, lastUpdate }` | Writer/Status/Version/Last Update metadata table. |
| `setResources` | `resources` | rows | `{ resource, link }[]` | Resource/Link table rows. |

`.pspt` forms for this Part:

- **object** — `object headerFooterLabels { key: value }` · `headerFooterLabels`, `metadata`
- **rows** — `table resources { ... }` + `rows resources [ ... ]` · `resources`

## Output

```js
await doc.generate('/path/to/output.docx'); // writes to disk
const buf = await doc.toBuffer(); // or get a Buffer directly
```

`MasterDocument` exposes the same two methods. Neither creates the output
directory — it must already exist.

## Images and diagrams

This suite embeds no images. Fields such as `erdDiagramNote` and
`gettingStartedScreenshot` are **caption strings**, not file paths — pass the
caption or link you want printed (`"[See erd.png]"`). Diagrams are authored
externally and referenced.
