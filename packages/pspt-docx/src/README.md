# Product Documentation Master — SDK Suite

This is a set of 16 Node.js modules that regenerate the "Product Documentation
Master" Word template **exactly** — same fonts, colors, table styles, page
setup, and header/footer behavior as the original — driven entirely by data
you provide. You never touch styling; you just call `set...()` methods with
your product's real content.

The master document is one 87-page Word file made of a PIC Matrix
(who-does-what for each doc type) followed by 14 numbered "Parts," each a
self-contained mini-document (its own title page, its own running header).
Each Part has its own SDK file here so you can generate just the ones you
need, in any order.

---

## 0. Folder structure

```
product-documentation-sdk/
├── README.md          ← this file
├── sdk/                ← all 16 .js modules — keep them together in one folder
└── sample-docs/        ← one example .docx per part, generated with placeholder content
                            (for reference — see what each part looks like before you fill it in)
```

Everything below refers to files inside `sdk/`.

## 1. What's in this folder

| File                         | What it generates                                                                                      |
| ---------------------------- | ------------------------------------------------------------------------------------------------------ |
| `pspt-core.js`               | Shared constants (fonts, colors, page size). Required by every other file — don't delete or rename it. |
| `pic-matrix-docx.js`         | Front-matter: cover info table, "1. Team" (RACI-style roles), "2. PIC Matrix," "3. Storage Guide."     |
| `style-guide-docx.js`        | Part 1 — Documentation Style Guide                                                                     |
| `project-brief-docx.js`      | Part 2 — Project Brief                                                                                 |
| `brd-docx.js`                | Part 3 — Business Requirements Document (BRD)                                                          |
| `prd-docx.js`                | Part 4 — Product Requirements Document (PRD)                                                           |
| `srs-docx.js`                | Part 5 — Software Requirements Specification (SRS)                                                     |
| `tech-doc-docx.js`           | Part 6 — Technical Documentation **+ Data Model (ERD) + API Specification** (3 mini-docs in one file)  |
| `uiux-docx.js`               | Part 7 — UI/UX Documentation                                                                           |
| `uat-docx.js`                | Part 8 — User Acceptance Testing (UAT)                                                                 |
| `deployment-guide-docx.js`   | Part 9 — Deployment Guide                                                                              |
| `user-manual-docx.js`        | Part 10 — User Manual                                                                                  |
| `changelog-docx.js`          | Part 11 — Changelog                                                                                    |
| `change-request-log-docx.js` | Part 12 — Change Request Log                                                                           |
| `glossary-docx.js`           | Part 13 — Glossary                                                                                     |
| `appendix-docx.js`           | Part 14 — Appendix                                                                                     |

Each file (except `pspt-core.js`) exports one SDK **class**. Each class has:

- A `static sectionGuide()` — call it to get a JSON list of every available
  section, what it's for, and an example payload. Handy if you (or another
  AI) want to inspect the API instead of reading source.
- `set...()` methods, one per section — call only the ones you have content
  for. Everything you don't call falls back to the template's own default
  placeholder text, so a part generated with zero calls looks exactly like
  the blank template.
- `.generate(outputPath)` — renders and writes the `.docx` file.

---

## 2. Setup (one-time)

You need Node.js and the `docx` npm package.

```bash
npm install docx
```

Keep all 16 `.js` files inside `sdk/` together — they `require('./pspt-core')`
and similar relative paths, so they must sit side by side. Run your own
scripts from inside that folder (or adjust the require paths if you move
them).

---

## 3. How to generate one document

Every module follows the same shape. Example — Project Brief:

```javascript
const ProjectBriefSDK = require('./project-brief-docx');

const doc = new ProjectBriefSDK();

doc.setHeaderFooterLabels({ productNameLabel: 'Acme Widget' });
doc.setMetadata({ writer: 'Jane Doe', status: 'Draft', version: 'V1', lastUpdate: 'Aug 31, 2026' });
doc.setOverview('Acme Widget lets small teams submit and approve expenses from their phones.');
doc.setBackgroundPains([
  'Manual spreadsheet tracking takes 3 hrs/week',
  'No audit trail for approvals',
]);
doc.setObjectives(['Cut expense processing time by 70%', 'Achieve 90% team adoption within Q4']);
doc.setKeyModules([
  { module: 'Receipts', features: 'Photo capture + OCR amount extraction' },
  { module: 'Approvals', features: 'One-tap approve/reject with comments' },
]);
doc.setTimeline([
  { phase: 'Phase 1 — Discovery & Design', duration: '2 weeks' },
  { phase: 'Phase 2 — MVP', duration: '4 weeks' },
]);
doc.setDeliverables(['Mobile app (iOS/Android)', 'Admin web dashboard']);
doc.setPreliminaryRisks([{ risk: 'Bank API rate limits', mitigation: 'Add caching + backoff' }]);

doc.generate('./02-project-brief.docx').then(() => console.log('Done'));
```

Run it with `node your-script.js`.

**Skipping sections is fine.** If you only have some of the content ready,
just don't call those methods — the document will render the template's
original `[bracketed placeholder]` text in their place, so nothing looks
broken or empty.

**To see exactly what a module accepts**, run this once per file:

```javascript
const ProjectBriefSDK = require('./project-brief-docx');
console.log(JSON.stringify(ProjectBriefSDK.sectionGuide(), null, 2));
```

This prints every `set...()` method, what section it fills, and an example
payload shape — the fastest way to know what to prepare (see §5).

---

## 4. Generating the whole 87-page master document

Right now each Part is a **separate `.docx` file** — there are 15 generator
scripts in total (front-matter + 14 Parts). To get one combined 87-page
document, you generate every part you want, then merge the resulting
`.docx` files, in this exact order:

1. **PIC Matrix front-matter** — `pic-matrix-docx.js`
2. **Part 1 — Documentation Style Guide** — `style-guide-docx.js`
3. **Part 2 — Project Brief** — `project-brief-docx.js`
4. **Part 3 — Business Requirements Document (BRD)** — `brd-docx.js`
5. **Part 4 — Product Requirements Document (PRD)** — `prd-docx.js`
6. **Part 5 — Software Requirements Specification (SRS)** — `srs-docx.js`
7. **Part 6 — Technical Documentation** — `tech-doc-docx.js`
   — this single file actually contains **three** mini-documents back to
   back: Technical Documentation (architecture + security), then Data
   Model (ERD), then API Specification. They come out of `.generate()`
   as one continuous `.docx`, so there's nothing extra to merge here —
   just treat this file's output as one block that slots into position 7.
8. **Part 7 — UI/UX Documentation** — `uiux-docx.js`
9. **Part 8 — User Acceptance Testing (UAT)** — `uat-docx.js`
10. **Part 9 — Deployment Guide** — `deployment-guide-docx.js`
11. **Part 10 — User Manual** — `user-manual-docx.js`
12. **Part 11 — Changelog** — `changelog-docx.js`
13. **Part 12 — Change Request Log** — `change-request-log-docx.js`
14. **Part 13 — Glossary** — `glossary-docx.js`
15. **Part 14 — Appendix** — `appendix-docx.js`

**How the merge itself works:** each generated `.docx` already contains the
correct title-only divider page plus its own content section, so merging is
purely concatenation — you are not supposed to re-style, re-number, or
re-order anything inside a part once it's generated. Any standard
DOCX-merging approach works (the `docx` skill's merge helper, `python-docx`
composing library, or opening each file in Word/LibreOffice and using
Insert → Text from File). Page numbers ("Page X of Y") are Word field codes
that recompute automatically once merged, so you don't need to fix those by
hand.

**If a part's content isn't ready yet**, you can still include it in the
merge — every module falls back to the original template's own placeholder
text for any section you didn't call, so an unfinished part still looks
like a clean, intentional blank template section rather than an error or a
missing page.

**Want this automated?** If you'd rather have one script that calls all 15
generators and stitches the results into a single `.docx` in one run
instead of merging by hand, just ask — it's a straightforward addition on
top of what's already built, and would take content for all 15 parts (or
defaults) as a single combined input.

---

## 5. What to prepare before you start

Nothing below is mandatory — call whichever `set...()` methods you have
data for and skip the rest; unset sections just render the template's own
placeholder text. But if your goal is a **complete master document with no
placeholder text left anywhere**, here is every section, part by part, with
the exact method name and the shape of data it expects.

### Front matter — PIC Matrix (`pic-matrix-docx.js`)

- **`setCoverInfo(...)`** — the cover identity table:
  - Product/application name
  - Product type (e.g. "SaaS", "Mobile App", "Internal Tool")
  - Current status (e.g. "Live", "In Development")
  - Application engineer name
  - Product lead/manager name
  - Writer name
  - Checker — name, and optionally an email (renders as a clickable mailto
    link if given)
  - Approver — name, and optionally an email (same mailto behavior)
  - Last-updated date
  - One-line "latest history" note (what changed most recently)
- **`setDocumentTitle(...)`** — only needed if you want to override the big
  cover title and the italic tagline underneath it; otherwise the
  template's own title/tagline is used.
- **`setTeam(...)`** — one entry per team role, each with:
  - Role name (e.g. "PM & Dev Manager")
  - One or more people filling that role (name + optional email each)
  - Their PIC function: Maker, Checker, or Approver
  - A list of responsibility bullets — each can have a `title`, an
    optional `highlight` (a substring of the title to bold+color), and a
    list of `children` sub-bullets
- **`setPicMatrix(...)`** — override which role is Maker/Checker/Approver
  for each of the 14 Parts' document types. If you skip this, the
  template's own default 13-row matrix (every Part from Style Guide through
  Appendices) is used exactly as-is — usually fine to leave alone.
- **`setStorageGuide(...)`** — override the "where files live" table
  (category + process pairs: shared drive location, folder path, naming
  convention, example filenames, etc.). Skipping this keeps the template's
  own default 6-row guide.
- **`setHeaderFooterLabels(...)`** — product name label for the running
  header, plus optional overrides for the doc-type label, company label,
  and team label in the footer.

### Part 1 — Documentation Style Guide (`style-guide-docx.js`)

This Part documents the template's own design system, so it's usually
fine to leave completely untouched. Only override it if your product
deliberately uses a different visual language than this template:

- **`setFontFamilies(...)`** — role/font pairs (Title, Headings, Body text,
  Code/technical values)
- **`setTypeScale(...)`** — element/size/weight/color/Word-style rows for
  every heading level and body text
- **`setFormattingRules(...)`** — category/rule pairs (when to use bold,
  italics, underlining, alignment)
- **`setColorPalette(...)`** — name/hex/RGB/use rows for your brand colors
- **`setStatusColors(...)`** — status/hex/example-use rows (Success,
  Warning, Error, etc.)
- **`setHeaderFooterGuidelines(...)`** — element/guideline rows describing
  header and footer layout rules
- **`setPageSetup(...)`** — setting/value rows (page size, margins, line
  spacing)
- **`setDocumentSuiteMap(...)`** — override the part/document/audience
  table listing all 13 other Parts and who reads each one
- **`setHeaderFooterLabels(...)`** — product name label for the header

### Part 2 — Project Brief (`project-brief-docx.js`)

- **`setOverview(...)`** — a 2–4 sentence elevator pitch: what's being
  built, who it's for, what problem it solves
- **`setBackgroundPains(...)`** — a bulleted list of current-state pain
  points that justify the project
- **`setObjectives(...)`** — a bulleted list of measurable objectives, each
  mapping back to a pain point above
- **`setKeyModules(...)`** — rows of `{ module, features }` — module name
  plus a one-line summary of its core features; keep names consistent with
  the BRD and PRD, since this table seeds both
- **`setTimeline(...)`** — rows of `{ phase, duration }` — a rough
  phase-by-phase estimate (Discovery, MVP, Testing, Deployment, etc.), not
  a detailed schedule
- **`setDeliverables(...)`** — a bulleted list of what ships (software,
  documentation, training, etc.)
- **`setPreliminaryRisks(...)`** — rows of `{ risk, mitigation }` for any
  risks visible before formal requirements exist
- **`setMetadata(...)`** — writer, status, version, last-update date
- **`setHeaderFooterLabels(...)`** — product name label

### Part 3 — Business Requirements Document / BRD (`brd-docx.js`)

The largest and most detail-hungry Part. Gather:

- **`setPurpose(...)`** — one paragraph on why this BRD exists
- **`setBusinessObjectives(...)`** — bulleted objectives tied to a business
  metric where possible
- **`setScope(...)`** — `{ inScope, outOfScope }` — which business
  processes are affected vs explicitly excluded
- **`setProjectRoles(...)`** — rows of `{ name, role, responsibility }` —
  one row per person named in the project
- **`setBusinessRequirements(...)`** — rows of `{ id, requirement,
objective, priority }` — numbered from the business's point of view
  ("the system shall allow..."), each traceable to an objective above;
  these IDs get referenced later by the SRS and Test Plan
- **`setProcessOverview(...)`** — a short paragraph naming the business
  process, its goal, trigger, and outcome
- **`setProcessScope(...)`** — exactly where the process starts and stops
- **`setRaci(...)`** — `{ roles: [...], steps: [{ step, values: [...] }] }`
  — a RACI matrix using functional titles (Customer, Sales Rep, System
  Administrator, etc.), not individual names; `values` are R/A/C/I markers
  aligned to the `roles` array, one array per step
- **`setExceptions(...)`** — bulleted edge cases (rejected steps, timeouts,
  missing information)
- **`setActors(...)`** — rows of `{ actor, description }` — every human
  role or external system that interacts with it
- **`setUseCaseList(...)`** — rows of `{ id, name, actor, description }` —
  a summary table of every use case (e.g. UC-01)
- **`setUseCaseSpecs(...)`** — the detailed spec **per use case** — this is
  the most time-consuming section to fill in properly. For each use case,
  provide: `id`, `name`, `primaryActor`, `preconditions`,
  `postconditions`, `mainFlow` (numbered steps), `alternateFlow`,
  `exceptionFlow`, `businessRules`, `relatedRequirements`
- **`setTraceabilityMatrix(...)`** — rows of `{ useCase, requirements }` —
  mapping each use case ID to the requirement ID(s) it satisfies
- **`setAssumptionsDependencies(...)`** — bulleted assumptions and external
  dependencies
- **`setRisks(...)`** — bulleted business-level risks to flag before
  sign-off (the full Risk Register with likelihood/impact lives in the PRD,
  not repeated here)
- **`setApprovals(...)`** — rows of `{ name, role, signature, date }` for
  the sign-off table
- **`setRevisionHistory(...)`** — rows of `{ version, date, author,
changes }`
- **`setMetadata(...)`** / **`setHeaderFooterLabels(...)`** — as above

### Part 4 — Product Requirements Document / PRD (`prd-docx.js`)

- **`setOverview(...)`** — 2–3 sentences a busy exec could read and
  understand the "what" and "why"
- **`setTargetAudience(...)`** — rows of `{ persona, needs, painPoints }`
  — primary and secondary personas
- **`setRequirements(...)`** — rows of `{ id, feature, description,
priority, status }` — priority uses MoSCoW (Must/Should/Could/Won't);
  the testable, detailed versions of these live in the SRS as FR-xxx,
  traced back to this ID
- **`setNfrNotes(...)`** — bulleted list of which non-functional categories
  matter for this feature (the full testable NFR table with targets lives
  once in the SRS, not duplicated here)
- **`setUserStories(...)`** — array of `{ story, bullets }` — `story` in
  "As a [persona], I want to [action], so that [benefit]" format;
  `bullets` are Given/When/Then flow lines under each story
- **`setDesignUx(...)`** — bulleted links to Figma/mockups rather than
  pasted images (full UX spec lives in Part 7)
- **`setTechnicalConsiderations(...)`** — bulleted high-level technical
  notes (full detail lives in Part 6)
- **`setSuccessMetrics(...)`** — rows of `{ metric, baseline, target,
owner }` — the project's single canonical measurable-success table, each
  metric tied back to a goal in the overview
- **`setRiskRegister(...)`** — rows of `{ risk, likelihood, impact,
mitigation }` — the project's single Risk Register, carrying forward
  risks flagged in the Project Brief and BRD plus any new ones
- **`setOpenQuestions(...)`** — bulleted open questions, ideally each
  noting an owner and due date
- **`setRoadmap(...)`** — rows of `{ horizon, focus }` — high-level themes
  by time horizon (Now/Next/Later), not detailed sprint planning
- **`setRevisionHistory(...)`** — rows of `{ version, date, author,
changes }`
- **`setMetadata(...)`** / **`setHeaderFooterLabels(...)`** — as above

### Part 5 — Software Requirements Specification / SRS (`srs-docx.js`)

- **`setPurposeAudience(...)`** — purpose of this SRS and its intended
  readers (typically engineering + QA)
- **`setScope(...)`** — brief restatement of scope, referencing the BRD's
  scope section rather than repeating it in full
- **`setReferences(...)`** — bulleted references to the Project Brief and
  BRD versions this SRS is built on
- **`setProductPerspective(...)`** — how the system fits into its
  environment (standalone vs integrated)
- **`setProductFunctions(...)`** — bulleted one-line functions
- **`setUserClasses(...)`** — rows of `{ userClass, description,
technicalLevel }` — reduced to the technical-access dimension; link back
  to the PRD's personas rather than re-describing needs/pain points
- **`setOperatingEnvironment(...)`** — bulleted notes: supported browsers,
  server/runtime, target devices
- **`setConstraints(...)`** — bulleted hard constraints (existing
  infrastructure, compliance requirements, etc.)
- **`setAssumptionsDependencies(...)`** — bulleted assumptions
- **`setFunctionalRequirements(...)`** — rows of `{ id, requirement,
tracesTo }` — atomic, one behaviour per line, testable, traced to a BRD
  requirement ID (e.g. FR-MOD-001 → BR-001)
- **`setNonFunctionalRequirements(...)`** — rows of `{ category,
requirement }` — Performance, Security, Usability, Reliability,
  Scalability, etc., each with a concrete testable target
- **`setUserInterfaces(...)`** — rows of `{ screenId, screenName,
relatedRequirements }` — a screen inventory (full visual specs live in
  Part 7)
- **`setHardwareInterfaces(...)`** — note on dedicated hardware interfaces,
  or state none exist
- **`setSoftwareInterfaces(...)`** — note on other systems this software
  talks to (full contract lives in Part 6's API spec)
- **`setCommunicationInterfaces(...)`** — network protocols, data formats,
  transport security
- **`setSystemFeature(...)`** — for complex features needing more than one
  requirement row: description, precondition, main flow, result
- **`setOtherRequirements(...)`** — bulleted legal/compliance/localisation
  requirements not covered elsewhere
- **`setAppendix(...)`** — open issues, each ideally tracked with an issue
  ID
- **`setRevisionHistory(...)`** — rows of `{ version, date, author,
changes }`
- **`setMetadata(...)`** / **`setHeaderFooterLabels(...)`** — as above

### Part 6 — Technical Documentation + ERD + API Spec (`tech-doc-docx.js`)

This one file generates three mini-documents in sequence. Gather content
for all three:

**6a. Technical Documentation (architecture + security)**

- **`setArchitectureOverview(...)`** — overall architecture style
  (monolith/microservices/client-server) and the reasoning behind it
- **`setComponents(...)`** — rows of `{ component, responsibility }`
- **`setTechStack(...)`** — rows of `{ layer, technology, version, notes }`
  — one row per layer: Frontend, Backend, Database, Infrastructure/Hosting,
  CI/CD, Monitoring & Logging
- **`setCodebaseTree(...)`** — an array of lines forming your repo's
  folder tree (rendered as a shaded monospace code block); use empty
  strings for blank separator lines if needed
- **`setNamingConventions(...)`** — rows of `{ item, convention, example }`
- **`setDataFlow(...)`** — a walkthrough of one or two representative
  request flows end-to-end
- **`setIntegrationPoints(...)`** — rows of `{ system, purpose, protocol,
criticality }` — every external system integration and its criticality
- **`setScalabilityNotes(...)`** — bulleted notes: expected load,
  caching/connection-pooling notes, known bottlenecks
- **`setDeploymentTopologyNote(...)`** — where components physically/
  logically run (complements the Deployment Guide's step-by-step
  instructions)
- **`setSecurityOverview(...)`** — scope, purpose, and any targeted
  compliance standards (SOC2, ISO 27001) or "None targeted in this phase"
- **`setAuthRequirements(...)`** — rows of `{ id, requirement }` — e.g.
  SEC-AUTH-01, testable authentication/authorization requirements
- **`setDataProtection(...)`** — rows of `{ id, requirement }` — e.g.
  SEC-DATA-01, encryption/credential-handling requirements
- **`setInputValidation(...)`** — rows of `{ id, requirement }` — e.g.
  SEC-INPUT-01, schema validation and injection-prevention requirements
- **`setInputValidationCode(...)`** — array of code lines illustrating safe
  vs unsafe query patterns (shaded monospace block)
- **`setSessionManagement(...)`** — bulleted session/token handling notes
- **`setAuditLogging(...)`** — rows of `{ event, fields }` — which events
  get logged and what fields are captured
- **`setComplianceRequirements(...)`** — bulleted data-retention policy and
  regulatory scope (GDPR, HIPAA, PCI-DSS, or none)
- **`setVulnerabilityManagement(...)`** — rows of `{ activity, frequency }`
  — dependency scans, penetration test cadence
- **`setSecurityChecklist(...)`** — bulleted pre-release checklist items
  (rendered with ☐ checkboxes)
- **`setRevisionHistory(...)`** — rows of `{ version, date, author,
changes }`

**6b. Data Model / ERD**

- **`setErdOverview(...)`** — what data domain this ERD covers and the
  target database engine (e.g. PostgreSQL 16)
- **`setEntityList(...)`** — rows of `{ entity, description }` — every
  table/collection with a one-line purpose
- **`setErdDiagramNote(...)`** — placeholder/link note for the actual ERD
  diagram image
- **`setEntityDetails(...)`** — array of `{ tableName, columns: [{ column,
type, constraints, description }] }` — repeat per entity, precise enough
  that a developer could write the `CREATE TABLE` statement directly from
  it
- **`setRelationships(...)`** — rows of `{ from, to, cardinality, rule }`
  — every foreign-key relationship and the business rule it represents
- **`setNormalisationNotes(...)`** — confirm normal form or explain
  deliberate denormalisation
- **`setIndexingStrategy(...)`** — rows of `{ table, index, reason }`

**6c. API Specification**

- **`setApiOverview(...)`** — API purpose, style (REST/GraphQL), intended
  consumers
- **`setApiBaseUrls(...)`** — array of lines (e.g. "Production:
  https://api.example.com/v1")
- **`setApiVersioningNote(...)`** — how versioning works and what
  constitutes a breaking change
- **`setApiAuthNote(...)`** — auth scheme (JWT/API key/OAuth) and how to
  obtain a token
- **`setApiAuthHeader(...)`** — array of code lines showing the auth header
  format
- **`setApiRolesNote(...)`** — role-based access summary
- **`setApiResponseFormat(...)`** — array of JSON lines showing the common
  success/error response envelope
- **`setApiErrorCodes(...)`** — rows of `{ status, meaning }` — HTTP status
  codes and what each means for this API
- **`setEndpoints(...)`** — array of `{ method, path, description,
requestBody: [...lines], responseBody: [...lines], errors }` — repeat
  per endpoint, grouped by resource, mirroring the ERD's entity structure
- **`setPaginationNote(...)`** — query parameters and defaults for list
  endpoints
- **`setRateLimiting(...)`** — rows of `{ limit, window, header }`
- **`setApiChangelog(...)`** — rows of `{ version, date, change }` — the
  API's own wire-contract version history, distinct from the product-wide
  Changelog in Part 11
- **`setMetadata(...)`** applies to section 6a only; **`setHeaderFooterLabels(...)`**
  applies across all three sub-documents

### Part 7 — UI/UX Documentation (`uiux-docx.js`)

- **`setOverviewBullets(...)`** — bulleted: primary user goal, entry
  points, what "done" looks like
- **`setUserFlowNote(...)`** — link to the interactive prototype plus a
  step-by-step happy-path narration
- **`setColorPalette(...)`** — rows of `{ role, hex, usage }` — Primary,
  Secondary, Background, Surface, Text roles, Success/Warning/Error
- **`setTypography(...)`** — rows of `{ style, spec, usage }` — Display/H1,
  H2/H3, Body, Caption/Label, each with font/size/weight
- **`setComponents(...)`** — rows of `{ component, states, notes }` — if
  using a component library (shadcn, Material, Figma AI), reference the
  source library instead of redefining from scratch
- **`setSpacingGridNote(...)`** — spacing scale, grid system, max content
  width
- **`setComponentBehaviour(...)`** — bulleted interaction notes (hover,
  tap, drag)
- **`setEdgeCases(...)`** — bulleted UI edge cases (long text, network
  timeouts, empty states)
- **`setAnimationNote(...)`** — transition specs, duration, easing
- **`setContentCopy(...)`** — rows of `{ element, copy, notes }` — exact
  UI copy strings with character limits/tone notes
- **`setAccessibilityHeading(...)`** — optional override of the section 6
  heading text (defaults to the template's own wording)
- **`setAccessibilityBullets(...)`** — bulleted: color contrast (WCAG 2.1
  AA), keyboard navigation, screen reader labels, touch target sizes
- **`setResponsiveNotes(...)`** — rows of `{ platform, notes }` — Desktop/
  Web, Mobile Web, Native iOS/Android layout differences
- **`setRevisionHistory(...)`** — rows of `{ version, date, author,
changes }`
- **`setMetadata(...)`** / **`setHeaderFooterLabels(...)`** — as above

### Part 8 — User Acceptance Testing / UAT (`uat-docx.js`)

- **`setObjectivesScope(...)`** — what this test plan verifies and why —
  a quality goal, not just "find bugs"
- **`setFeaturesToTest(...)`** — bulleted list of features/flows in scope
- **`setTestStrategy(...)`** — rows of `{ level, description, owner }` —
  Unit, Integration, System, UAT levels and who owns each
- **`setTestEnvironments(...)`** — rows of `{ environment, url, database,
notes }`
- **`setTestAccounts(...)`** — bulleted: where test credentials live
  (never paste passwords directly), supported devices/browsers
- **`setTestSchedule(...)`** — rows of `{ activity, start, end }`
- **`setRolesResponsibilities(...)`** — rows of `{ role, responsibility }`
  — QA Engineer, Dev Lead, Business/UAT Rep, etc.
- **`setTestCases(...)`** — rows of `{ id, relatedUseCase, steps,
expectedResult, status }` — one row per test case, ideally traced back
  to a BRD use case ID
- **`setNonFunctionalTesting(...)`** — rows of `{ type, criteria,
required, status }` — Performance, Accessibility, Security criteria
- **`setBugSeverityDefinitions(...)`** — rows of `{ severity, definition }`
  — S0/Blocker through S3/Minor
- **`setSignoff(...)`** — rows of `{ role, name, approved, date }`
- **`setMetadata(...)`** / **`setHeaderFooterLabels(...)`** — as above

### Part 9 — Deployment Guide (`deployment-guide-docx.js`)

- **`setOverview(...)`** — what's deployed, to where, how often
  (continuous deployment vs manual releases)
- **`setPrerequisites(...)`** — rows of `{ requirement, details }` — server
  access, runtime version, where environment variables come from
- **`setEnvironments(...)`** — rows of `{ environment, purpose, branch }`
  — Development, Staging, Production and their git branches
- **`setEnvExample(...)`** — array of lines forming a `.env.example`
  block (rendered as a shaded code block)
- **`setDeploymentSteps(...)`** — array of shell-command lines forming the
  actual deploy procedure; use empty strings `''` for blank separator
  lines between numbered steps
- **`setConfiguration(...)`** — rows of `{ variable, description,
required }` — every environment variable the app reads
- **`setRollbackSteps(...)`** — array of shell-command lines for rolling
  back to a previous release
- **`setRollbackNote(...)`** — a note on what to do if the rollback
  involves a database migration
- **`setPostDeploymentChecklist(...)`** — bulleted post-deploy checks
  (rendered with ☐ checkboxes) — health check endpoint, login smoke test,
  log review
- **`setMonitoring(...)`** — rows of `{ what, tool, alertRecipient }` —
  uptime, error logs, and who gets alerted
- **`setMetadata(...)`** / **`setHeaderFooterLabels(...)`** — as above

### Part 10 — User Manual (`user-manual-docx.js`)

- **`setIntroduction(...)`** — what the application is and who the manual
  is for, in plain language
- **`setGettingStarted(...)`** — how to access the system and log in for
  the first time, including prerequisites
- **`setGettingStartedScreenshot(...)`** — caption text for the login
  screenshot (e.g. "[Insert screenshot of the login screen]")
- **`setNavigating(...)`** — orientation to the main layout: sidebar,
  header, common buttons, before task walkthroughs
- **`setNavigatingScreenshot(...)`** — caption text for the annotated
  layout screenshot
- **`setFeatureWalkthroughs(...)`** — array of `{ taskTitle, steps,
screenshot }` — **one "How to [task]" block per feature/module**, each
  with numbered, task-oriented steps a first-time user could follow
  without assistance, plus an optional screenshot caption
- **`setFaq(...)`** — Q:/A: format frequently asked questions
- **`setTroubleshooting(...)`** — rows of `{ problem, solution }`
- **`setSupportContact(...)`** — how to reach support (channel or
  help-desk portal)
- **`setMetadata(...)`** / **`setHeaderFooterLabels(...)`** — as above

### Part 11 — Changelog (`changelog-docx.js`)

- **`setUnreleased(...)`** — bulleted list of merged-but-not-yet-shipped
  changes, each prefixed with a tag: ADDED, CHANGED, FIXED, DEPRECATED,
  REMOVED, or SECURITY
- **`setReleases(...)`** — array of `{ heading, entries }`, **newest
  release first**. `heading` follows "vX.Y.Z [YYYY-MM-DD] — Optional
  one-line theme" (Semantic Versioning: breaking change → MAJOR, new
  feature → MINOR, bug fix → PATCH). `entries` are tagged bullets in the
  same ADDED/CHANGED/FIXED/... format as Unreleased
- **`setMetadata(...)`** / **`setHeaderFooterLabels(...)`** — as above
- (The "How to Use This Changelog" guidance and the Category Legend table
  are fixed template content and don't need input.)

### Part 12 — Change Request Log (`change-request-log-docx.js`)

- **`setChangeRequests(...)`** — rows of `{ crId, dateRequested,
requestedBy, description, affectedFlow, status }` — one row per change
  request as they come in (e.g. CR-001, date, name/role of requester,
  description of the change, which flow/feature it affects, and its
  current status such as Pending/Approved/Rejected)
- **`setMetadata(...)`** / **`setHeaderFooterLabels(...)`** — as above

### Part 13 — Glossary (`glossary-docx.js`)

- **`setTerms(...)`** — rows of `{ term, definition }` — one consolidated
  glossary for the whole document: business terms (in plain language),
  acronyms (spelled out), and technical jargon your business and technical
  readers might otherwise define differently
- **`setMetadata(...)`** / **`setHeaderFooterLabels(...)`** — as above

### Part 14 — Appendix (`appendix-docx.js`)

- **`setResources(...)`** — rows of `{ resource, link }` — links out to
  living/detailed resources instead of duplicating them in the document:
  timeline/project tracker, design files (Figma), full API reference
  (OpenAPI/Swagger), data schema/ERD source, runbooks, research
  repository, analytics dashboard, CI/CD pipeline, incident/postmortem log
- **`setMetadata(...)`** / **`setHeaderFooterLabels(...)`** — as above

---

## 6. Tips

- **Dates and free-text fields are plain strings** — pass them exactly as
  you want them to appear (e.g. `'Aug 31, 2026'`); there's no date parsing
  or reformatting happening internally.
- **Table-shaped sections are arrays of plain objects** — the exact key
  names for each method are documented above and in that method's JSDoc
  comment / `sectionGuide()` entry. Match the key names exactly (e.g.
  `{ requirement: '...', tracesTo: '...' }` for SRS functional
  requirements) — extra keys are ignored, and missing keys render as a
  blank cell rather than throwing an error.
- **Bulleted-list sections are just arrays of strings** —
  `['First point', 'Second point']` — no object wrapper needed.
- **Multi-line content inside a single table cell**: put `\n` inside the
  string; each line becomes its own paragraph within that cell.
- **Code-block sections** (repo trees, shell commands, JSON examples,
  `.env` files) take an **array of lines**, rendered as one continuous
  shaded monospace block. Use an empty string `''` as an array entry
  wherever you want a blank line inside that same block (e.g. to separate
  numbered steps in a deployment script) — don't start a new method call
  for that, since each call renders as its own separate code box.
  This applies to: `setCodebaseTree`, `setInputValidationCode`,
  `setEnvExample`, `setDeploymentSteps`, `setRollbackSteps`,
  `setApiBaseUrls`, `setApiAuthHeader`, `setApiResponseFormat`, and each
  endpoint's `requestBody`/`responseBody` arrays in `setEndpoints`.
- **Checklist sections** (Security Checklist in Part 6, Post-Deployment
  Verification in Part 9) take a plain array of strings — each one is
  automatically prefixed with a ☐ box, don't add your own checkbox
  character.
- **Always call `setHeaderFooterLabels({ productNameLabel: '...' })`** on
  every single Part you generate — it's the one call that has to be
  repeated across all 15 files, since each generates its own independent
  header. Skipping it leaves the placeholder text "Product name/logo" (or
  "Product Name" in the PIC Matrix) visible in that Part's running header.
- **RACI tables (BRD §6.3)**: the `values` array inside each `steps` entry
  must be the same length and in the same order as the `roles` array —
  position 0 in `values` is that step's marker for `roles[0]`, and so on.
- **Use-case IDs and traceability**: keep the same ID scheme across Parts
  — BR-xxx (BRD business requirements) → FR-MOD-xxx or FR-xxx (SRS/PRD
  functional requirements) → UC-xx (BRD use cases) → TC-xx (UAT test
  cases) — the templates cross-reference these IDs in their guidance text,
  so consistent numbering makes the whole document suite traceable
  end-to-end.
- **Hardware products**: none of these modules assume or force hardware
  content — Part 6 as built here is written for pure software products. If
  your product has a physical hardware component, say so and the SDK can
  be extended with hardware-specific sections (BOM, mechanical/electrical
  specs, certifications, etc.), similar to what exists in the earlier
  Product Specification SDK's optional hardware section.
- **Don't hand-edit `pspt-core.js`** unless you deliberately want to change
  fonts, colors, or page geometry for the _entire_ suite at once — every
  other file imports its constants from there, so a change there ripples
  through all 15 Parts.
