/**
 * Example: the Product Documentation Master, driven entirely through the SDK's
 * public functions. This demonstrates the shape every method expects — swap in
 * your own product's content; the styling never changes.
 *
 * Two ways to use the suite are shown:
 *   1. `MasterDocument` — fill in whichever Parts you have content for and get
 *      one assembled document. Parts you never touch still render, as their own
 *      blank template, so the output is always complete.
 *   2. A single Part on its own, when you only need (say) the Project Brief.
 *
 * Run with `npm run example:docx` from the repo root.
 */
const { MasterDocument, ProjectBriefSDK } = require('pspt-docx');

const PRODUCT = 'Alban Translator';

function fillProjectBrief(brief) {
  brief
    .setMetadata({
      writer: 'Jane Doe',
      status: 'Draft',
      version: 'V1 (Phase 1)',
      lastUpdate: 'Sept 1, 2026',
    })
    .setOverview(
      'Alban Translator turns spoken Albanian into subtitled English in real time, ' +
        'so mixed-language teams can run meetings without an interpreter in the room.',
    )
    .setBackgroundPains([
      'Meetings stall while someone paraphrases for the non-Albanian speakers',
      'Recordings are never transcribed, so decisions get lost',
    ])
    .setObjectives([
      'Deliver subtitles within 2 seconds of speech',
      'Reach 95% word accuracy on conversational speech',
    ])
    .setKeyModules([
      { module: 'Capture', features: 'Live microphone capture with speaker diarisation' },
      { module: 'Translate', features: 'Streaming Albanian to English with punctuation' },
      { module: 'Archive', features: 'Searchable transcript stored per meeting' },
    ])
    .setTimeline([
      { phase: 'Phase 1 — Discovery & Design', duration: '2 weeks' },
      { phase: 'Phase 2 — MVP', duration: '6 weeks' },
      { phase: 'Phase 3 — Testing & UAT', duration: '3 weeks' },
    ])
    .setDeliverables([
      'Desktop application (macOS/Windows)',
      'Technical documentation (this document)',
      'Onboarding session for meeting hosts',
    ])
    .setPreliminaryRisks([
      {
        risk: 'Dialect coverage outside Tosk/Gheg',
        mitigation: 'Scope Phase 1 to standard Albanian',
      },
      { risk: 'Latency budget under poor network', mitigation: 'Fall back to on-device model' },
    ]);
}

function fillBrd(brd) {
  brd
    .setPurpose(
      'Defines what the business needs Alban Translator to do before any build work starts.',
    )
    .setBusinessObjectives(['Cut meeting overrun caused by live paraphrasing to near zero'])
    .setScope({
      inScope: 'Live meeting translation and transcript archival.',
      outOfScope: 'Document translation, and any language pair other than Albanian to English.',
    })
    .setBusinessRequirements([
      {
        id: 'BR-001',
        requirement: 'The system shall display subtitles during a live meeting',
        objective: 'Cut meeting overrun',
        priority: 'Must',
      },
      {
        id: 'BR-002',
        requirement: 'The system shall store a searchable transcript per meeting',
        objective: 'Preserve decisions',
        priority: 'Should',
      },
    ])
    .setActors([
      { actor: 'Meeting host', description: 'Starts and ends the translation session' },
      { actor: 'Participant', description: 'Reads subtitles; does not configure anything' },
    ])
    .setUseCaseList([
      {
        id: 'UC-01',
        name: 'Start a translated meeting',
        actor: 'Meeting host',
        description: 'Opens a session and begins capture',
      },
    ]);
}

function fillGlossary(glossary) {
  glossary.setTerms([
    { term: 'Diarisation', definition: 'Splitting an audio stream by who is speaking' },
    { term: 'Tosk / Gheg', definition: 'The two principal dialect groups of Albanian' },
    { term: 'WER', definition: 'Word Error Rate — the accuracy measure used for transcription' },
  ]);
}

async function main() {
  // ---------------------------------------------------------------------
  // 1. The assembled master document.
  // ---------------------------------------------------------------------
  const master = new MasterDocument();

  // The one call worth making for every Part — it fills the running header.
  master.setHeaderFooterLabels({ productNameLabel: PRODUCT });

  master.part('picMatrix').setCoverInfo({
    productName: PRODUCT,
    productType: 'Desktop application',
    status: 'In Development',
    writer: 'Jane Doe',
    lastUpdated: 'Sept 1, 2026',
  });

  fillProjectBrief(master.part('projectBrief'));
  fillBrd(master.part('brd'));
  fillGlossary(master.part('glossary'));

  // Every Part left untouched above still renders, as its blank template.
  await master.generate('./alban-translator-documentation.docx');
  console.log('wrote ./alban-translator-documentation.docx');

  // ---------------------------------------------------------------------
  // 2. The same Project Brief on its own, as a standalone file.
  // ---------------------------------------------------------------------
  const brief = new ProjectBriefSDK();
  brief.setHeaderFooterLabels({ productNameLabel: PRODUCT });
  fillProjectBrief(brief);
  await brief.generate('./alban-translator-project-brief.docx');
  console.log('wrote ./alban-translator-project-brief.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
