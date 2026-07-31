/**
 * Example: fully populated Product Specification, driven entirely through
 * the SDK's public functions. This demonstrates the shape every method
 * expects. Swap in your own product's content — the styling never changes.
 */
const ProductSpecSDK = require('pspt-docx');

async function main() {
  const doc = new ProductSpecSDK();

  doc.setCoverPage({
    productName: 'Alban Translator',
    shortDescription: 'Real-time push-to-talk voice translation and meeting transcription.',
    lastUpdated: '2025-06-30',
    status: 'In Review',
  });

  doc.setProductInfo({
    productName: 'Alban Translator',
    version: 'v2.0 (Landing Page revision — prior: v2.0 Admin/Reseller, v1.0 Initial Handoff)',
    status:
      'In Progress for production cost — Reseller Landing Page & Dashboard, Alban Revision Landing Page, and Admin Page in active build (June 2025 sprint)',
  });

  doc.setExecutiveSummary(
    'Alban Translator provides real-time push-to-talk voice translation for guest/customer-service interactions, ' +
      'plus meeting recording with live transcription and translated summaries ("Smart Minutes"). A reseller channel ' +
      'lets partners run their own branded admin and landing pages on top of the same platform.',
  );

  doc.setProductRoadmap([
    { phase: 'Handoff v1.0', theme: 'Initial Handoff', timeframe: '2024', status: 'Complete' },
    {
      phase: 'v2.0',
      theme: 'Landing page, Admin page, Reseller admin, Reseller landing page',
      timeframe: '2025',
      status: 'Complete',
    },
  ]);

  doc.setTargetMarket([
    {
      segment: 'End users / Guests',
      description:
        'Individuals needing real-time push-to-talk voice translation (e.g., hotel front-desk / guest check-in scenario).',
    },
    {
      segment: 'Business customers',
      description:
        'Companies running international meetings needing live transcription + translated summaries ("Smart Minutes").',
    },
    {
      segment: 'Resellers / Partners',
      description: 'Partners distributing Alban under their own branded admin & landing pages.',
    },
  ]);

  doc.setCustomerPersonas([
    {
      persona: 'Guest / Front-desk user',
      role: 'Traveler or hotel guest',
      goals: 'Communicate instantly in an unfamiliar language',
      painPoints: 'No shared language with staff',
      buyingBehavior: 'Free trial → light usage',
    },
    {
      persona: 'Business / Meeting host',
      role: 'Team lead running cross-border meetings',
      goals: 'Accurate live meeting transcription & translated summaries',
      painPoints: 'Manual note-taking, lost context across languages',
      buyingBehavior: 'Business ($50) / Enterprise ($100) credit plans',
    },
    {
      persona: 'Reseller Admin',
      role: 'Partner managing sub-accounts',
      goals: 'Manage users, credits, and view conversation history',
      painPoints: 'Needs its own branded dashboard & controls',
      buyingBehavior: 'Reseller admin panel access',
    },
  ]);

  doc.setUserJourney({
    caption: 'Guest push-to-talk flow',
  });

  doc.setUseCases([
    {
      useCase: 'Push-to-talk live voice translation between guest and staff',
      actor: 'Guest / Customer Service',
      trigger: 'User taps mic / Push to Talk',
      outcome: 'Translated speech shown/spoken in the target language',
    },
    {
      useCase: 'Record New Meeting with live transcription & translation',
      actor: 'Meeting host',
      trigger: 'User starts "Record New Meeting", selects input/output language',
      outcome: 'Live Transcription + Live Translated Transcription displayed and saved',
    },
    {
      useCase: 'Monitor account usage',
      actor: 'Business user / Admin',
      trigger: 'User opens Dashboard',
      outcome:
        "See credit balance, active users, today's/total conversations, usage chart, top languages used",
    },
    {
      useCase: 'Manage reseller sub-accounts',
      actor: 'Reseller Admin',
      trigger: 'Admin logs into Reseller Dashboard',
      outcome: 'Views/edits users, credits, referral history, withdrawals',
    },
  ]);

  doc.setCompetitiveAnalysis([
    {
      competitor: 'Generic Translation App',
      strengths: 'Wide language support',
      weaknesses: 'No meeting-specific workflow',
      pricing: 'Freemium',
      differentiator: 'Push-to-talk UX built for service interactions',
    },
  ]);

  doc.setPricingModel({
    modelDescription: 'Subscription/credit-based plans plus a reseller channel.',
    tiers: [
      {
        tier: 'Business',
        price: '$50',
        includes: 'Meeting transcription credits',
        targetSegment: 'Business teams',
      },
      {
        tier: 'Enterprise',
        price: '$100',
        includes: 'Higher credit allotment, priority support',
        targetSegment: 'Larger orgs',
      },
    ],
  });

  doc.setRevenueModel(
    'Subscription/credit-based SaaS revenue, plus a reseller channel where partners resell access and manage their ' +
      'own end users (referral/withdrawal tracked in Reseller Dashboard).',
  );

  doc.setFeatures([
    {
      name: 'Push-to-Talk Live Translation',
      description:
        'Real-time voice translation between two speakers (guest/customer service) with flag + language selector.',
      priority: 'Must',
    },
    {
      name: 'Record New Meeting / Smart Minutes',
      description:
        'Records a meeting, selects input/output languages, and produces Live Transcription + Live Translated Transcription.',
      priority: 'Must',
    },
    {
      name: 'User Dashboard',
      description:
        "Shows credit balance, active users, today's/total conversations, conversations-overview chart, top languages used, user management list.",
      priority: 'Must',
    },
    {
      name: 'Reseller Admin Panel',
      description:
        'Reseller-branded dashboard: credits, users, withdrawal, referral history; integrates with Alban Admin Panel.',
      priority: 'Should',
    },
    {
      name: 'Landing Page / Marketing Site',
      description:
        'Public site — Home, Pricing, How It Works, Contact Us, Free Trial, Sign Up/Sign In.',
      priority: 'Must',
    },
    {
      name: 'Multi-language support',
      description: 'English, Indonesian, Chinese, Arabic, Japanese, Dutch, French.',
      priority: 'Must',
    },
  ]);

  doc.setSystemArchitecture({
    description: 'API gateway fronting a Node.js backend with a translation pipeline microservice.',
  });

  doc.setTechnologyStack([
    {
      layer: 'Frontend',
      technology: 'React / Next.js',
      justification: 'Fast iteration on the marketing site and dashboards.',
    },
    {
      layer: 'Backend',
      technology: 'Node.js',
      justification: 'Shared language with frontend team.',
    },
  ]);

  doc.setApis({
    rows: [
      {
        endpoint: '/api/v1/auth',
        method: 'HTTP',
        description: 'Authentication',
        authRequired: 'No',
        role: 'None',
      },
      {
        endpoint: '/api/v1/users',
        method: 'HTTP',
        description: 'User Management Endpoint',
        authRequired: 'Yes',
        role: 'Admin',
      },
    ],
    docsLink: 'Postman collection — internal link',
  });

  doc.setDatabaseDesign({
    rows: [{ table: 'users', purpose: 'User accounts', keyFields: 'id, email', url: '' }],
  });

  doc.setAuthentication('JWT-based session auth; reseller admin accounts use scoped roles.');

  doc.setSecurity({
    measures: [
      { title: 'Cloudflare', children: ['Turnstile'] },
      { title: 'Cloud Armor', children: ['Rate limiting / WAF rules'] },
    ],
    notes: 'TLS in transit; secrets stored in a managed secrets vault.',
  });

  doc.setMonitoring('Uptime and error-rate dashboards for the API and translation pipeline.');
  doc.setLogging('Centralized structured logging; PII masked in conversation logs.');
  doc.setBackupStrategy('Nightly database backups, 30-day retention.');

  // Hardware section exercised deliberately for parity coverage.
  doc.setEnableHardwareSection(true);
  doc.setHardwareOverview('A pocket translator companion device (optional add-on hardware).');
  doc.setHardwareComponents([
    { component: 'MCU', function: 'Main compute', vendor: 'Nordic', altVendor: 'Espressif' },
  ]);
  doc.setBillOfMaterials([
    {
      partNo: 'R-1001',
      description: '10k resistor',
      qty: '4',
      unitCost: '$0.01',
      vendor: 'Digikey',
      leadTime: '2 weeks',
    },
  ]);
  doc.setMechanicalDesign({ description: 'ABS enclosure, IP54.' });
  doc.setElectricalSpecification({ description: '5V input, 200mA typical draw.' });
  doc.setSensors([
    {
      sensor: 'Microphone array',
      type: 'MEMS',
      range: 'n/a',
      accuracy: 'n/a',
      purpose: 'Voice capture',
    },
  ]);
  doc.setConnectivity('BLE 5.0 for pairing, Wi-Fi for OTA updates.');
  doc.setPowerRequirements('Li-ion 500mAh, ~7 days typical use, USB-C charging.');
  doc.setFirmware('A/B partition OTA with automatic rollback on boot failure.');
  doc.setCertifications([
    { certification: 'FCC', region: 'US', status: 'Pending', targetDate: 'Q4 2026' },
  ]);
  doc.setEnvironmentalRequirements('-10C to 45C operating range, IP54, 1m drop tested.');
  doc.setManufacturingNotes('Contract manufactured in Shenzhen; target yield 98%.');
  doc.setMaintenance('Battery is field-replaceable; 2-year expected lifespan.');

  doc.setDesignTools([
    { category: 'Design File', tool: 'Figma', link: '[Figma Link — to be inserted]' },
    { category: 'Live Project', tool: 'Deployment', link: 'alban-translator.com' },
  ]);

  doc.setDesignPrinciples(
    'Clarity and speed of communication first: minimal-friction push-to-talk interaction, clear language/flag ' +
      'indicators, and dashboard views that surface usage (credits, conversations) at a glance.',
  );

  doc.setLayoutGrid([
    { property: 'Breakpoint (Website)', value: '1512 × 982 px' },
    { property: 'Grid System', value: '12-column grid' },
    { property: 'Margin', value: '120px' },
    { property: 'Gutter', value: '24px' },
  ]);

  doc.setTypography([
    { weight: 'Font family', sizes: 'Plus Jakarta Sans' },
    { weight: 'Regular', sizes: '12px, 16px, 20px, 24px' },
    { weight: 'Semibold', sizes: '12px, 16px, 20px, 24px, 32px' },
  ]);

  doc.setColorPalette([
    { role: 'Primary', color: 'Blue', hex: '#0081FF' },
    { role: 'Secondary', color: 'Pink/Magenta', hex: '#EE48B1' },
    { role: 'Black', color: 'Black', hex: '#000000' },
    { role: 'White', color: 'White', hex: '#FFFFFF' },
  ]);

  doc.setComponentsStates([
    { component: 'Primary Button', state: 'Default', behavior: 'Blue background, white text' },
  ]);
  doc.setResponsiveBehavior([
    { breakpoint: '<768px', device: 'Mobile', notes: 'Single column layout' },
  ]);
  doc.setInteractionAnimation([{ aspect: 'Transitions', notes: 'Fade in, 200ms' }]);

  doc.setRevisionHistory([
    { version: 'v1.0', date: '2024', changes: 'Initial Handoff' },
    {
      version: 'v2.0',
      date: '2025',
      changes: 'Landing page; Admin page; Reseller admin; Reseller landing page',
    },
    { version: 'v3.0', date: '2025', changes: 'Landing Page revision' },
  ]);

  doc.setTestStrategy(
    'Manual exploratory testing each sprint, plus smoke tests on staging before release.',
  );
  doc.setTestPlan([
    {
      phase: 'Regression',
      scope: 'Full app',
      entryCriteria: 'Feature complete',
      exitCriteria: 'Zero P1 bugs',
    },
  ]);
  doc.setTestCases([
    {
      id: 'TC-001',
      description: 'Push to talk translates speech',
      steps: '1. Tap mic 2. Speak 3. Release',
      expectedResult: 'Translated speech is played',
      status: 'Pass',
    },
  ]);
  doc.setBugTracking([{ severity: 'Critical', definition: 'Data loss or outage', sla: '4 hours' }]);
  doc.setSecurityTesting('Ad-hoc penetration testing before major releases.');
  doc.setHardwareValidation([
    {
      type: 'Drop Test',
      description: '1m onto concrete',
      standard: 'IEC 60068-2-32',
      result: 'Pass',
    },
  ]);

  await doc.generate(require('path').join(__dirname, 'output', 'filled_out.docx'));
  console.log('wrote filled_out.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
