/**
 * Example: Generating a Project Tracker & Gantt Chart spreadsheet (.xlsx)
 * matching the user's reference design with colored rows, red weekend columns,
 * and blue timeline bars.
 */
const path = require('path');
const ExcelTrackerSDK = require('pspt-xlsx');

async function main() {
  const tracker = new ExcelTrackerSDK();

  tracker.setTitle('Project Tracker & Gantt Chart');

  // Section 1: Project Tracker (Light Green & Light Yellow & Light Pink rows)
  tracker.addSection({
    title: 'Project Tracker',
    bannerColor: '2D7D46', // Dark Green banner
    rowColor: 'green',     // Light green rows
    tasks: [
      { no: 1, name: 'Frontend Architecture Design', detail: '', checklist: 'Completed', mandays: 1, total: 1, startDate: '2026-06-09', endDate: '2026-06-09', deliverables: 'Design Arc', notes: '', check: true },
      { no: 2, name: 'Frontend UI Implementation', detail: '', checklist: 'Completed', mandays: 1, total: 1, startDate: '2026-06-10', endDate: '2026-06-10', deliverables: '', notes: '', check: true },
    ],
  });

  tracker.addSection({
    title: 'Backend Core Tasks',
    bannerColor: '2D7D46',
    rowColor: 'yellow',    // Light yellow rows
    tasks: [
      { no: 3, name: 'Backend Architecture Design', detail: '', checklist: 'Completed', mandays: 1, total: 1, startDate: '2026-06-11', endDate: '2026-06-11', deliverables: 'Design Arc', notes: '', check: true },
      { no: 4, name: 'Backend Auth Design', detail: '', checklist: 'Completed', mandays: 1, total: 1, startDate: '2026-06-12', endDate: '2026-06-12', deliverables: 'Granular RBAC', notes: '', check: true },
      { no: 5, name: 'Backend Implementation', detail: '', checklist: 'Completed', mandays: 1, total: 1, startDate: '2026-06-15', endDate: '2026-06-15', deliverables: '', notes: '', check: true },
    ],
  });

  tracker.addSection({
    title: 'Integrations',
    bannerColor: '2D7D46',
    rowColor: 'green',     // Light green rows
    tasks: [
      { no: 6, name: 'Auth Integration', detail: '', checklist: 'Completed', mandays: 1, total: 1, startDate: '2026-06-17', endDate: '2026-06-17', deliverables: '', notes: '', check: true },
      { no: 7, name: 'Project Integration', detail: '', checklist: 'Completed', mandays: 2, total: 2, startDate: '2026-06-18', endDate: '2026-06-19', deliverables: '', notes: '', check: true },
      { no: 8, name: 'Activity Integration', detail: '', checklist: 'Completed', mandays: 2, total: 2, startDate: '2026-06-22', endDate: '2026-06-23', deliverables: '', notes: '', check: true },
      { no: 9, name: 'Logs Integration', detail: '', checklist: 'Completed', mandays: 2, total: 2, startDate: '2026-06-24', endDate: '2026-06-25', deliverables: '', notes: '', check: true },
    ],
  });

  tracker.addSection({
    title: 'AI & Automation Tasks',
    bannerColor: '2D7D46',
    rowColor: 'pink',      // Light pink rows
    tasks: [
      { no: 12, name: 'AI Integration', detail: '', checklist: 'Completed', mandays: 2, total: 2, startDate: '2026-06-29', endDate: '2026-06-30', deliverables: 'R&D', notes: '', check: true, commitCount: 5, linesAdded: 320, linesRemoved: 40 },
      { no: 13, name: 'AI Project Brainstorming', detail: '', checklist: 'Completed', mandays: 2, total: 2, startDate: '2026-07-01', endDate: '2026-07-02', deliverables: '', notes: '', check: true },
      { no: 14, name: 'AI Project Automate Creation', detail: '', checklist: 'Completed', mandays: 2, total: 2, startDate: '2026-07-03', endDate: '2026-07-06', deliverables: '', notes: '', check: true },
    ],
  });

  // Section 2: Updated Role and Field (spanning 29 July till 3 August as requested!)
  tracker.addSection({
    title: 'Updated Role and Field',
    bannerColor: '2D7D46',
    rowColor: 'yellow',    // Light yellow rows
    tasks: [
      { no: 1, name: 'Redesign Organization Hierarchy (Design)', detail: '', checklist: 'Completed', mandays: 2, total: 2, startDate: '2026-06-24', endDate: '2026-06-29', deliverables: '', notes: '', check: true },
      { no: 2, name: 'Redesign Organization Hierarchy (Implementation)', detail: '', checklist: 'Completed', mandays: 2, total: 2, startDate: '2026-06-24', endDate: '2026-06-29', deliverables: '', notes: '', check: true },
      { no: 3, name: 'Enhanced AI Integration with Gemini 2.5', detail: '', checklist: 'Completed', mandays: 6, total: 6, startDate: '2026-07-29', endDate: '2026-08-03', deliverables: 'Gemini 2.5', notes: 'Spans July 29 to Aug 3', check: true },
      { no: 4, name: 'Review and enhance security and code optimization', detail: '', checklist: 'Completed', mandays: 2, total: 2, startDate: '2026-06-30', endDate: '2026-07-01', deliverables: '', notes: '', check: true },
      { no: 5, name: 'Auto select project at task/subtask creation', detail: '', checklist: 'Completed', mandays: 1, total: 1, startDate: '2026-07-01', endDate: '2026-07-01', deliverables: '', notes: '', check: true },
    ],
  });

  // addCallout is pre-existing dead code (stored but never rendered) — exercised
  // here for parity with the original SDK's exact (non-)behavior.
  tracker.addCallout({ dateStr: '2026-07-04', text: 'US Holiday', color: 'CC0000' });

  const outputPath = path.join(__dirname, 'output', 'project-tracker-gantt.xlsx');
  await tracker.generate(outputPath);
  console.log(`Successfully generated Project Tracker Excel sheet at: ${outputPath}`);
}

main().catch(err => {
  console.error('Error generating Excel file:', err);
  process.exit(1);
});
