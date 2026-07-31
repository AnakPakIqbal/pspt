'use strict';
/**
 * =============================================================================
 * DESIGN TOKENS — shared across pspt-docx and pspt-xlsx.
 * Extracted directly from the original product-spec-sdk.js / excel-tracker-sdk.js
 * source templates. Do not change these unless the source templates change.
 * =============================================================================
 */

const PLACEHOLDER_ROW_TEXT = '[...]';

// ---------------------------------------------------------------------------
// docx tokens (product-spec-sdk.js)
// ---------------------------------------------------------------------------

const DOCX_FONT = 'Arial';

const DOCX_COLOR = {
  h1: '1f3864', // Heading 1 (navy)
  h2: '2e5395', // Heading 2 (blue)
  h3: '3d3d3d', // Heading 3 (dark grey)
  h4: '2e74b5', // Heading 4 (italic blue)
  purposeLabel: '2e5395', // "Purpose: " label
  purposeText: '6b6b6b', // Purpose description text
  coverValue: '8a8a8a', // cover page placeholder values (grey italic)
  tableHeaderBg: '1f3864', // table header row fill
  tableHeaderText: 'ffffff', // table header row text
  border: '000000',
  black: '000000',
  headerFooterText: '6b6b6b',
  confidential: 'cc0000',
  imageBoxFill: 'CFE2F3', // cover placeholder graphic fill
};

const DOCX_SIZE = {
  // half-points (Word convention: 22 half-points = 11pt)
  appName: 80, // 40pt — cover page application name
  docTitle: 40, // 20pt — "Product Specification" on cover
  h1: 32, // 16pt
  h2: 26, // 13pt
  h3: 23, // 11.5pt
  h4: 22, // 11pt (italic)
  body: 22, // 11pt
  headerFooter: 18, // 9pt
};

const DOCX_PAGE = {
  width: 12240,
  height: 15840, // US Letter, portrait (DXA)
  margin: { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 708, footer: 708 },
};

// Usable content width for tables (US Letter, 1in margins both sides)
const DOCX_CONTENT_WIDTH = DOCX_PAGE.width - DOCX_PAGE.margin.left - DOCX_PAGE.margin.right; // 9360 dxa = 6.5in

// ---------------------------------------------------------------------------
// xlsx tokens (excel-tracker-sdk.js)
// ---------------------------------------------------------------------------

const XLSX_COLORS = {
  navyHeader: 'FF1B365D', // Main header bar (Navy Blue)
  darkGreenBanner: 'FF2D7D46', // Section header (Dark Green)
  redWeekend: 'FFCC0000', // Weekend column (Red)
  blueGantt: 'FF4A90E2', // Gantt timeline bar (Blue)
  gridBorder: 'FF595959', // Soft black grid border
  white: 'FFFFFFFF',
  darkText: 'FF333333',
};

// Color presets for section rows
const XLSX_ROW_PRESETS = {
  green: 'FFE2EFDA',
  yellow: 'FFF2CC',
  pink: 'FFFCE4D6',
  blue: 'FFDDEBF7',
  purple: 'FFEAD1DC',
  gray: 'FFF2F2F2',
};

module.exports = {
  PLACEHOLDER_ROW_TEXT,
  DOCX_FONT,
  DOCX_COLOR,
  DOCX_SIZE,
  DOCX_PAGE,
  DOCX_CONTENT_WIDTH,
  XLSX_COLORS,
  XLSX_ROW_PRESETS,
};
