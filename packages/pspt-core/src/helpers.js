'use strict';
/**
 * Shared border/date/color helpers used by both the docx and xlsx adapters.
 */

const { XLSX_ROW_PRESETS, XLSX_COLORS, DOCX_COLOR } = require('./tokens');

// ---------------------------------------------------------------------------
// docx border helper
// ---------------------------------------------------------------------------

/** Builds the standard single 0.5pt black border set used by every docx table. */
function docxBorderSet(docx) {
  const { BorderStyle } = docx;
  const edge = { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLOR.border };
  return {
    top: edge,
    bottom: edge,
    left: edge,
    right: edge,
    insideHorizontal: edge,
    insideVertical: edge,
  };
}

// ---------------------------------------------------------------------------
// xlsx border helper
// ---------------------------------------------------------------------------

/** Applies the standard thin grid border to an exceljs cell. */
function applyXlsxBorder(cell, gridBorderArgb) {
  const color = gridBorderArgb || XLSX_COLORS.gridBorder;
  cell.border = {
    top: { style: 'thin', color: { argb: color } },
    left: { style: 'thin', color: { argb: color } },
    bottom: { style: 'thin', color: { argb: color } },
    right: { style: 'thin', color: { argb: color } },
  };
}

// ---------------------------------------------------------------------------
// date helpers (xlsx)
// ---------------------------------------------------------------------------

function parseDate(dateVal) {
  if (!dateVal) return new Date();
  if (dateVal instanceof Date) return dateVal;
  if (typeof dateVal === 'string') {
    const parts = dateVal.split('-');
    if (parts.length === 3) {
      return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    }
  }
  return new Date(dateVal);
}

function resetTime(date) {
  const reset = new Date(date.getTime());
  reset.setHours(0, 0, 0, 0);
  return reset;
}

function formatDateDisplay(date) {
  if (!date) return '';
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

// ---------------------------------------------------------------------------
// color resolution (xlsx)
// ---------------------------------------------------------------------------

/**
 * Resolves a color input (preset name or raw hex) into an ARGB hex string.
 * @param {string} colorInput
 * @param {string} defaultHex
 * @param {Object} [rowPresets] - defaults to XLSX_ROW_PRESETS
 */
function resolveColor(colorInput, defaultHex, rowPresets) {
  const presets = rowPresets || XLSX_ROW_PRESETS;
  if (!colorInput) return defaultHex.startsWith('FF') ? defaultHex : `FF${defaultHex}`;
  if (presets[colorInput]) return presets[colorInput];
  const hex = colorInput.replace('#', '');
  return hex.length === 6 ? `FF${hex}` : hex;
}

module.exports = {
  docxBorderSet,
  applyXlsxBorder,
  parseDate,
  resetTime,
  formatDateDisplay,
  resolveColor,
};
