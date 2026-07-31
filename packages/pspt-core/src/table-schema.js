'use strict';
/**
 * Generalized `{key, header, weight}` column-schema idiom shared by the docx
 * table builder and the xlsx task table. This module only distributes column
 * widths and resolves the real-rows-vs-placeholder-rows fallback; the actual
 * cell rendering stays format-specific (docx TableCell vs. xlsx worksheet cell).
 */

/**
 * @param {Array<{key:string, header:string, weight?:number}>} columns
 * @param {number} totalWidth
 * @returns {number[]} widths per column, summing exactly to totalWidth
 */
function distributeColumnWidths(columns, totalWidth) {
  const totalWeight = columns.reduce((s, c) => s + (c.weight || 1), 0);
  const widths = columns.map((c) => Math.floor((totalWidth * (c.weight || 1)) / totalWeight));
  const drift = totalWidth - widths.reduce((a, b) => a + b, 0);
  widths[widths.length - 1] += drift;
  return widths;
}

/**
 * Resolves which row set to render and whether the placeholder convention
 * applies (see PLACEHOLDER_ROW_TEXT in ./tokens.js).
 * @param {Array<Object>|undefined} rows - real data rows
 * @param {Array<Object>} placeholderRows - fallback row(s)
 * @returns {{hasData:boolean, rows:Array<Object>}}
 */
function resolveRows(rows, placeholderRows) {
  const hasData = Array.isArray(rows) && rows.length > 0;
  return { hasData, rows: hasData ? rows : (placeholderRows || []) };
}

module.exports = { distributeColumnWidths, resolveRows };
