'use strict';

const ExcelJS = require('exceljs');
const {
  XLSX_COLORS,
  XLSX_ROW_PRESETS,
  parseDate,
  resetTime,
  formatDateDisplay,
  resolveColor,
  applyXlsxBorder,
} = require('pspt-core');

/**
 * ExcelTrackerSDK
 * Programmatically generates Project Tracker & Gantt Chart spreadsheets
 * matching navy headers, colored category section banners, red weekend indicators,
 * blue Gantt timeline bars with Commit Count & Lines Added/Removed (+/-) tracking.
 */
class ExcelTrackerSDK {
  constructor() {
    this.workbook = new ExcelJS.Workbook();
    this.title = 'Project Tracker & Gantt Timeline';
    this.sections = [];
    this.customCalendarRange = null; // { startDate, endDate }
    this.callouts = []; // [{ dateStr, text, color }]
    this.expandFullMonths = true; // Auto expand calendar to 1st of month & last day of month
    this.minMonthsSpan = 3; // Default min 3 months span (e.g., June, July, August)

    // Design Tokens & Color Palettes
    this.colors = { ...XLSX_COLORS };

    // Color presets for section rows
    this.rowPresets = { ...XLSX_ROW_PRESETS };
  }

  /**
   * Set document main title
   * @param {string} title
   */
  setTitle(title) {
    this.title = title;
    return this;
  }

  /**
   * Optionally set an explicit calendar range (e.g. '2026-06-01' to '2026-08-31').
   * If omitted, range is automatically computed from min/max task dates.
   * @param {string|Date} startDate
   * @param {string|Date} endDate
   */
  setCalendarRange(startDate, endDate) {
    this.customCalendarRange = {
      start: this._parseDate(startDate),
      end: this._parseDate(endDate),
    };
    return this;
  }

  /**
   * Set minimum months span to render in the calendar.
   * @param {number} count
   */
  setMinMonthsSpan(count = 3) {
    this.minMonthsSpan = count;
    return this;
  }

  /**
   * Toggle auto-expansion of calendar to full month boundaries.
   * @param {boolean} enabled
   */
  setExpandFullMonths(enabled = true) {
    this.expandFullMonths = enabled;
    return this;
  }

  /**
   * Add a section with tasks and visual row styling.
   * @param {Object} section
   * @param {string} section.title - Section title e.g. 'Core Payment Engine'
   * @param {string} [section.bannerColor] - Hex or preset name for section banner
   * @param {string} [section.rowColor] - Hex or preset name for task rows background
   * @param {Array<Object>} section.tasks - List of tasks
   */
  addSection({ title, bannerColor = '2D7D46', rowColor = 'green', tasks = [] }) {
    this.sections.push({
      title,
      bannerColor: this._resolveColor(bannerColor, '2D7D46'),
      rowColor: this._resolveColor(rowColor, 'FFE2EFDA'),
      tasks: tasks.map((task) => ({
        no: task.no,
        name: task.name || '',
        detail: task.detail || '',
        checklist: task.checklist || 'Completed',
        mandays: task.mandays || 1,
        total: task.total || task.mandays || 1,
        startDate: this._parseDate(task.startDate),
        endDate: this._parseDate(task.endDate),
        deliverables: task.deliverables || '',
        notes: task.notes || '',
        commitCount: task.commitCount !== undefined ? task.commitCount : 1,
        lines:
          task.lines ||
          (task.linesAdded !== undefined
            ? `+${task.linesAdded} / -${task.linesRemoved || 0}`
            : '+0 / -0'),
        check: task.check !== undefined ? task.check : true,
      })),
    });
    return this;
  }

  /**
   * Add a custom callout text overlaid on a specific date column.
   * @param {Object} callout
   */
  addCallout({ dateStr, text, color = 'CC0000' }) {
    this.callouts.push({
      date: this._parseDate(dateStr),
      text,
      color: this._resolveColor(color, 'CC0000'),
    });
    return this;
  }

  /**
   * Generate the Excel file and write to disk
   * @param {string} filePath
   */
  async generate(filePath) {
    const worksheet = this.workbook.addWorksheet(this.title || 'Project Tracker');

    // 1. Calculate Calendar Dates
    const { dates, monthGroups } = this._computeCalendarTimeline();

    // 2. Setup Columns metadata (Left Table: Col A to M = 13 Columns)
    const leftColCount = 13;
    const totalCols = leftColCount + dates.length;

    // Set Column widths
    worksheet.getColumn(1).width = 6; // No
    worksheet.getColumn(2).width = 32; // FEATURE NAME
    worksheet.getColumn(3).width = 25; // FEATURE DETAIL
    worksheet.getColumn(4).width = 14; // Checklist
    worksheet.getColumn(5).width = 14; // MANDAYS
    worksheet.getColumn(6).width = 10; // TOTAL
    worksheet.getColumn(7).width = 14; // Start Date
    worksheet.getColumn(8).width = 14; // End Date
    worksheet.getColumn(9).width = 20; // Deliverables
    worksheet.getColumn(10).width = 22; // Notes
    worksheet.getColumn(11).width = 12; // Commit Count
    worksheet.getColumn(12).width = 18; // Line Added & Removed
    worksheet.getColumn(13).width = 8; // Check

    // Timeline column widths (small grid cells)
    for (let col = leftColCount + 1; col <= totalCols; col++) {
      worksheet.getColumn(col).width = 3.5;
    }

    // -------------------------------------------------------------
    // ROW 1 & ROW 2: MAIN HEADERS
    // -------------------------------------------------------------

    // Row 1: Navy Header Bar for Left Table + Month Banners for Right Timeline
    const row1 = worksheet.getRow(1);
    row1.height = 28;

    // Merge G1:H1 for DATELINE
    worksheet.mergeCells('G1:H1');

    // Left Table Headers
    const topHeaderLabels = [
      { col: 1, text: '' },
      { col: 2, text: 'FEATURE NAME' },
      { col: 3, text: 'FEATURE DETAIL' },
      { col: 4, text: 'Checklist' },
      { col: 5, text: 'MANDAYS (Days)' },
      { col: 6, text: 'TOTAL' },
      { col: 7, text: 'DATELINE' },
      { col: 9, text: 'Deliverables' },
      { col: 10, text: 'Notes' },
      { col: 11, text: 'Commits' },
      { col: 12, text: 'Lines (+/-)' },
      { col: 13, text: 'Check' },
    ];

    topHeaderLabels.forEach((hdr) => {
      const cell = row1.getCell(hdr.col);
      cell.value = hdr.text;
    });

    for (let col = 1; col <= leftColCount; col++) {
      const cell = row1.getCell(col);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: this.colors.navyHeader } };
      cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: this.colors.white } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this._applyBorder(cell);
    }

    // Row 2: Sub-headers for G2 ('Start') & H2 ('End')
    const row2 = worksheet.getRow(2);
    row2.height = 20;
    row2.getCell(7).value = 'Start';
    row2.getCell(8).value = 'End';

    for (let col = 1; col <= leftColCount; col++) {
      const cell = row2.getCell(col);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: this.colors.navyHeader } };
      cell.font = { name: 'Arial', size: 9, bold: true, color: { argb: this.colors.white } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this._applyBorder(cell);
    }

    // Timeline Header (Row 1 Month Banners & Row 2 Day Numbers)
    monthGroups.forEach((monthGroup) => {
      const startCol = leftColCount + monthGroup.startIndex + 1;
      const endCol = leftColCount + monthGroup.endIndex + 1;

      if (startCol < endCol) {
        worksheet.mergeCells(1, startCol, 1, endCol);
      }

      const mCell = row1.getCell(startCol);
      mCell.value = monthGroup.monthName;
      mCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: monthGroup.bannerColor } };
      mCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: this.colors.white } };
      mCell.alignment = { vertical: 'middle', horizontal: 'center' };

      for (let col = startCol; col <= endCol; col++) {
        this._applyBorder(row1.getCell(col));
      }
    });

    dates.forEach((dObj, idx) => {
      const colIdx = leftColCount + idx + 1;
      const dCell = row2.getCell(colIdx);
      dCell.value = dObj.date.getDate();
      dCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: dObj.isWeekend ? this.colors.redWeekend : dObj.monthBannerColor },
      };
      dCell.font = { name: 'Arial', size: 8, bold: true, color: { argb: this.colors.white } };
      dCell.alignment = { vertical: 'middle', horizontal: 'center' };
      this._applyBorder(dCell);
    });

    // -------------------------------------------------------------
    // RENDER SECTIONS & TASKS
    // -------------------------------------------------------------

    let currentRowIndex = 3;

    for (const section of this.sections) {
      // Render Section Header Banner Row
      const sRow = worksheet.getRow(currentRowIndex);
      sRow.height = 24;

      // Merge Section Header across Left Table
      worksheet.mergeCells(currentRowIndex, 1, currentRowIndex, leftColCount);
      const bannerCell = sRow.getCell(1);
      bannerCell.value = section.title;
      bannerCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: section.bannerColor },
      };
      bannerCell.font = { name: 'Arial', size: 11, bold: true, color: { argb: this.colors.white } };
      bannerCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

      for (let col = 1; col <= leftColCount; col++) {
        this._applyBorder(sRow.getCell(col));
      }

      // Render Section Header Timeline Cells
      dates.forEach((dObj, idx) => {
        const colIdx = leftColCount + idx + 1;
        const cell = sRow.getCell(colIdx);
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: dObj.isWeekend ? this.colors.redWeekend : section.bannerColor },
        };
        this._applyBorder(cell);
      });

      currentRowIndex++;

      // Render Task Rows
      for (const task of section.tasks) {
        const tRow = worksheet.getRow(currentRowIndex);
        tRow.height = 22;

        const cellValues = [
          task.no,
          task.name,
          task.detail,
          task.checklist,
          task.mandays,
          task.total,
          this._formatDateDisplay(task.startDate),
          this._formatDateDisplay(task.endDate),
          task.deliverables,
          task.notes,
          task.commitCount,
          task.lines,
          task.check ? '☑' : '☐',
        ];

        cellValues.forEach((val, i) => {
          const colIdx = i + 1;
          const cell = tRow.getCell(colIdx);

          this._formatCellValue(cell, val);

          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: section.rowColor } };
          cell.font = cell.font || {
            name: 'Arial',
            size: 9,
            color: { argb: this.colors.darkText },
          };

          // Alignment specifics
          if (
            colIdx === 1 ||
            colIdx === 4 ||
            colIdx === 5 ||
            colIdx === 6 ||
            colIdx === 7 ||
            colIdx === 8 ||
            colIdx === 11 ||
            colIdx === 12 ||
            colIdx === 13
          ) {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
          } else {
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
          }

          this._applyBorder(cell);
        });

        // Calculate Task Date Range for Gantt Plotting
        const taskStart = this._resetTime(task.startDate);
        const taskEnd = this._resetTime(task.endDate);

        // Find Start Index and End Index in dates array
        let startIdx = -1;
        let endIdx = -1;

        dates.forEach((dObj, dIdx) => {
          const curTime = dObj.date.getTime();
          if (curTime >= taskStart.getTime() && startIdx === -1) {
            startIdx = dIdx;
          }
          if (curTime <= taskEnd.getTime()) {
            endIdx = dIdx;
          }
        });

        if (startIdx !== -1 && endIdx === -1) endIdx = startIdx;

        // Render Gantt Chart Cells for this Task
        dates.forEach((dObj, dIdx) => {
          const colIdx = leftColCount + dIdx + 1;
          const gCell = tRow.getCell(colIdx);

          const isWithinRange = dIdx >= startIdx && dIdx <= endIdx && startIdx !== -1;

          if (isWithinRange) {
            // Render Gantt bar cell (special color for weekend commit if on a weekend)
            const ganttColor = dObj.isWeekend ? 'FF1E5C9B' : this.colors.blueGantt;
            gCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ganttColor } };

            // Show Commit Count badge inside the first Gantt cell
            if (dIdx === startIdx) {
              gCell.value = task.commitCount ? `${task.commitCount}` : '';
              gCell.font = {
                name: 'Arial',
                size: 8,
                bold: true,
                color: { argb: this.colors.white },
              };
              gCell.alignment = { vertical: 'middle', horizontal: 'center' };
            }

            // Rich Tooltip Note on Gantt Bar Cell
            gCell.note = `Task: ${task.name}\nCommits: ${task.commitCount}\nLines (+/-): ${task.lines}\nDates: ${this._formatDateDisplay(task.startDate)} - ${this._formatDateDisplay(task.endDate)}${dObj.isWeekend ? ' (Weekend Commit)' : ''}`;
          } else if (dObj.isWeekend) {
            gCell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: this.colors.redWeekend },
            };
          } else {
            gCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: section.rowColor } };
          }

          this._applyBorder(gCell);
        });

        currentRowIndex++;
      }
    }

    // Write file
    await this.workbook.xlsx.writeFile(filePath);
    return filePath;
  }

  // -------------------------------------------------------------
  // HELPER METHODS
  // -------------------------------------------------------------

  _formatCellValue(cell, value) {
    if (value && typeof value === 'object' && value.hyperlink) {
      cell.value = {
        text: value.text || value.hyperlink,
        hyperlink: value.hyperlink,
      };
      cell.font = {
        name: 'Arial',
        size: 9,
        color: { argb: 'FF0563C1' },
        underline: true,
      };
    } else {
      cell.value = value;
    }
  }

  _computeCalendarTimeline() {
    let minDate = null;
    let maxDate = null;

    if (this.customCalendarRange) {
      minDate = this.customCalendarRange.start;
      maxDate = this.customCalendarRange.end;
    } else {
      this.sections.forEach((sec) => {
        sec.tasks.forEach((t) => {
          if (!minDate || t.startDate < minDate) minDate = t.startDate;
          if (!maxDate || t.endDate > maxDate) maxDate = t.endDate;
        });
      });
    }

    if (!minDate) minDate = new Date();
    if (!maxDate) maxDate = new Date();

    minDate = this._resetTime(minDate);
    maxDate = this._resetTime(maxDate);

    // Expand to 1st of minDate month and Last Day of maxDate month if expandFullMonths is true
    const startDate = new Date(minDate.getFullYear(), minDate.getMonth(), 1);
    let endDate = new Date(maxDate.getFullYear(), maxDate.getMonth() + 1, 0);

    // Enforce minimum months span
    if (this.minMonthsSpan > 1) {
      const monthDiff =
        (endDate.getFullYear() - startDate.getFullYear()) * 12 +
        (endDate.getMonth() - startDate.getMonth()) +
        1;
      if (monthDiff < this.minMonthsSpan) {
        const monthsToAdd = this.minMonthsSpan - monthDiff;
        endDate = new Date(endDate.getFullYear(), endDate.getMonth() + monthsToAdd + 1, 0);
      }
    }

    const dates = [];
    const monthMap = new Map();

    const curr = new Date(startDate.getTime());
    let index = 0;

    const bannerColors = ['1B365D', '2D7D46', '1B365D', '2D7D46'];

    while (curr <= endDate) {
      const isWeekend = curr.getDay() === 0 || curr.getDay() === 6;
      const mKey = `${curr.getFullYear()}-${curr.getMonth() + 1}`;
      const mName = curr
        .toLocaleString('default', { month: 'long', year: 'numeric' })
        .toUpperCase();

      if (!monthMap.has(mKey)) {
        const colorIdx = monthMap.size % bannerColors.length;
        monthMap.set(mKey, {
          monthName: mName,
          bannerColor: bannerColors[colorIdx],
          startIndex: index,
          endIndex: index,
        });
      } else {
        monthMap.get(mKey).endIndex = index;
      }

      const mColor = monthMap.get(mKey).bannerColor;

      dates.push({
        date: new Date(curr.getTime()),
        isWeekend,
        monthKey: mKey,
        monthBannerColor: mColor,
      });

      curr.setDate(curr.getDate() + 1);
      index++;
    }

    return {
      dates,
      monthGroups: Array.from(monthMap.values()),
    };
  }

  _parseDate(dateVal) {
    return parseDate(dateVal);
  }

  _resetTime(date) {
    return resetTime(date);
  }

  _formatDateDisplay(date) {
    return formatDateDisplay(date);
  }

  _resolveColor(colorInput, defaultHex) {
    return resolveColor(colorInput, defaultHex, this.rowPresets);
  }

  _applyBorder(cell) {
    applyXlsxBorder(cell, this.colors.gridBorder);
  }
}

module.exports = ExcelTrackerSDK;
