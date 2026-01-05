// ═══════════════════════════════════════════════════════════════════════════
// REPORTER ENGINE - Generate comprehensive reports
// ═══════════════════════════════════════════════════════════════════════════

const ExcelJS = require('exceljs');
const { FORMATTING, SEVERITY } = require('../utils/constants');
const helpers = require('../utils/helpers');

class Reporter {
  constructor() {
    this.workbook = null;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // MAIN REPORT GENERATION
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Generate complete analysis report workbook
   * @param {Object} params - All data needed for report
   * @returns {ExcelJS.Workbook}
   */
  async generateReport(params) {
    const {
      originalData,
      cleanedData,
      analysisResult,
      cleaningResult,
      options = {},
    } = params;

    this.workbook = new ExcelJS.Workbook();
    this.workbook.creator = 'Excel Intelligent Bot';
    this.workbook.created = new Date();

    // ─────────────────────────────────────────────────────────────────────
    // Sheet 1: Executive Summary
    // ─────────────────────────────────────────────────────────────────────
    await this._createSummarySheet(params);

    // ─────────────────────────────────────────────────────────────────────
    // Sheet 2: Cleaned Data
    // ─────────────────────────────────────────────────────────────────────
    await this._createDataSheet(cleanedData, 'Cleaned Data', '00B050');

    // ─────────────────────────────────────────────────────────────────────
    // Sheet 3: Issues Found
    // ─────────────────────────────────────────────────────────────────────
    if (analysisResult?.issues?.length > 0) {
      await this._createIssuesSheet(analysisResult);
    }

    // ─────────────────────────────────────────────────────────────────────
    // Sheet 4: Changes Log
    // ─────────────────────────────────────────────────────────────────────
    if (cleaningResult?.changes?.length > 0) {
      await this._createChangesSheet(cleaningResult);
    }

    // ─────────────────────────────────────────────────────────────────────
    // Sheet 5: Original Data (Backup)
    // ─────────────────────────────────────────────────────────────────────
    if (options.includeOriginal !== false) {
      await this._createDataSheet(originalData, 'Original Data', 'FFC000');
    }

    // ─────────────────────────────────────────────────────────────────────
    // Sheet 6: Statistics
    // ─────────────────────────────────────────────────────────────────────
    if (options.includeStats !== false) {
      await this._createStatsSheet(cleanedData, analysisResult);
    }

    return this.workbook;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // SUMMARY SHEET
  // ─────────────────────────────────────────────────────────────────────────

  async _createSummarySheet(params) {
    const { originalData, cleanedData, analysisResult, cleaningResult, options } = params;
    
    const ws = this.workbook.addWorksheet('Summary', {
      properties: { tabColor: { argb: '2B579A' } },
    });

    let row = 1;

    // ─────────────────────────────────────────────────────────────────────
    // Title Section
    // ─────────────────────────────────────────────────────────────────────
    ws.mergeCells(`A${row}:F${row}`);
    const titleCell = ws.getCell(`A${row}`);
    titleCell.value = '📊 INTELLIGENT ANALYSIS REPORT';
    titleCell.font = { size: 22, bold: true, color: { argb: 'FF2B579A' } };
    titleCell.alignment = { horizontal: 'center' };
    row += 2;

    // Generated info
    ws.getCell(`A${row}`).value = 'Generated:';
    ws.getCell(`B${row}`).value = new Date().toLocaleString('id-ID', { 
      dateStyle: 'full', 
      timeStyle: 'medium' 
    });
    row++;

    ws.getCell(`A${row}`).value = 'Source File:';
    ws.getCell(`B${row}`).value = originalData.fileName || 'Unknown';
    row += 2;

    // ─────────────────────────────────────────────────────────────────────
    // Quality Score Section
    // ─────────────────────────────────────────────────────────────────────
    ws.mergeCells(`A${row}:F${row}`);
    ws.getCell(`A${row}`).value = '📈 DATA QUALITY SCORE';
    ws.getCell(`A${row}`).font = { size: 16, bold: true };
    ws.getCell(`A${row}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE2EFDA' },
    };
    row += 2;

    const beforeScore = analysisResult?.qualityScore?.score || 0;
    const afterScore = this._calculateAfterScore(beforeScore, cleaningResult);

    // Before score
    ws.getCell(`A${row}`).value = 'Before Analysis:';
    ws.getCell(`B${row}`).value = `${beforeScore}%`;
    ws.getCell(`C${row}`).value = this._getScoreBar(beforeScore);
    ws.getCell(`D${row}`).value = this._getScoreLabel(beforeScore);
    this._styleScoreCell(ws.getCell(`B${row}`), beforeScore);
    row++;

    // After score
    ws.getCell(`A${row}`).value = 'After Cleaning:';
    ws.getCell(`B${row}`).value = `${afterScore}%`;
    ws.getCell(`C${row}`).value = this._getScoreBar(afterScore);
    ws.getCell(`D${row}`).value = this._getScoreLabel(afterScore);
    this._styleScoreCell(ws.getCell(`B${row}`), afterScore);
    row++;

    // Improvement
    const improvement = afterScore - beforeScore;
    ws.getCell(`A${row}`).value = 'Improvement:';
    ws.getCell(`B${row}`).value = `+${improvement}%`;
    ws.getCell(`B${row}`).font = { bold: true, color: { argb: 'FF00B050' } };
    row += 2;

    // ─────────────────────────────────────────────────────────────────────
    // Data Overview Section
    // ─────────────────────────────────────────────────────────────────────
    ws.mergeCells(`A${row}:F${row}`);
    ws.getCell(`A${row}`).value = '📋 DATA OVERVIEW';
    ws.getCell(`A${row}`).font = { size: 16, bold: true };
    ws.getCell(`A${row}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFDCE6F1' },
    };
    row += 2;

    const stats = [
      ['Original Rows', originalData.data?.length || 0],
      ['Cleaned Rows', cleanedData.data?.length || 0],
      ['Rows Removed', (originalData.data?.length || 0) - (cleanedData.data?.length || 0)],
      ['Columns', cleanedData.headers?.length || 0],
    ];

    stats.forEach(([label, value]) => {
      ws.getCell(`A${row}`).value = label + ':';
      ws.getCell(`B${row}`).value = value;
      ws.getCell(`B${row}`).numFmt = '#,##0';
      row++;
    });
    row++;

    // ─────────────────────────────────────────────────────────────────────
    // Issues Summary Section
    // ─────────────────────────────────────────────────────────────────────
    ws.mergeCells(`A${row}:F${row}`);
    ws.getCell(`A${row}`).value = '🔍 ISSUES SUMMARY';
    ws.getCell(`A${row}`).font = { size: 16, bold: true };
    ws.getCell(`A${row}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFF2CC' },
    };
    row += 2;

    const categorized = analysisResult?.categorizedIssues || {};
    
    // Auto-fixed
    ws.getCell(`A${row}`).value = '🟢 Auto-Fixed:';
    ws.getCell(`B${row}`).value = categorized.autoFix?.length || 0;
    ws.getCell(`C${row}`).value = 'issues automatically resolved';
    ws.getCell(`A${row}`).font = { color: { argb: 'FF00B050' } };
    row++;

    // Needs review
    ws.getCell(`A${row}`).value = '🟡 Needs Review:';
    ws.getCell(`B${row}`).value = categorized.needsReview?.length || 0;
    ws.getCell(`C${row}`).value = 'items require your attention';
    ws.getCell(`A${row}`).font = { color: { argb: 'FFFFC000' } };
    row++;

    // Critical
    ws.getCell(`A${row}`).value = '🔴 Critical:';
    ws.getCell(`B${row}`).value = categorized.critical?.length || 0;
    ws.getCell(`C${row}`).value = 'issues need manual fix';
    ws.getCell(`A${row}`).font = { color: { argb: 'FFFF0000' } };
    row += 2;

    // ─────────────────────────────────────────────────────────────────────
    // Changes Summary Section
    // ─────────────────────────────────────────────────────────────────────
    if (cleaningResult?.changesByType) {
      ws.mergeCells(`A${row}:F${row}`);
      ws.getCell(`A${row}`).value = '✅ CHANGES APPLIED';
      ws.getCell(`A${row}`).font = { size: 16, bold: true };
      ws.getCell(`A${row}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE2EFDA' },
      };
      row += 2;

      Object.entries(cleaningResult.changesByType).forEach(([operation, count]) => {
        ws.getCell(`A${row}`).value = `✓ ${operation}:`;
        ws.getCell(`B${row}`).value = count;
        ws.getCell(`C${row}`).value = count === 1 ? 'change' : 'changes';
        row++;
      });
    }

    // ─────────────────────────────────────────────────────────────────────
    // Column widths
    // ─────────────────────────────────────────────────────────────────────
    ws.getColumn(1).width = 25;
    ws.getColumn(2).width = 15;
    ws.getColumn(3).width = 30;
    ws.getColumn(4).width = 15;
    ws.getColumn(5).width = 15;
    ws.getColumn(6).width = 15;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // DATA SHEET
  // ─────────────────────────────────────────────────────────────────────────

  async _createDataSheet(data, sheetName, tabColor) {
    const { headers, data: rows } = data;
    
    const ws = this.workbook.addWorksheet(sheetName, {
      properties: { tabColor: { argb: tabColor } },
      views: [{ state: 'frozen', ySplit: 1 }],
    });

    // Add headers
    ws.addRow(headers);

    // Style header
    const headerRow = ws.getRow(1);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF2B579A' },
      };
      cell.font = {
        bold: true,
        color: { argb: 'FFFFFFFF' },
      };
      cell.alignment = { horizontal: 'center' };
    });
    headerRow.height = 25;

    // Add data rows
    rows.forEach((row, index) => {
      const rowData = headers.map(h => row[h]);
      const addedRow = ws.addRow(rowData);
      
      // Zebra stripes
      if (index % 2 === 0) {
        addedRow.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF5F5F5' },
          };
        });
      }
    });

    // Auto-fit columns
    headers.forEach((header, i) => {
      const col = ws.getColumn(i + 1);
      let maxLen = header.length;
      rows.slice(0, 100).forEach(row => {
        const val = String(row[header] || '');
        if (val.length > maxLen) maxLen = val.length;
      });
      col.width = Math.min(50, Math.max(10, maxLen + 2));
    });

    // Add borders
    const lastRow = rows.length + 1;
    const lastCol = headers.length;
    
    for (let r = 1; r <= lastRow; r++) {
      for (let c = 1; c <= lastCol; c++) {
        ws.getCell(r, c).border = {
          top: { style: 'thin', color: { argb: 'FFD4D4D4' } },
          left: { style: 'thin', color: { argb: 'FFD4D4D4' } },
          bottom: { style: 'thin', color: { argb: 'FFD4D4D4' } },
          right: { style: 'thin', color: { argb: 'FFD4D4D4' } },
        };
      }
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // ISSUES SHEET
  // ─────────────────────────────────────────────────────────────────────────

  async _createIssuesSheet(analysisResult) {
    const ws = this.workbook.addWorksheet('Issues Found', {
      properties: { tabColor: { argb: 'FFFF0000' } },
    });

    const headers = ['Severity', 'Type', 'Location', 'Message', 'Suggestion', 'Status'];
    ws.addRow(headers);

    // Style header
    const headerRow = ws.getRow(1);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF6B6B' },
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    });

    // Add issues
    const issues = analysisResult.issues || [];
    issues.forEach(issue => {
      const location = issue.row 
        ? `Row ${issue.row}${issue.column ? `, Column: ${issue.column}` : ''}`
        : issue.column 
          ? `Column: ${issue.column}` 
          : 'Multiple';

      const status = issue.severity === SEVERITY.AUTO_FIX 
        ? '✅ Fixed' 
        : issue.severity === SEVERITY.NEEDS_REVIEW 
          ? '⚠️ Review' 
          : '❌ Manual';

      const severityIcon = issue.severity === SEVERITY.AUTO_FIX 
        ? '🟢' 
        : issue.severity === SEVERITY.NEEDS_REVIEW 
          ? '🟡' 
          : '🔴';

      const row = ws.addRow([
        `${severityIcon} ${issue.severity}`,
        issue.type,
        location,
        issue.message,
        issue.suggestion || '-',
        status,
      ]);

      // Color code by severity
      const color = issue.severity === SEVERITY.AUTO_FIX 
        ? 'FFE2EFDA'
        : issue.severity === SEVERITY.NEEDS_REVIEW 
          ? 'FFFFF2CC'
          : 'FFFFC7CE';
      
      row.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: color },
        };
      });
    });

    // Column widths
    ws.getColumn(1).width = 15;
    ws.getColumn(2).width = 25;
    ws.getColumn(3).width = 25;
    ws.getColumn(4).width = 50;
    ws.getColumn(5).width = 40;
    ws.getColumn(6).width = 12;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // CHANGES LOG SHEET
  // ─────────────────────────────────────────────────────────────────────────

  async _createChangesSheet(cleaningResult) {
    const ws = this.workbook.addWorksheet('Changes Log', {
      properties: { tabColor: { argb: 'FF00B050' } },
    });

    const headers = ['Operation', 'Count', 'Details'];
    ws.addRow(headers);

    // Style header
    const headerRow = ws.getRow(1);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF00B050' },
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    });

    // Add summary changes
    const summaryChanges = cleaningResult.changes.filter(c => c.type === 'SUMMARY');
    summaryChanges.forEach(change => {
      ws.addRow([
        `✓ ${change.operation}`,
        change.count,
        change.message,
      ]);
    });

    // Column widths
    ws.getColumn(1).width = 30;
    ws.getColumn(2).width = 15;
    ws.getColumn(3).width = 60;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // STATISTICS SHEET
  // ─────────────────────────────────────────────────────────────────────────

  async _createStatsSheet(data, analysisResult) {
    const { headers, data: rows } = data;
    const columnStats = analysisResult?.columnStats || {};
    const columnTypes = analysisResult?.columnTypes || {};

    const ws = this.workbook.addWorksheet('Statistics', {
      properties: { tabColor: { argb: 'FF7030A0' } },
    });

    let currentRow = 1;

    // Title
    ws.mergeCells(`A${currentRow}:E${currentRow}`);
    ws.getCell(`A${currentRow}`).value = '📊 COLUMN STATISTICS';
    ws.getCell(`A${currentRow}`).font = { size: 18, bold: true };
    currentRow += 2;

    // For each column
    headers.forEach(header => {
      const stats = columnStats[header] || {};
      const type = columnTypes[header]?.type || 'unknown';

      // Column header
      ws.getCell(`A${currentRow}`).value = header;
      ws.getCell(`A${currentRow}`).font = { bold: true, size: 14 };
      ws.getCell(`A${currentRow}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFDCE6F1' },
      };
      ws.mergeCells(`A${currentRow}:E${currentRow}`);
      currentRow++;

      // Type
      ws.getCell(`A${currentRow}`).value = 'Data Type:';
      ws.getCell(`B${currentRow}`).value = type;
      currentRow++;

      // Basic stats
      ws.getCell(`A${currentRow}`).value = 'Total Values:';
      ws.getCell(`B${currentRow}`).value = stats.totalCount || 0;
      currentRow++;

      ws.getCell(`A${currentRow}`).value = 'Non-Empty:';
      ws.getCell(`B${currentRow}`).value = stats.nonEmptyCount || 0;
      currentRow++;

      ws.getCell(`A${currentRow}`).value = 'Empty:';
      ws.getCell(`B${currentRow}`).value = stats.emptyCount || 0;
      ws.getCell(`C${currentRow}`).value = `(${stats.emptyPercentage || 0}%)`;
      currentRow++;

      ws.getCell(`A${currentRow}`).value = 'Unique Values:';
      ws.getCell(`B${currentRow}`).value = stats.uniqueCount || 0;
      currentRow++;

      // Numeric stats if applicable
      if (stats.numeric) {
        ws.getCell(`A${currentRow}`).value = 'Sum:';
        ws.getCell(`B${currentRow}`).value = stats.numeric.sum;
        ws.getCell(`B${currentRow}`).numFmt = '#,##0.00';
        currentRow++;

        ws.getCell(`A${currentRow}`).value = 'Average:';
        ws.getCell(`B${currentRow}`).value = stats.numeric.average;
        ws.getCell(`B${currentRow}`).numFmt = '#,##0.00';
        currentRow++;

        ws.getCell(`A${currentRow}`).value = 'Min:';
        ws.getCell(`B${currentRow}`).value = stats.numeric.min;
        currentRow++;

        ws.getCell(`A${currentRow}`).value = 'Max:';
        ws.getCell(`B${currentRow}`).value = stats.numeric.max;
        currentRow++;

        ws.getCell(`A${currentRow}`).value = 'Median:';
        ws.getCell(`B${currentRow}`).value = stats.numeric.median;
        currentRow++;
      }

      // Most common values
      if (stats.mostCommon?.length > 0) {
        ws.getCell(`A${currentRow}`).value = 'Most Common:';
        currentRow++;
        
        stats.mostCommon.slice(0, 3).forEach(({ value, count }) => {
          ws.getCell(`B${currentRow}`).value = value;
          ws.getCell(`C${currentRow}`).value = `(${count}x)`;
          currentRow++;
        });
      }

      currentRow++; // Spacing between columns
    });

    // Column widths
    ws.getColumn(1).width = 20;
    ws.getColumn(2).width = 25;
    ws.getColumn(3).width = 15;
    ws.getColumn(4).width = 15;
    ws.getColumn(5).width = 15;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HELPER METHODS
  // ─────────────────────────────────────────────────────────────────────────

  _calculateAfterScore(beforeScore, cleaningResult) {
    if (!cleaningResult) return beforeScore;
    
    // Improvement based on fixes applied
    const fixes = cleaningResult.stats?.totalChanges || 0;
    const improvement = Math.min(50, fixes * 0.5);
    
    return Math.min(100, Math.round(beforeScore + improvement));
  }

  _getScoreBar(score) {
    const filled = Math.round(score / 5);
    const empty = 20 - filled;
    return '█'.repeat(filled) + '░'.repeat(empty);
  }

  _getScoreLabel(score) {
    if (score >= 90) return '✅ Excellent';
    if (score >= 80) return '✅ Good';
    if (score >= 70) return '⚠️ Fair';
    if (score >= 60) return '⚠️ Poor';
    return '❌ Critical';
  }

  _styleScoreCell(cell, score) {
    cell.font = { 
      bold: true, 
      size: 14,
      color: { 
        argb: score >= 80 ? 'FF00B050' : score >= 60 ? 'FFFFC000' : 'FFFF0000' 
      },
    };
  }

  /**
   * Convert workbook to buffer
   */
  async toBuffer() {
    return await this.workbook.xlsx.writeBuffer();
  }
}

module.exports = new Reporter();
