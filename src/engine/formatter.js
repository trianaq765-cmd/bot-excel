// ═══════════════════════════════════════════════════════════════════════════
// FORMATTER ENGINE - Professional Excel Styling
// ═══════════════════════════════════════════════════════════════════════════

const ExcelJS = require('exceljs');
const { FORMATTING, COLUMN_TYPES } = require('../utils/constants');
const helpers = require('../utils/helpers');

class Formatter {
  constructor() {
    this.workbook = null;
    this.worksheet = null;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // MAIN FORMAT METHOD
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Create professionally formatted Excel workbook
   * @param {Object} data - { headers, data, columnTypes }
   * @param {Object} options - Formatting options
   * @returns {ExcelJS.Workbook} Formatted workbook
   */
  async format(data, options = {}) {
    const { headers, data: rows, columnTypes = {} } = data;
    
    this.workbook = new ExcelJS.Workbook();
    this.workbook.creator = 'Excel Intelligent Bot';
    this.workbook.created = new Date();

    // ─────────────────────────────────────────────────────────────────────
    // Create main data sheet
    // ─────────────────────────────────────────────────────────────────────
    
    this.worksheet = this.workbook.addWorksheet('Data', {
      properties: { tabColor: { argb: '2B579A' } },
      views: [{ state: 'frozen', ySplit: 1 }], // Freeze header row
    });

    // Add headers
    this.worksheet.addRow(headers);

    // Add data rows
    rows.forEach(row => {
      const rowData = headers.map(h => row[h]);
      this.worksheet.addRow(rowData);
    });

    // ─────────────────────────────────────────────────────────────────────
    // Apply formatting
    // ─────────────────────────────────────────────────────────────────────

    // Header styling
    this._formatHeader(headers, options);

    // Column formatting based on data types
    this._formatColumns(headers, columnTypes, options);

    // Row styling (zebra stripes)
    if (options.zebraStripes !== false) {
      this._applyZebraStripes(rows.length);
    }

    // Auto-fit column widths
    this._autoFitColumns(headers, rows);

    // Add borders
    if (options.borders !== false) {
      this._addBorders(headers.length, rows.length);
    }

    // Add total row if requested
    if (options.addTotals) {
      this._addTotalRow(headers, rows, columnTypes);
    }

    // ─────────────────────────────────────────────────────────────────────
    // Create summary sheet if requested
    // ─────────────────────────────────────────────────────────────────────

    if (options.addSummary) {
      this._createSummarySheet(headers, rows, columnTypes, options);
    }

    return this.workbook;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HEADER FORMATTING
  // ─────────────────────────────────────────────────────────────────────────

  _formatHeader(headers, options) {
    const headerRow = this.worksheet.getRow(1);
    
    const headerColor = options.headerColor || FORMATTING.HEADER_COLOR;
    const fontColor = options.headerFontColor || FORMATTING.HEADER_FONT_COLOR;

    headerRow.eachCell((cell, colNumber) => {
      // Fill color
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: headerColor },
      };

      // Font
      cell.font = {
        name: FORMATTING.FONT_NAME,
        size: FORMATTING.HEADER_FONT_SIZE,
        bold: true,
        color: { argb: fontColor },
      };

      // Alignment
      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center',
        wrapText: true,
      };

      // Border
      cell.border = {
        top: { style: 'medium', color: { argb: headerColor } },
        bottom: { style: 'medium', color: { argb: headerColor } },
        left: { style: 'thin', color: { argb: 'FFFFFFFF' } },
        right: { style: 'thin', color: { argb: 'FFFFFFFF' } },
      };
    });

    // Set row height
    headerRow.height = 30;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // COLUMN FORMATTING
  // ─────────────────────────────────────────────────────────────────────────

  _formatColumns(headers, columnTypes, options) {
    headers.forEach((header, index) => {
      const colNumber = index + 1;
      const column = this.worksheet.getColumn(colNumber);
      const type = columnTypes[header]?.type || 'string';

      // Apply number format based on type
      switch (type) {
        case 'currency':
          column.numFmt = '"Rp "#,##0';
          column.alignment = { horizontal: 'right' };
          break;
          
        case 'number':
          column.numFmt = '#,##0';
          column.alignment = { horizontal: 'right' };
          break;
          
        case 'percentage':
          column.numFmt = '0.00%';
          column.alignment = { horizontal: 'right' };
          break;
          
        case 'date':
          column.numFmt = 'DD-MMM-YYYY';
          column.alignment = { horizontal: 'center' };
          break;
          
        case 'email':
        case 'phone':
          column.alignment = { horizontal: 'left' };
          break;
          
        default:
          column.alignment = { horizontal: 'left' };
      }
    });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // ZEBRA STRIPES
  // ─────────────────────────────────────────────────────────────────────────

  _applyZebraStripes(rowCount) {
    for (let i = 2; i <= rowCount + 1; i++) {
      if (i % 2 === 0) {
        const row = this.worksheet.getRow(i);
        row.eachCell((cell) => {
          if (!cell.fill || cell.fill.type !== 'pattern') {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: FORMATTING.ALT_ROW_COLOR },
            };
          }
        });
      }
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // AUTO-FIT COLUMNS
  // ─────────────────────────────────────────────────────────────────────────

  _autoFitColumns(headers, rows) {
    headers.forEach((header, index) => {
      const colNumber = index + 1;
      const column = this.worksheet.getColumn(colNumber);
      
      // Calculate max width
      let maxWidth = header.length;
      
      rows.forEach(row => {
        const value = row[header];
        if (value !== null && value !== undefined) {
          const length = String(value).length;
          if (length > maxWidth) {
            maxWidth = length;
          }
        }
      });

      // Set width with padding (min 10, max 50)
      column.width = Math.min(50, Math.max(10, maxWidth + 2));
    });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // BORDERS
  // ─────────────────────────────────────────────────────────────────────────

  _addBorders(colCount, rowCount) {
    const borderStyle = {
      style: 'thin',
      color: { argb: FORMATTING.BORDER_COLOR },
    };

    for (let row = 2; row <= rowCount + 1; row++) {
      for (let col = 1; col <= colCount; col++) {
        const cell = this.worksheet.getCell(row, col);
        cell.border = {
          top: borderStyle,
          left: borderStyle,
          bottom: borderStyle,
          right: borderStyle,
        };
      }
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // TOTAL ROW
  // ─────────────────────────────────────────────────────────────────────────

  _addTotalRow(headers, rows, columnTypes) {
    const totalRow = ['TOTAL'];
    const rowNumber = rows.length + 2;

    headers.forEach((header, index) => {
      if (index === 0) return; // Skip first column (label)
      
      const type = columnTypes[header]?.type;
      
      if (['number', 'currency'].includes(type)) {
        // Add SUM formula
        const colLetter = this._getColumnLetter(index + 1);
        totalRow.push({ formula: `SUM(${colLetter}2:${colLetter}${rowNumber - 1})` });
      } else {
        totalRow.push('');
      }
    });

    const addedRow = this.worksheet.addRow(totalRow);
    
    // Style total row
    addedRow.eachCell((cell) => {
      cell.font = {
        name: FORMATTING.FONT_NAME,
        size: FORMATTING.FONT_SIZE,
        bold: true,
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE2EFDA' }, // Light green
      };
      cell.border = {
        top: { style: 'double', color: { argb: 'FF000000' } },
        bottom: { style: 'double', color: { argb: 'FF000000' } },
      };
    });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // SUMMARY SHEET
  // ─────────────────────────────────────────────────────────────────────────

  _createSummarySheet(headers, rows, columnTypes, options) {
    const summarySheet = this.workbook.addWorksheet('Summary', {
      properties: { tabColor: { argb: '00B050' } },
    });

    // Title
    summarySheet.mergeCells('A1:D1');
    const titleCell = summarySheet.getCell('A1');
    titleCell.value = options.reportTitle || 'Data Summary Report';
    titleCell.font = { size: 18, bold: true, color: { argb: '2B579A' } };
    titleCell.alignment = { horizontal: 'center' };

    // Generated date
    summarySheet.getCell('A3').value = 'Generated:';
    summarySheet.getCell('B3').value = new Date().toLocaleString('id-ID');

    // Data statistics
    summarySheet.getCell('A5').value = 'Total Rows:';
    summarySheet.getCell('B5').value = rows.length;
    
    summarySheet.getCell('A6').value = 'Total Columns:';
    summarySheet.getCell('B6').value = headers.length;

    // Column statistics for numeric columns
    let currentRow = 8;
    summarySheet.getCell(`A${currentRow}`).value = 'Column Statistics';
    summarySheet.getCell(`A${currentRow}`).font = { bold: true, size: 14 };
    currentRow += 2;

    headers.forEach(header => {
      const type = columnTypes[header]?.type;
      
      if (['number', 'currency'].includes(type)) {
        const values = rows.map(r => helpers.parseNumber(r[header])).filter(v => !isNaN(v));
        
        if (values.length > 0) {
          const stats = helpers.calculateStats(values);
          
          summarySheet.getCell(`A${currentRow}`).value = header;
          summarySheet.getCell(`A${currentRow}`).font = { bold: true };
          currentRow++;
          
          summarySheet.getCell(`A${currentRow}`).value = '  Sum:';
          summarySheet.getCell(`B${currentRow}`).value = stats.sum;
          summarySheet.getCell(`B${currentRow}`).numFmt = type === 'currency' ? '"Rp "#,##0' : '#,##0';
          currentRow++;
          
          summarySheet.getCell(`A${currentRow}`).value = '  Average:';
          summarySheet.getCell(`B${currentRow}`).value = stats.average;
          summarySheet.getCell(`B${currentRow}`).numFmt = type === 'currency' ? '"Rp "#,##0' : '#,##0.00';
          currentRow++;
          
          summarySheet.getCell(`A${currentRow}`).value = '  Min:';
          summarySheet.getCell(`B${currentRow}`).value = stats.min;
          currentRow++;
          
          summarySheet.getCell(`A${currentRow}`).value = '  Max:';
          summarySheet.getCell(`B${currentRow}`).value = stats.max;
          currentRow += 2;
        }
      }
    });

    // Auto-fit columns
    summarySheet.getColumn(1).width = 20;
    summarySheet.getColumn(2).width = 25;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // CUSTOM FORMATTING FROM INSTRUCTIONS
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Apply custom formatting based on user instructions
   */
  async applyCustomFormat(workbook, instructions) {
    const worksheet = workbook.worksheets[0];
    
    instructions.forEach(instruction => {
      switch (instruction.type) {
        case 'headerColor':
          this._setHeaderColor(worksheet, instruction.color);
          break;
          
        case 'columnFormat':
          this._setColumnFormat(worksheet, instruction.column, instruction.format);
          break;
          
        case 'freezePane':
          worksheet.views = [{ state: 'frozen', ySplit: instruction.row || 1 }];
          break;
          
        case 'addTotals':
          // Handled separately
          break;
          
        case 'conditionalFormat':
          this._addConditionalFormatting(worksheet, instruction);
          break;
      }
    });

    return workbook;
  }

  _setHeaderColor(worksheet, color) {
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: color.replace('#', 'FF') },
      };
    });
  }

  _setColumnFormat(worksheet, columnIndex, format) {
    const column = worksheet.getColumn(columnIndex);
    column.numFmt = format;
  }

  _addConditionalFormatting(worksheet, instruction) {
    const { column, rule, style } = instruction;
    
    worksheet.addConditionalFormatting({
      ref: `${column}2:${column}1000`,
      rules: [{
        type: 'expression',
        formulae: [rule],
        style: style,
      }],
    });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // UTILITY METHODS
  // ─────────────────────────────────────────────────────────────────────────

  _getColumnLetter(colNumber) {
    let letter = '';
    while (colNumber > 0) {
      const mod = (colNumber - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      colNumber = Math.floor((colNumber - 1) / 26);
    }
    return letter;
  }

  /**
   * Convert workbook to buffer
   */
  async toBuffer(workbook) {
    return await workbook.xlsx.writeBuffer();
  }

  /**
   * Write workbook to file
   */
  async toFile(workbook, filePath) {
    await workbook.xlsx.writeFile(filePath);
  }
}

module.exports = new Formatter();
