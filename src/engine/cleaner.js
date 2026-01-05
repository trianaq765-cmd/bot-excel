// ═══════════════════════════════════════════════════════════════════════════
// CLEANER ENGINE - Auto-fix data issues
// ═══════════════════════════════════════════════════════════════════════════

const { ISSUE_TYPES, SEVERITY, INDONESIA } = require('../utils/constants');
const helpers = require('../utils/helpers');

class Cleaner {
  constructor() {
    this.changes = [];
    this.stats = {
      totalChanges: 0,
      rowsAffected: 0,
      cellsModified: 0,
    };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // MAIN CLEAN METHOD
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Clean data based on analysis results
   * @param {Object} parsedData - Data from FileParser
   * @param {Object} analysisResult - Result from Analyzer
   * @param {Object} options - Cleaning options
   * @returns {Object} Cleaned data with change log
   */
  clean(parsedData, analysisResult, options = {}) {
    const startTime = Date.now();
    this.changes = [];
    this.stats = { totalChanges: 0, rowsAffected: new Set(), cellsModified: 0 };

    // Deep clone data to avoid mutation
    let { headers, data } = JSON.parse(JSON.stringify(parsedData));
    const { columnTypes, categorizedIssues } = analysisResult;

    // Get auto-fixable issues
    const autoFixIssues = categorizedIssues?.autoFix || [];

    // ─────────────────────────────────────────────────────────────────────
    // Apply fixes based on issue types
    // ─────────────────────────────────────────────────────────────────────

    // 1. Remove duplicate rows
    if (this._hasIssueType(autoFixIssues, ISSUE_TYPES.DUPLICATE_ROW)) {
      data = this._removeDuplicates(headers, data);
    }

    // 2. Remove empty rows
    if (this._hasIssueType(autoFixIssues, ISSUE_TYPES.EMPTY_ROW)) {
      data = this._removeEmptyRows(headers, data);
    }

    // 3. Fix whitespace issues
    if (this._hasIssueType(autoFixIssues, ISSUE_TYPES.WHITESPACE)) {
      data = this._fixWhitespace(headers, data);
    }

    // 4. Standardize date formats
    if (this._hasIssueType(autoFixIssues, ISSUE_TYPES.DATE_INCONSISTENT)) {
      data = this._standardizeDates(headers, data, columnTypes, options.dateFormat);
    }

    // 5. Standardize number/currency formats
    if (this._hasIssueType(autoFixIssues, ISSUE_TYPES.NUMBER_FORMAT) ||
        this._hasIssueType(autoFixIssues, ISSUE_TYPES.CURRENCY_FORMAT)) {
      data = this._standardizeNumbers(headers, data, columnTypes);
    }

    // 6. Standardize phone numbers
    if (this._hasIssueType(autoFixIssues, ISSUE_TYPES.PHONE_FORMAT)) {
      data = this._standardizePhones(headers, data, columnTypes);
    }

    // 7. Fix email formatting
    if (this._hasIssueType(autoFixIssues, ISSUE_TYPES.EMAIL_FORMAT)) {
      data = this._standardizeEmails(headers, data, columnTypes);
    }

    // 8. Fix calculation errors
    if (this._hasIssueType(autoFixIssues, ISSUE_TYPES.CALCULATION_ERROR)) {
      const calcIssue = autoFixIssues.find(i => i.type === ISSUE_TYPES.CALCULATION_ERROR);
      if (calcIssue?.fixInfo) {
        data = this._fixCalculations(data, calcIssue.fixInfo);
      }
    }

    // 9. Fix tax calculations
    if (this._hasIssueType(autoFixIssues, ISSUE_TYPES.TAX_CALCULATION)) {
      const taxIssue = autoFixIssues.find(i => i.type === ISSUE_TYPES.TAX_CALCULATION);
      if (taxIssue?.fixInfo) {
        data = this._fixTaxCalculations(data, taxIssue.fixInfo);
      }
    }

    // 10. Format NPWP
    if (this._hasIssueType(autoFixIssues, ISSUE_TYPES.NPWP_INVALID)) {
      data = this._formatNPWP(headers, data, columnTypes);
    }

    // ─────────────────────────────────────────────────────────────────────
    // Optional: Apply text case standardization
    // ─────────────────────────────────────────────────────────────────────
    
    if (options.textCase) {
      data = this._standardizeTextCase(headers, data, columnTypes, options.textCase);
    }

    // ─────────────────────────────────────────────────────────────────────
    // Generate summary
    // ─────────────────────────────────────────────────────────────────────

    return {
      success: true,
      cleanTime: Date.now() - startTime,
      cleanTimeFormatted: helpers.formatDuration(Date.now() - startTime),
      
      // Cleaned data
      headers,
      data,
      
      // Statistics
      stats: {
        totalChanges: this.changes.length,
        rowsAffected: this.stats.rowsAffected.size,
        cellsModified: this.stats.cellsModified,
        originalRowCount: parsedData.data.length,
        cleanedRowCount: data.length,
        rowsRemoved: parsedData.data.length - data.length,
      },
      
      // Change log
      changes: this.changes,
      
      // Summary by type
      changesByType: this._summarizeChangesByType(),
    };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // CLEANING OPERATIONS
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Remove exact duplicate rows
   */
  _removeDuplicates(headers, data) {
    const seen = new Map();
    const uniqueData = [];
    let removed = 0;

    data.forEach((row, index) => {
      const key = headers.map(h => String(row[h] || '').toLowerCase().trim()).join('|');
      
      if (!seen.has(key)) {
        seen.set(key, row._rowIndex);
        uniqueData.push(row);
      } else {
        removed++;
        this._logChange({
          type: 'REMOVE_DUPLICATE',
          row: row._rowIndex,
          message: `Removed duplicate of row ${seen.get(key)}`,
        });
      }
    });

    if (removed > 0) {
      this.changes.push({
        type: 'SUMMARY',
        operation: 'Remove Duplicates',
        count: removed,
        message: `Removed ${removed} duplicate rows`,
      });
    }

    return uniqueData;
  }

  /**
   * Remove empty rows
   */
  _removeEmptyRows(headers, data) {
    const cleanedData = [];
    let removed = 0;

    data.forEach(row => {
      const hasContent = headers.some(h => {
        const val = row[h];
        return val !== '' && val !== null && val !== undefined;
      });

      if (hasContent) {
        cleanedData.push(row);
      } else {
        removed++;
        this._logChange({
          type: 'REMOVE_EMPTY',
          row: row._rowIndex,
          message: `Removed empty row`,
        });
      }
    });

    if (removed > 0) {
      this.changes.push({
        type: 'SUMMARY',
        operation: 'Remove Empty Rows',
        count: removed,
        message: `Removed ${removed} empty rows`,
      });
    }

    return cleanedData;
  }

  /**
   * Fix whitespace issues (trim + remove multiple spaces)
   */
  _fixWhitespace(headers, data) {
    let fixed = 0;

    data.forEach(row => {
      headers.forEach(header => {
        const val = row[header];
        if (typeof val === 'string') {
          const cleaned = helpers.cleanString(val);
          if (cleaned !== val) {
            row[header] = cleaned;
            fixed++;
            this.stats.cellsModified++;
            this.stats.rowsAffected.add(row._rowIndex);
          }
        }
      });
    });

    if (fixed > 0) {
      this.changes.push({
        type: 'SUMMARY',
        operation: 'Fix Whitespace',
        count: fixed,
        message: `Fixed whitespace in ${fixed} cells`,
      });
    }

    return data;
  }

  /**
   * Standardize date formats
   */
  _standardizeDates(headers, data, columnTypes, targetFormat = 'DD-MMM-YYYY') {
    let fixed = 0;

    headers.forEach(header => {
      if (columnTypes[header]?.type === 'date') {
        data.forEach(row => {
          const val = row[header];
          if (val && val !== '') {
            const formatted = helpers.formatDate(val, targetFormat);
            if (formatted !== val) {
              row[header] = formatted;
              fixed++;
              this.stats.cellsModified++;
              this.stats.rowsAffected.add(row._rowIndex);
            }
          }
        });
      }
    });

    if (fixed > 0) {
      this.changes.push({
        type: 'SUMMARY',
        operation: 'Standardize Dates',
        count: fixed,
        message: `Standardized ${fixed} dates to ${targetFormat} format`,
      });
    }

    return data;
  }

  /**
   * Standardize number and currency formats
   */
  _standardizeNumbers(headers, data, columnTypes) {
    let fixed = 0;

    headers.forEach(header => {
      const type = columnTypes[header]?.type;
      
      if (type === 'currency') {
        data.forEach(row => {
          const val = row[header];
          if (val && val !== '') {
            const num = helpers.parseNumber(val);
            if (!isNaN(num)) {
              const formatted = helpers.formatCurrency(num);
              if (formatted !== val) {
                row[header] = formatted;
                fixed++;
                this.stats.cellsModified++;
                this.stats.rowsAffected.add(row._rowIndex);
              }
            }
          }
        });
      } else if (type === 'number') {
        data.forEach(row => {
          const val = row[header];
          if (val && val !== '' && typeof val === 'string') {
            const num = helpers.parseNumber(val);
            if (!isNaN(num)) {
              // Keep as number, formatted with thousand separators
              row[header] = num; // Store as actual number
              fixed++;
              this.stats.cellsModified++;
              this.stats.rowsAffected.add(row._rowIndex);
            }
          }
        });
      }
    });

    if (fixed > 0) {
      this.changes.push({
        type: 'SUMMARY',
        operation: 'Standardize Numbers',
        count: fixed,
        message: `Standardized ${fixed} number/currency values`,
      });
    }

    return data;
  }

  /**
   * Standardize phone numbers
   */
  _standardizePhones(headers, data, columnTypes) {
    let fixed = 0;

    headers.forEach(header => {
      if (columnTypes[header]?.type === 'phone') {
        data.forEach(row => {
          const val = row[header];
          if (val && val !== '') {
            const formatted = helpers.formatPhoneID(val);
            if (formatted !== val && formatted !== String(val)) {
              row[header] = formatted;
              fixed++;
              this.stats.cellsModified++;
              this.stats.rowsAffected.add(row._rowIndex);
            }
          }
        });
      }
    });

    if (fixed > 0) {
      this.changes.push({
        type: 'SUMMARY',
        operation: 'Standardize Phones',
        count: fixed,
        message: `Standardized ${fixed} phone numbers to +62 format`,
      });
    }

    return data;
  }

  /**
   * Standardize email formatting
   */
  _standardizeEmails(headers, data, columnTypes) {
    let fixed = 0;

    headers.forEach(header => {
      if (columnTypes[header]?.type === 'email') {
        data.forEach(row => {
          const val = row[header];
          if (val && val !== '') {
            const cleaned = helpers.fixEmail(val);
            if (cleaned !== val) {
              row[header] = cleaned;
              fixed++;
              this.stats.cellsModified++;
              this.stats.rowsAffected.add(row._rowIndex);
            }
          }
        });
      }
    });

    if (fixed > 0) {
      this.changes.push({
        type: 'SUMMARY',
        operation: 'Standardize Emails',
        count: fixed,
        message: `Fixed ${fixed} email addresses`,
      });
    }

    return data;
  }

  /**
   * Fix calculation errors (Qty × Price = Total)
   */
  _fixCalculations(data, fixInfo) {
    const { qtyCol, priceCol, totalCol } = fixInfo;
    let fixed = 0;

    data.forEach(row => {
      const qty = helpers.parseNumber(row[qtyCol]);
      const price = helpers.parseNumber(row[priceCol]);
      const currentTotal = helpers.parseNumber(row[totalCol]);

      if (!isNaN(qty) && !isNaN(price)) {
        const expectedTotal = qty * price;
        
        // Check if current total is wrong
        if (isNaN(currentTotal) || Math.abs(expectedTotal - currentTotal) > 1) {
          row[totalCol] = expectedTotal;
          fixed++;
          this.stats.cellsModified++;
          this.stats.rowsAffected.add(row._rowIndex);
          
          this._logChange({
            type: 'FIX_CALCULATION',
            row: row._rowIndex,
            column: totalCol,
            oldValue: currentTotal,
            newValue: expectedTotal,
            message: `Fixed: ${qty} × ${price} = ${expectedTotal}`,
          });
        }
      }
    });

    if (fixed > 0) {
      this.changes.push({
        type: 'SUMMARY',
        operation: 'Fix Calculations',
        count: fixed,
        message: `Recalculated ${fixed} total values`,
      });
    }

    return data;
  }

  /**
   * Fix tax calculations (PPN = DPP × 11%)
   */
  _fixTaxCalculations(data, fixInfo) {
    const { dppCol, ppnCol, rate } = fixInfo;
    let fixed = 0;

    data.forEach(row => {
      const dpp = helpers.parseNumber(row[dppCol]);
      const currentPPN = helpers.parseNumber(row[ppnCol]);

      if (!isNaN(dpp) && dpp > 0) {
        const expectedPPN = Math.round(dpp * rate);
        
        if (isNaN(currentPPN) || Math.abs(expectedPPN - currentPPN) > 1) {
          row[ppnCol] = expectedPPN;
          fixed++;
          this.stats.cellsModified++;
          this.stats.rowsAffected.add(row._rowIndex);
          
          this._logChange({
            type: 'FIX_TAX',
            row: row._rowIndex,
            column: ppnCol,
            oldValue: currentPPN,
            newValue: expectedPPN,
            message: `Fixed PPN: ${dpp} × ${rate * 100}% = ${expectedPPN}`,
          });
        }
      }
    });

    if (fixed > 0) {
      this.changes.push({
        type: 'SUMMARY',
        operation: 'Fix Tax Calculations',
        count: fixed,
        message: `Recalculated ${fixed} PPN values (${rate * 100}%)`,
      });
    }

    return data;
  }

  /**
   * Format NPWP numbers
   */
  _formatNPWP(headers, data, columnTypes) {
    let fixed = 0;

    headers.forEach(header => {
      const lower = header.toLowerCase();
      if (columnTypes[header]?.type === 'npwp' || lower.includes('npwp')) {
        data.forEach(row => {
          const val = row[header];
          if (val && val !== '') {
            const validation = helpers.validateNPWP(val);
            if (validation.isValid && validation.formatted !== val) {
              row[header] = validation.formatted;
              fixed++;
              this.stats.cellsModified++;
              this.stats.rowsAffected.add(row._rowIndex);
            }
          }
        });
      }
    });

    if (fixed > 0) {
      this.changes.push({
        type: 'SUMMARY',
        operation: 'Format NPWP',
        count: fixed,
        message: `Formatted ${fixed} NPWP numbers`,
      });
    }

    return data;
  }

  /**
   * Standardize text case
   */
  _standardizeTextCase(headers, data, columnTypes, targetCase) {
    let fixed = 0;

    headers.forEach(header => {
      if (columnTypes[header]?.type === 'string') {
        data.forEach(row => {
          const val = row[header];
          if (val && typeof val === 'string') {
            let newVal;
            
            switch (targetCase) {
              case 'upper':
                newVal = val.toUpperCase();
                break;
              case 'lower':
                newVal = val.toLowerCase();
                break;
              case 'title':
                newVal = helpers.toTitleCase(val);
                break;
              case 'sentence':
                newVal = helpers.toSentenceCase(val);
                break;
              default:
                newVal = val;
            }
            
            if (newVal !== val) {
              row[header] = newVal;
              fixed++;
              this.stats.cellsModified++;
              this.stats.rowsAffected.add(row._rowIndex);
            }
          }
        });
      }
    });

    if (fixed > 0) {
      this.changes.push({
        type: 'SUMMARY',
        operation: 'Standardize Text Case',
        count: fixed,
        message: `Changed ${fixed} cells to ${targetCase} case`,
      });
    }

    return data;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // STANDALONE CLEANING METHODS (Can be called individually)
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Clean data without analysis (basic cleaning)
   */
  basicClean(headers, data, options = {}) {
    this.changes = [];
    this.stats = { totalChanges: 0, rowsAffected: new Set(), cellsModified: 0 };

    let cleanedData = [...data];

    // Remove duplicates
    if (options.removeDuplicates !== false) {
      cleanedData = this._removeDuplicates(headers, cleanedData);
    }

    // Remove empty rows
    if (options.removeEmpty !== false) {
      cleanedData = this._removeEmptyRows(headers, cleanedData);
    }

    // Fix whitespace
    if (options.trimWhitespace !== false) {
      cleanedData = this._fixWhitespace(headers, cleanedData);
    }

    // Text case
    if (options.textCase) {
      // First detect column types
      const fakeColumnTypes = {};
      headers.forEach(h => {
        fakeColumnTypes[h] = { type: 'string' };
      });
      cleanedData = this._standardizeTextCase(headers, cleanedData, fakeColumnTypes, options.textCase);
    }

    return {
      headers,
      data: cleanedData,
      stats: {
        totalChanges: this.changes.length,
        rowsAffected: this.stats.rowsAffected.size,
        cellsModified: this.stats.cellsModified,
        originalRowCount: data.length,
        cleanedRowCount: cleanedData.length,
        rowsRemoved: data.length - cleanedData.length,
      },
      changes: this.changes,
    };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HELPER METHODS
  // ─────────────────────────────────────────────────────────────────────────

  _hasIssueType(issues, type) {
    return issues.some(issue => issue.type === type);
  }

  _logChange(change) {
    change.timestamp = new Date().toISOString();
    this.changes.push(change);
  }

  _summarizeChangesByType() {
    const summary = {};
    
    this.changes
      .filter(c => c.type === 'SUMMARY')
      .forEach(c => {
        summary[c.operation] = c.count;
      });
    
    return summary;
  }
}

module.exports = new Cleaner();
