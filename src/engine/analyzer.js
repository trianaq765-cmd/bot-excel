// ═══════════════════════════════════════════════════════════════════════════
// ANALYZER ENGINE - Intelligent Analysis Core
// ═══════════════════════════════════════════════════════════════════════════

const { ISSUE_TYPES, SEVERITY, PATTERNS, COLUMN_TYPES, INDONESIA } = require('../utils/constants');
const helpers = require('../utils/helpers');

class Analyzer {
  constructor() {
    this.issues = [];
    this.stats = {};
    this.analysisMode = 'auto';
  }

  // ─────────────────────────────────────────────────────────────────────────
  // MAIN ANALYSIS METHOD
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Run full intelligent analysis on data
   * @param {Object} parsedData - Data from FileParser
   * @param {Object} options - Analysis options
   * @returns {Object} Analysis results
   */
  analyze(parsedData, options = {}) {
    const startTime = Date.now();
    this.issues = [];
    this.analysisMode = options.mode || 'auto';

    const { headers, data, columnTypes, columnStats } = parsedData;

    if (!data || data.length === 0) {
      return {
        success: false,
        error: 'No data to analyze',
        issues: [],
        stats: {},
      };
    }

    // ─────────────────────────────────────────────────────────────────────
    // Run all analysis modules
    // ─────────────────────────────────────────────────────────────────────

    // 1. Structure Analysis
    this._analyzeStructure(headers, data);

    // 2. Format Consistency Analysis
    this._analyzeFormatConsistency(headers, data, columnTypes);

    // 3. Data Quality Analysis
    this._analyzeDataQuality(headers, data);

    // 4. Duplicate Analysis
    this._analyzeDuplicates(headers, data);

    // 5. Outlier Analysis
    this._analyzeOutliers(headers, data, columnTypes);

    // 6. Logic & Calculation Analysis
    this._analyzeLogic(headers, data, columnTypes);

    // 7. Indonesia-Specific Validation
    this._analyzeIndonesiaSpecific(headers, data, columnTypes);

    // ─────────────────────────────────────────────────────────────────────
    // Calculate summary statistics
    // ─────────────────────────────────────────────────────────────────────

    const summary = this._calculateSummary(headers, data, parsedData);

    // ─────────────────────────────────────────────────────────────────────
    // Categorize issues by severity
    // ─────────────────────────────────────────────────────────────────────

    const categorizedIssues = {
      autoFix: this.issues.filter(i => i.severity === SEVERITY.AUTO_FIX),
      needsReview: this.issues.filter(i => i.severity === SEVERITY.NEEDS_REVIEW),
      critical: this.issues.filter(i => i.severity === SEVERITY.CRITICAL),
    };

    // Calculate data quality score
    const qualityScore = this._calculateQualityScore(data, this.issues);

    return {
      success: true,
      analysisTime: Date.now() - startTime,
      analysisTimeFormatted: helpers.formatDuration(Date.now() - startTime),
      mode: this.analysisMode,
      
      // Summary
      summary,
      qualityScore,
      
      // Issues
      totalIssues: this.issues.length,
      issues: this.issues,
      categorizedIssues,
      
      // Counts by severity
      autoFixCount: categorizedIssues.autoFix.length,
      needsReviewCount: categorizedIssues.needsReview.length,
      criticalCount: categorizedIssues.critical.length,
      
      // Column info
      columnTypes,
      columnStats,
    };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // STRUCTURE ANALYSIS
  // ─────────────────────────────────────────────────────────────────────────

  _analyzeStructure(headers, data) {
    // Check for empty/duplicate headers
    const headerIssues = [];
    const seen = new Set();

    headers.forEach((header, index) => {
      // Empty header
      if (!header || header.startsWith('Column_')) {
        this._addIssue({
          type: ISSUE_TYPES.NO_HEADER,
          severity: SEVERITY.NEEDS_REVIEW,
          column: index + 1,
          message: `Column ${index + 1} has no proper header`,
          suggestion: 'Add a descriptive header name',
        });
      }

      // Duplicate header
      const lowerHeader = header.toLowerCase();
      if (seen.has(lowerHeader)) {
        this._addIssue({
          type: ISSUE_TYPES.DUPLICATE_HEADER,
          severity: SEVERITY.AUTO_FIX,
          column: header,
          message: `Duplicate header found: "${header}"`,
          suggestion: 'Rename to unique name',
          autoFix: true,
        });
      }
      seen.add(lowerHeader);
    });

    // Check for empty rows in middle of data
    data.forEach((row, index) => {
      const values = headers.map(h => row[h]);
      const nonEmpty = values.filter(v => v !== '' && v !== null && v !== undefined);
      
      if (nonEmpty.length === 0 && index < data.length - 1) {
        this._addIssue({
          type: ISSUE_TYPES.EMPTY_ROW,
          severity: SEVERITY.AUTO_FIX,
          row: row._rowIndex,
          message: `Empty row found at row ${row._rowIndex}`,
          suggestion: 'Remove empty row',
          autoFix: true,
        });
      }
    });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // FORMAT CONSISTENCY ANALYSIS
  // ─────────────────────────────────────────────────────────────────────────

  _analyzeFormatConsistency(headers, data, columnTypes) {
    headers.forEach(header => {
      const type = columnTypes[header]?.type;
      const values = data.map(row => ({ value: row[header], rowIndex: row._rowIndex }))
                        .filter(v => v.value !== '' && v.value !== null);

      if (values.length === 0) return;

      switch (type) {
        case 'date':
          this._checkDateConsistency(header, values);
          break;
        case 'currency':
        case 'number':
          this._checkNumberConsistency(header, values, type);
          break;
        case 'phone':
          this._checkPhoneConsistency(header, values);
          break;
        case 'email':
          this._checkEmailConsistency(header, values);
          break;
        case 'string':
          this._checkTextConsistency(header, values);
          break;
      }
    });
  }

  _checkDateConsistency(header, values) {
    const formats = new Set();
    
    values.forEach(({ value, rowIndex }) => {
      const str = String(value);
      
      // Detect format
      if (/^\d{4}-\d{2}-\d{2}/.test(str)) formats.add('ISO');
      else if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(str)) formats.add('DD/MM/YYYY');
      else if (/^\d{1,2}-\d{1,2}-\d{4}/.test(str)) formats.add('DD-MM-YYYY');
      else if (/^\d{1,2}\s+\w+\s+\d{4}/.test(str)) formats.add('DD MMM YYYY');
      else formats.add('OTHER');

      // Check for future dates
      if (helpers.isFutureDate(value)) {
        this._addIssue({
          type: ISSUE_TYPES.FUTURE_DATE,
          severity: SEVERITY.NEEDS_REVIEW,
          row: rowIndex,
          column: header,
          value: value,
          message: `Future date detected: ${value}`,
          suggestion: 'Verify if this date is correct',
        });
      }
    });

    // Flag inconsistent formats
    if (formats.size > 1) {
      this._addIssue({
        type: ISSUE_TYPES.DATE_INCONSISTENT,
        severity: SEVERITY.AUTO_FIX,
        column: header,
        message: `Inconsistent date formats in column "${header}"`,
        details: `Found formats: ${Array.from(formats).join(', ')}`,
        suggestion: 'Standardize to DD-MMM-YYYY format',
        autoFix: true,
        affectedRows: values.length,
      });
    }
  }

  _checkNumberConsistency(header, values, type) {
    let hasInconsistentFormat = false;
    const issues = [];

    values.forEach(({ value, rowIndex }) => {
      const str = String(value);
      
      // Check for numbers stored as text
      if (typeof value === 'string' && !isNaN(helpers.parseNumber(value))) {
        issues.push({ rowIndex, value, issue: 'number_as_text' });
      }

      // Check currency format
      if (type === 'currency') {
        if (!/^[Rr][Pp]\.?\s?[\d.,]+$/.test(str) && !str.includes('Rp')) {
          issues.push({ rowIndex, value, issue: 'currency_format' });
        }
      }
    });

    if (issues.length > 0) {
      this._addIssue({
        type: type === 'currency' ? ISSUE_TYPES.CURRENCY_FORMAT : ISSUE_TYPES.NUMBER_FORMAT,
        severity: SEVERITY.AUTO_FIX,
        column: header,
        message: `Inconsistent ${type} format in column "${header}"`,
        suggestion: type === 'currency' ? 'Standardize to Rp X,XXX,XXX format' : 'Standardize number format',
        autoFix: true,
        affectedRows: issues.length,
        details: issues.slice(0, 5),
      });
    }
  }

  _checkPhoneConsistency(header, values) {
    const issues = [];

    values.forEach(({ value, rowIndex }) => {
      const str = String(value);
      
      if (!helpers.isValidPhoneID(str)) {
        issues.push({ rowIndex, value, issue: 'invalid_format' });
      } else {
        // Check if format is not standardized
        const formatted = helpers.formatPhoneID(str);
        if (formatted !== str) {
          issues.push({ rowIndex, value, issue: 'needs_formatting', suggested: formatted });
        }
      }
    });

    if (issues.length > 0) {
      this._addIssue({
        type: ISSUE_TYPES.PHONE_FORMAT,
        severity: SEVERITY.AUTO_FIX,
        column: header,
        message: `Phone numbers need formatting in column "${header}"`,
        suggestion: 'Standardize to +62 XXX-XXXX-XXXX format',
        autoFix: true,
        affectedRows: issues.length,
      });
    }
  }

  _checkEmailConsistency(header, values) {
    const issues = [];

    values.forEach(({ value, rowIndex }) => {
      const str = String(value).trim();
      
      // Check valid format
      if (!helpers.isValidEmail(str)) {
        // Check if fixable
        const fixed = helpers.fixEmail(str);
        if (helpers.isValidEmail(fixed)) {
          issues.push({ rowIndex, value, issue: 'fixable', suggested: fixed });
        } else {
          this._addIssue({
            type: ISSUE_TYPES.EMAIL_FORMAT,
            severity: SEVERITY.CRITICAL,
            row: rowIndex,
            column: header,
            value: value,
            message: `Invalid email format: ${value}`,
            suggestion: 'Enter valid email address',
          });
        }
      } else if (str !== str.toLowerCase()) {
        issues.push({ rowIndex, value, issue: 'case', suggested: str.toLowerCase() });
      }
    });

    if (issues.filter(i => i.issue !== 'invalid').length > 0) {
      this._addIssue({
        type: ISSUE_TYPES.EMAIL_FORMAT,
        severity: SEVERITY.AUTO_FIX,
        column: header,
        message: `Email formatting issues in column "${header}"`,
        suggestion: 'Standardize to lowercase and fix minor issues',
        autoFix: true,
        affectedRows: issues.length,
      });
    }
  }

  _checkTextConsistency(header, values) {
    // Check for whitespace issues
    const whitespaceIssues = [];
    const caseIssues = { upper: 0, lower: 0, mixed: 0, title: 0 };

    values.forEach(({ value, rowIndex }) => {
      const str = String(value);
      
      // Check whitespace
      if (str !== str.trim() || /\s{2,}/.test(str)) {
        whitespaceIssues.push({ rowIndex, value });
      }

      // Detect case pattern
      if (str === str.toUpperCase()) caseIssues.upper++;
      else if (str === str.toLowerCase()) caseIssues.lower++;
      else if (str === helpers.toTitleCase(str)) caseIssues.title++;
      else caseIssues.mixed++;
    });

    // Flag whitespace issues
    if (whitespaceIssues.length > 0) {
      this._addIssue({
        type: ISSUE_TYPES.WHITESPACE,
        severity: SEVERITY.AUTO_FIX,
        column: header,
        message: `Whitespace issues in column "${header}"`,
        suggestion: 'Trim and remove extra spaces',
        autoFix: true,
        affectedRows: whitespaceIssues.length,
      });
    }

    // Flag case inconsistency (if not intentionally mixed)
    const total = values.length;
    const dominantCase = Object.entries(caseIssues).sort((a, b) => b[1] - a[1])[0];
    
    if (dominantCase[1] < total * 0.8 && caseIssues.mixed < total * 0.5) {
      this._addIssue({
        type: ISSUE_TYPES.TEXT_CASE,
        severity: SEVERITY.NEEDS_REVIEW,
        column: header,
        message: `Inconsistent text case in column "${header}"`,
        details: caseIssues,
        suggestion: `Consider standardizing to ${dominantCase[0]} case`,
      });
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // DATA QUALITY ANALYSIS
  // ─────────────────────────────────────────────────────────────────────────

  _analyzeDataQuality(headers, data) {
    // Check for missing values in each column
    headers.forEach(header => {
      const emptyCount = data.filter(row => {
        const val = row[header];
        return val === '' || val === null || val === undefined;
      }).length;

      const emptyPercent = (emptyCount / data.length) * 100;

      if (emptyPercent > 0 && emptyPercent < 100) {
        // Some missing values
        if (emptyPercent > 20) {
          this._addIssue({
            type: ISSUE_TYPES.EMPTY_CELL,
            severity: SEVERITY.NEEDS_REVIEW,
            column: header,
            message: `Column "${header}" has ${emptyPercent.toFixed(1)}% empty values`,
            suggestion: 'Review if these should be filled',
            affectedRows: emptyCount,
          });
        }
      }
    });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // DUPLICATE ANALYSIS
  // ─────────────────────────────────────────────────────────────────────────

  _analyzeDuplicates(headers, data) {
    // Exact duplicates
    const seen = new Map();
    const exactDuplicates = [];

    data.forEach((row, index) => {
      // Create key from all values (excluding _rowIndex)
      const key = headers.map(h => String(row[h] || '').toLowerCase().trim()).join('|');
      
      if (seen.has(key)) {
        exactDuplicates.push({
          row: row._rowIndex,
          duplicateOf: seen.get(key),
        });
      } else {
        seen.set(key, row._rowIndex);
      }
    });

    if (exactDuplicates.length > 0) {
      this._addIssue({
        type: ISSUE_TYPES.DUPLICATE_ROW,
        severity: SEVERITY.AUTO_FIX,
        message: `Found ${exactDuplicates.length} exact duplicate rows`,
        suggestion: 'Remove duplicate rows',
        autoFix: true,
        affectedRows: exactDuplicates.length,
        details: exactDuplicates.slice(0, 10),
      });
    }

    // Fuzzy duplicates (similar but not exact)
    this._analyzeFuzzyDuplicates(headers, data);
  }

  _analyzeFuzzyDuplicates(headers, data) {
    // Find potential key columns (name, email, phone, etc.)
    const keyColumns = headers.filter(h => {
      const lower = h.toLowerCase();
      return lower.includes('name') || lower.includes('email') || 
             lower.includes('phone') || lower.includes('id') ||
             lower.includes('nama') || lower.includes('telepon');
    });

    if (keyColumns.length === 0) return;

    const fuzzyDuplicates = [];
    const checked = new Set();

    for (let i = 0; i < data.length; i++) {
      for (let j = i + 1; j < data.length; j++) {
        const key = `${i}-${j}`;
        if (checked.has(key)) continue;
        checked.add(key);

        // Compare key columns
        let similarCount = 0;
        let totalChecked = 0;

        keyColumns.forEach(col => {
          const val1 = String(data[i][col] || '').toLowerCase().trim();
          const val2 = String(data[j][col] || '').toLowerCase().trim();
          
          if (val1 && val2) {
            totalChecked++;
            const similarity = helpers.stringSimilarity(val1, val2);
            if (similarity > 0.85) similarCount++;
          }
        });

        if (totalChecked > 0 && similarCount === totalChecked) {
          fuzzyDuplicates.push({
            row1: data[i]._rowIndex,
            row2: data[j]._rowIndex,
            matchedColumns: keyColumns,
          });
        }
      }

      // Limit check for performance
      if (fuzzyDuplicates.length >= 50) break;
    }

    if (fuzzyDuplicates.length > 0) {
      this._addIssue({
        type: ISSUE_TYPES.DUPLICATE_FUZZY,
        severity: SEVERITY.NEEDS_REVIEW,
        message: `Found ${fuzzyDuplicates.length} potential duplicate entries`,
        suggestion: 'Review these entries - they may be duplicates with slight differences',
        affectedRows: fuzzyDuplicates.length * 2,
        details: fuzzyDuplicates.slice(0, 10),
      });
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // OUTLIER ANALYSIS
  // ─────────────────────────────────────────────────────────────────────────

  _analyzeOutliers(headers, data, columnTypes) {
    headers.forEach(header => {
      const type = columnTypes[header]?.type;
      
      if (!['number', 'currency'].includes(type)) return;

      const values = data.map(row => ({
        value: helpers.parseNumber(row[header]),
        rowIndex: row._rowIndex,
        original: row[header],
      })).filter(v => !isNaN(v.value));

      if (values.length < 10) return; // Need enough data for outlier detection

      const numbers = values.map(v => v.value);
      const outliers = [];

      values.forEach(({ value, rowIndex, original }) => {
        if (helpers.isOutlier(value, numbers)) {
          outliers.push({ rowIndex, value, original });
        }
      });

      if (outliers.length > 0) {
        const stats = helpers.calculateStats(numbers);
        
        outliers.forEach(outlier => {
          this._addIssue({
            type: ISSUE_TYPES.NUMERIC_OUTLIER,
            severity: SEVERITY.NEEDS_REVIEW,
            row: outlier.rowIndex,
            column: header,
            value: outlier.original,
            message: `Outlier detected: ${outlier.original} (average: ${helpers.formatNumber(stats.average)})`,
            suggestion: 'Verify if this value is correct',
            details: {
              value: outlier.value,
              average: stats.average,
              stdDev: stats.stdDev,
              deviation: Math.abs(outlier.value - stats.average) / stats.stdDev,
            },
          });
        });
      }
    });

    // Check for negative values in typically positive columns
    const positiveColumns = headers.filter(h => {
      const lower = h.toLowerCase();
      return lower.includes('qty') || lower.includes('quantity') ||
             lower.includes('jumlah') || lower.includes('stock') ||
             lower.includes('price') || lower.includes('harga') ||
             lower.includes('amount') || lower.includes('total');
    });

    positiveColumns.forEach(header => {
      data.forEach(row => {
        const value = helpers.parseNumber(row[header]);
        if (!isNaN(value) && value < 0) {
          this._addIssue({
            type: ISSUE_TYPES.NEGATIVE_INVALID,
            severity: SEVERITY.NEEDS_REVIEW,
            row: row._rowIndex,
            column: header,
            value: row[header],
            message: `Negative value in typically positive column: ${row[header]}`,
            suggestion: 'Verify if this negative value is correct',
          });
        }
      });
    });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // LOGIC & CALCULATION ANALYSIS
  // ─────────────────────────────────────────────────────────────────────────

  _analyzeLogic(headers, data, columnTypes) {
    // Detect potential calculation columns (Total = Qty × Price)
    const lowerHeaders = headers.map(h => h.toLowerCase());
    
    const hasQty = lowerHeaders.some(h => h.includes('qty') || h.includes('quantity') || h.includes('jumlah'));
    const hasPrice = lowerHeaders.some(h => h.includes('price') || h.includes('harga') || h.includes('unit'));
    const hasTotal = lowerHeaders.some(h => h.includes('total') || h.includes('amount') || h.includes('subtotal'));

    if (hasQty && hasPrice && hasTotal) {
      const qtyCol = headers.find(h => h.toLowerCase().includes('qty') || h.toLowerCase().includes('quantity') || h.toLowerCase().includes('jumlah'));
      const priceCol = headers.find(h => h.toLowerCase().includes('price') || h.toLowerCase().includes('harga') || h.toLowerCase().includes('unit'));
      const totalCol = headers.find(h => h.toLowerCase().includes('total') || h.toLowerCase().includes('amount') || h.toLowerCase().includes('subtotal'));

      if (qtyCol && priceCol && totalCol) {
        const calcErrors = [];

        data.forEach(row => {
          const qty = helpers.parseNumber(row[qtyCol]);
          const price = helpers.parseNumber(row[priceCol]);
          const total = helpers.parseNumber(row[totalCol]);

          if (!isNaN(qty) && !isNaN(price) && !isNaN(total)) {
            const expected = qty * price;
            // Allow 1% tolerance for rounding
            if (Math.abs(expected - total) > Math.max(1, expected * 0.01)) {
              calcErrors.push({
                rowIndex: row._rowIndex,
                qty, price, total, expected,
                difference: total - expected,
              });
            }
          }
        });

        if (calcErrors.length > 0) {
          this._addIssue({
            type: ISSUE_TYPES.CALCULATION_ERROR,
            severity: SEVERITY.AUTO_FIX,
            message: `Found ${calcErrors.length} calculation errors (${qtyCol} × ${priceCol} ≠ ${totalCol})`,
            suggestion: 'Recalculate totals automatically',
            autoFix: true,
            affectedRows: calcErrors.length,
            details: calcErrors.slice(0, 5),
            fixInfo: { qtyCol, priceCol, totalCol },
          });
        }
      }
    }

    // Check for sequence gaps (invoice numbers, etc.)
    const seqColumns = headers.filter(h => {
      const lower = h.toLowerCase();
      return lower.includes('no') || lower.includes('number') || 
             lower.includes('invoice') || lower.includes('id');
    });

    seqColumns.forEach(header => {
      this._checkSequence(header, data);
    });
  }

  _checkSequence(header, data) {
    const values = data.map(row => row[header]).filter(v => v !== '' && v !== null);
    
    // Extract numeric parts
    const numericParts = values.map(v => {
      const match = String(v).match(/\d+/g);
      return match ? parseInt(match[match.length - 1]) : null;
    }).filter(n => n !== null);

    if (numericParts.length < 3) return;

    // Sort and find gaps
    const sorted = [...new Set(numericParts)].sort((a, b) => a - b);
    const gaps = [];

    for (let i = 1; i < sorted.length; i++) {
      const diff = sorted[i] - sorted[i - 1];
      if (diff > 1) {
        // Found gap
        for (let j = sorted[i - 1] + 1; j < sorted[i]; j++) {
          gaps.push(j);
        }
      }
    }

    if (gaps.length > 0 && gaps.length <= 10) {
      this._addIssue({
        type: ISSUE_TYPES.SEQUENCE_GAP,
        severity: SEVERITY.NEEDS_REVIEW,
        column: header,
        message: `Sequence gaps detected in column "${header}"`,
        suggestion: 'Check if missing entries: ' + gaps.slice(0, 5).join(', ') + (gaps.length > 5 ? '...' : ''),
        details: { missingNumbers: gaps },
      });
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // INDONESIA-SPECIFIC VALIDATION
  // ─────────────────────────────────────────────────────────────────────────

  _analyzeIndonesiaSpecific(headers, data, columnTypes) {
    headers.forEach(header => {
      const type = columnTypes[header]?.type;
      const lower = header.toLowerCase();

      // NIK Validation
      if (type === 'nik' || lower.includes('nik') || lower.includes('ktp')) {
        this._validateNIKColumn(header, data);
      }

      // NPWP Validation
      if (type === 'npwp' || lower.includes('npwp')) {
        this._validateNPWPColumn(header, data);
      }
    });

    // Tax calculation validation
    this._validateTaxCalculations(headers, data);
  }

  _validateNIKColumn(header, data) {
    const issues = { invalid: [], needsFormat: [] };

    data.forEach(row => {
      const value = row[header];
      if (!value || value === '') return;

      const validation = helpers.validateNIK(value);
      
      if (!validation.isValid) {
        issues.invalid.push({
          rowIndex: row._rowIndex,
          value,
          errors: validation.errors,
        });
      } else {
        // Check if needs formatting
        const formatted = helpers.formatNIK(value);
        if (formatted !== value && !value.includes('.')) {
          issues.needsFormat.push({ rowIndex: row._rowIndex, value, formatted });
        }
      }
    });

    if (issues.invalid.length > 0) {
      this._addIssue({
        type: ISSUE_TYPES.NIK_INVALID,
        severity: SEVERITY.CRITICAL,
        column: header,
        message: `Found ${issues.invalid.length} invalid NIK numbers`,
        suggestion: 'Verify and correct these NIK numbers',
        affectedRows: issues.invalid.length,
        details: issues.invalid.slice(0, 5),
      });
    }
  }

  _validateNPWPColumn(header, data) {
    const issues = { invalid: [], needsFormat: [] };

    data.forEach(row => {
      const value = row[header];
      if (!value || value === '') return;

      const validation = helpers.validateNPWP(value);
      
      if (!validation.isValid) {
        issues.invalid.push({
          rowIndex: row._rowIndex,
          value,
          errors: validation.errors,
        });
      } else if (validation.formatted !== value) {
        issues.needsFormat.push({ rowIndex: row._rowIndex, value, formatted: validation.formatted });
      }
    });

    if (issues.invalid.length > 0) {
      this._addIssue({
        type: ISSUE_TYPES.NPWP_INVALID,
        severity: SEVERITY.CRITICAL,
        column: header,
        message: `Found ${issues.invalid.length} invalid NPWP numbers`,
        suggestion: 'Verify and correct these NPWP numbers',
        affectedRows: issues.invalid.length,
        details: issues.invalid.slice(0, 5),
      });
    }

    if (issues.needsFormat.length > 0) {
      this._addIssue({
        type: ISSUE_TYPES.NPWP_INVALID,
        severity: SEVERITY.AUTO_FIX,
        column: header,
        message: `Found ${issues.needsFormat.length} NPWP numbers needing format standardization`,
        suggestion: 'Format to XX.XXX.XXX.X-XXX.XXX',
        autoFix: true,
        affectedRows: issues.needsFormat.length,
      });
    }
  }

  _validateTaxCalculations(headers, data) {
    const lowerHeaders = headers.map(h => h.toLowerCase());
    
    // Check PPN (11%)
    const hasDPP = lowerHeaders.some(h => h.includes('dpp') || h.includes('dasar'));
    const hasPPN = lowerHeaders.some(h => h.includes('ppn') || h.includes('vat') || h.includes('tax'));

    if (hasDPP && hasPPN) {
      const dppCol = headers.find(h => h.toLowerCase().includes('dpp') || h.toLowerCase().includes('dasar'));
      const ppnCol = headers.find(h => h.toLowerCase().includes('ppn') || h.toLowerCase().includes('vat') || h.toLowerCase().includes('tax'));

      if (dppCol && ppnCol) {
        const taxErrors = [];

        data.forEach(row => {
          const dpp = helpers.parseNumber(row[dppCol]);
          const ppn = helpers.parseNumber(row[ppnCol]);

          if (!isNaN(dpp) && !isNaN(ppn) && dpp > 0) {
            const expected = dpp * INDONESIA.PPN_RATE;
            // Allow 1 rupiah tolerance
            if (Math.abs(expected - ppn) > 1) {
              taxErrors.push({
                rowIndex: row._rowIndex,
                dpp, ppn, expected,
                difference: ppn - expected,
              });
            }
          }
        });

        if (taxErrors.length > 0) {
          this._addIssue({
            type: ISSUE_TYPES.TAX_CALCULATION,
            severity: SEVERITY.AUTO_FIX,
            message: `Found ${taxErrors.length} PPN calculation errors (should be 11% of DPP)`,
            suggestion: 'Recalculate PPN = DPP × 11%',
            autoFix: true,
            affectedRows: taxErrors.length,
            details: taxErrors.slice(0, 5),
            fixInfo: { dppCol, ppnCol, rate: INDONESIA.PPN_RATE },
          });
        }
      }
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HELPER METHODS
  // ─────────────────────────────────────────────────────────────────────────

  _addIssue(issue) {
    issue.id = helpers.generateId();
    issue.timestamp = new Date().toISOString();
    this.issues.push(issue);
  }

  _calculateSummary(headers, data, parsedData) {
    return {
      totalRows: data.length,
      totalColumns: headers.length,
      totalCells: data.length * headers.length,
      
      // Empty stats
      emptyRows: data.filter(row => 
        headers.every(h => row[h] === '' || row[h] === null || row[h] === undefined)
      ).length,
      
      emptyCells: headers.reduce((sum, h) => 
        sum + data.filter(row => row[h] === '' || row[h] === null || row[h] === undefined).length
      , 0),
      
      // File info
      fileName: parsedData.fileName,
      fileSize: parsedData.fileSizeFormatted,
      sheetCount: parsedData.sheetCount,
    };
  }

  _calculateQualityScore(data, issues) {
    // Base score
    let score = 100;

    // Deduct points based on issues
    issues.forEach(issue => {
      const affectedRows = issue.affectedRows || 1;
      const impactPercent = (affectedRows / data.length) * 100;

      switch (issue.severity) {
        case SEVERITY.CRITICAL:
          score -= Math.min(20, impactPercent * 2);
          break;
        case SEVERITY.NEEDS_REVIEW:
          score -= Math.min(10, impactPercent);
          break;
        case SEVERITY.AUTO_FIX:
          score -= Math.min(5, impactPercent * 0.5);
          break;
      }
    });

    return {
      score: Math.max(0, Math.round(score)),
      grade: score >= 90 ? 'A' : score >= 80 ? 'B' : score >= 70 ? 'C' : score >= 60 ? 'D' : 'F',
      label: score >= 90 ? 'Excellent' : score >= 80 ? 'Good' : score >= 70 ? 'Fair' : score >= 60 ? 'Poor' : 'Critical',
    };
  }
}

module.exports = new Analyzer();
