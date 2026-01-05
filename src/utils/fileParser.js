// ═══════════════════════════════════════════════════════════════════════════
// FILE PARSER - Parse XLSX, XLS, CSV files
// ═══════════════════════════════════════════════════════════════════════════

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const { FILE, COLUMN_TYPES, PATTERNS } = require('./constants');
const helpers = require('./helpers');

class FileParser {
  constructor() {
    this.supportedExtensions = FILE.ALLOWED_EXTENSIONS;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // MAIN PARSE METHOD
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Parse file and return structured data
   * @param {string|Buffer} input - File path or buffer
   * @param {Object} options - Parse options
   * @returns {Object} Parsed data with metadata
   */
  parse(input, options = {}) {
    const startTime = Date.now();
    
    try {
      let workbook;
      let fileName = options.fileName || 'unknown';
      let fileSize = 0;

      // Handle buffer or file path
      if (Buffer.isBuffer(input)) {
        workbook = XLSX.read(input, { type: 'buffer', cellDates: true });
        fileSize = input.length;
      } else if (typeof input === 'string') {
        if (!fs.existsSync(input)) {
          throw new Error(`File not found: ${input}`);
        }
        const stats = fs.statSync(input);
        fileSize = stats.size;
        fileName = path.basename(input);
        workbook = XLSX.readFile(input, { cellDates: true });
      } else {
        throw new Error('Invalid input: expected file path or buffer');
      }

      // Get sheet names
      const sheetNames = workbook.SheetNames;
      if (sheetNames.length === 0) {
        throw new Error('No sheets found in workbook');
      }

      // Parse all sheets
      const sheets = {};
      let totalRows = 0;
      let totalCols = 0;

      for (const sheetName of sheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const parsed = this._parseSheet(sheet, options);
        sheets[sheetName] = parsed;
        totalRows += parsed.rowCount;
        totalCols = Math.max(totalCols, parsed.columnCount);
      }

      // Use first sheet as primary
      const primarySheet = sheets[sheetNames[0]];

      const result = {
        success: true,
        fileName,
        fileSize,
        fileSizeFormatted: helpers.formatBytes(fileSize),
        parseTime: Date.now() - startTime,
        parseTimeFormatted: helpers.formatDuration(Date.now() - startTime),
        
        // Sheet info
        sheetCount: sheetNames.length,
        sheetNames,
        activeSheet: sheetNames[0],
        
        // Data from primary sheet
        headers: primarySheet.headers,
        data: primarySheet.data,
        rawData: primarySheet.rawData,
        
        // Metadata
        rowCount: primarySheet.rowCount,
        columnCount: primarySheet.columnCount,
        cellCount: primarySheet.rowCount * primarySheet.columnCount,
        
        // Column analysis
        columnTypes: primarySheet.columnTypes,
        columnStats: primarySheet.columnStats,
        
        // All sheets (for multi-sheet files)
        sheets,
      };

      return result;

    } catch (error) {
      return {
        success: false,
        error: error.message,
        fileName: options.fileName || 'unknown',
        parseTime: Date.now() - startTime,
      };
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // PARSE SINGLE SHEET
  // ─────────────────────────────────────────────────────────────────────────

  _parseSheet(sheet, options = {}) {
    // Convert to JSON (array of arrays)
    const rawData = XLSX.utils.sheet_to_json(sheet, { 
      header: 1, 
      defval: '',
      raw: false,
    });

    if (rawData.length === 0) {
      return {
        headers: [],
        data: [],
        rawData: [],
        rowCount: 0,
        columnCount: 0,
        columnTypes: {},
        columnStats: {},
      };
    }

    // Detect header row
    const headerRowIndex = this._detectHeaderRow(rawData);
    const headers = rawData[headerRowIndex] || [];
    
    // Clean headers
    const cleanedHeaders = this._cleanHeaders(headers);
    
    // Extract data rows (after header)
    const dataRows = rawData.slice(headerRowIndex + 1);
    
    // Convert to array of objects
    const data = dataRows.map((row, rowIndex) => {
      const obj = { _rowIndex: headerRowIndex + rowIndex + 2 }; // Excel row number
      cleanedHeaders.forEach((header, colIndex) => {
        obj[header] = row[colIndex] !== undefined ? row[colIndex] : '';
      });
      return obj;
    });

    // Analyze column types
    const columnTypes = this._analyzeColumnTypes(cleanedHeaders, data);
    
    // Calculate column statistics
    const columnStats = this._calculateColumnStats(cleanedHeaders, data, columnTypes);

    return {
      headers: cleanedHeaders,
      originalHeaders: headers,
      data,
      rawData,
      rowCount: data.length,
      columnCount: cleanedHeaders.length,
      headerRowIndex,
      columnTypes,
      columnStats,
    };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HEADER DETECTION
  // ─────────────────────────────────────────────────────────────────────────

  _detectHeaderRow(rawData) {
    // Simple heuristic: first non-empty row with mostly string values
    for (let i = 0; i < Math.min(10, rawData.length); i++) {
      const row = rawData[i];
      if (!row || row.length === 0) continue;
      
      // Count non-empty cells
      const nonEmpty = row.filter(cell => cell !== '' && cell !== null && cell !== undefined);
      if (nonEmpty.length === 0) continue;
      
      // Check if most cells are strings (headers are usually strings)
      const stringCells = nonEmpty.filter(cell => 
        typeof cell === 'string' && isNaN(parseFloat(cell))
      );
      
      if (stringCells.length >= nonEmpty.length * 0.5) {
        return i;
      }
    }
    
    return 0; // Default to first row
  }

  _cleanHeaders(headers) {
    const cleaned = [];
    const seen = {};

    headers.forEach((header, index) => {
      let name = String(header || `Column_${index + 1}`).trim();
      
      // Remove special characters
      name = name.replace(/[^\w\s]/g, '_').replace(/\s+/g, '_');
      
      // Handle empty headers
      if (!name || name === '_') {
        name = `Column_${index + 1}`;
      }
      
      // Handle duplicates
      if (seen[name]) {
        let counter = 2;
        while (seen[`${name}_${counter}`]) {
          counter++;
        }
        name = `${name}_${counter}`;
      }
      
      seen[name] = true;
      cleaned.push(name);
    });

    return cleaned;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // COLUMN TYPE ANALYSIS
  // ─────────────────────────────────────────────────────────────────────────

  _analyzeColumnTypes(headers, data) {
    const types = {};
    
    headers.forEach(header => {
      const values = data.map(row => row[header]).filter(v => v !== '' && v !== null && v !== undefined);
      
      if (values.length === 0) {
        types[header] = { type: COLUMN_TYPES.EMPTY, confidence: 1 };
        return;
      }

      // Count each type
      const typeCounts = {
        number: 0,
        currency: 0,
        date: 0,
        email: 0,
        phone: 0,
        nik: 0,
        npwp: 0,
        percentage: 0,
        boolean: 0,
        string: 0,
      };

      values.forEach(value => {
        const strValue = String(value).trim();
        
        // Check each type
        if (PATTERNS.EMAIL.test(strValue)) {
          typeCounts.email++;
        } else if (PATTERNS.NIK.test(strValue.replace(/\D/g, ''))) {
          typeCounts.nik++;
        } else if (PATTERNS.NPWP.test(strValue.replace(/\D/g, ''))) {
          typeCounts.npwp++;
        } else if (PATTERNS.CURRENCY_ID.test(strValue) || /^[Rr]p/i.test(strValue)) {
          typeCounts.currency++;
        } else if (PATTERNS.PHONE_ID.test(strValue) || /^(\+62|62|08)\d+/.test(strValue.replace(/\D/g, ''))) {
          typeCounts.phone++;
        } else if (/%$/.test(strValue) || (parseFloat(strValue) >= 0 && parseFloat(strValue) <= 1 && strValue.includes('.'))) {
          typeCounts.percentage++;
        } else if (helpers.parseDate(value) !== null) {
          typeCounts.date++;
        } else if (!isNaN(helpers.parseNumber(strValue))) {
          typeCounts.number++;
        } else if (['true', 'false', 'yes', 'no', 'ya', 'tidak', '1', '0'].includes(strValue.toLowerCase())) {
          typeCounts.boolean++;
        } else {
          typeCounts.string++;
        }
      });

      // Determine primary type
      const total = values.length;
      let maxType = 'string';
      let maxCount = 0;

      Object.entries(typeCounts).forEach(([type, count]) => {
        if (count > maxCount) {
          maxCount = count;
          maxType = type;
        }
      });

      types[header] = {
        type: maxType,
        confidence: maxCount / total,
        distribution: typeCounts,
        sampleSize: total,
      };
    });

    return types;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // COLUMN STATISTICS
  // ─────────────────────────────────────────────────────────────────────────

  _calculateColumnStats(headers, data, columnTypes) {
    const stats = {};

    headers.forEach(header => {
      const values = data.map(row => row[header]);
      const nonEmpty = values.filter(v => v !== '' && v !== null && v !== undefined);
      const type = columnTypes[header]?.type || 'string';

      const columnStat = {
        totalCount: values.length,
        nonEmptyCount: nonEmpty.length,
        emptyCount: values.length - nonEmpty.length,
        emptyPercentage: ((values.length - nonEmpty.length) / values.length * 100).toFixed(1),
        uniqueCount: new Set(nonEmpty.map(v => String(v).toLowerCase().trim())).size,
      };

      // Add numeric stats if applicable
      if (['number', 'currency', 'percentage'].includes(type)) {
        const numbers = nonEmpty.map(v => helpers.parseNumber(v)).filter(n => !isNaN(n));
        if (numbers.length > 0) {
          const numStats = helpers.calculateStats(numbers);
          columnStat.numeric = numStats;
        }
      }

      // Add most common values
      const valueCounts = {};
      nonEmpty.forEach(v => {
        const key = String(v).toLowerCase().trim();
        valueCounts[key] = (valueCounts[key] || 0) + 1;
      });
      
      const sorted = Object.entries(valueCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);
      
      columnStat.mostCommon = sorted.map(([value, count]) => ({ value, count }));

      stats[header] = columnStat;
    });

    return stats;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // UTILITY METHODS
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Validate file extension
   */
  isValidExtension(fileName) {
    const ext = path.extname(fileName).toLowerCase();
    return this.supportedExtensions.includes(ext);
  }

  /**
   * Get file extension
   */
  getExtension(fileName) {
    return path.extname(fileName).toLowerCase();
  }

  /**
   * Convert parsed data back to workbook
   */
  toWorkbook(data, headers) {
    const ws = XLSX.utils.json_to_sheet(data, { header: headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    return wb;
  }

  /**
   * Write workbook to buffer
   */
  toBuffer(workbook, format = 'xlsx') {
    return XLSX.write(workbook, { type: 'buffer', bookType: format });
  }

  /**
   * Write workbook to file
   */
  toFile(workbook, filePath) {
    XLSX.writeFile(workbook, filePath);
  }
}

module.exports = new FileParser();
