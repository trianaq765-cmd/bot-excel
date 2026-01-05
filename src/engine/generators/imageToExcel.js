// ═══════════════════════════════════════════════════════════════════════════
// IMAGE TO EXCEL - OCR-based table extraction (Optional Feature)
// ═══════════════════════════════════════════════════════════════════════════

// Lazy load tesseract.js - only when needed
let Tesseract = null;

class ImageToExcel {
  constructor() {
    this.worker = null;
    this.initialized = false;
    this.available = false;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // CHECK AVAILABILITY
  // ─────────────────────────────────────────────────────────────────────────

  async checkAvailability() {
    if (Tesseract !== null) {
      return this.available;
    }

    try {
      Tesseract = require('tesseract.js');
      this.available = true;
      console.log('✅ OCR (tesseract.js) is available');
    } catch (error) {
      Tesseract = false;
      this.available = false;
      console.log('⚠️ OCR (tesseract.js) not available - Image extraction disabled');
    }

    return this.available;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // INITIALIZATION
  // ─────────────────────────────────────────────────────────────────────────

  async initialize() {
    if (this.initialized) return true;

    const available = await this.checkAvailability();
    if (!available) {
      return false;
    }

    try {
      this.worker = await Tesseract.createWorker('ind+eng');
      this.initialized = true;
      return true;
    } catch (error) {
      console.error('Failed to initialize OCR:', error.message);
      return false;
    }
  }

  async terminate() {
    if (this.worker) {
      await this.worker.terminate();
      this.worker = null;
      this.initialized = false;
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // MAIN EXTRACTION METHOD
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Extract table data from image
   * @param {Buffer|string} image - Image buffer or path
   * @param {Object} options - Extraction options
   * @returns {Object} Extracted data
   */
  async extract(image, options = {}) {
    const startTime = Date.now();

    // Check if OCR is available
    const available = await this.checkAvailability();
    if (!available) {
      return {
        success: false,
        error: 'OCR feature not available. Please install tesseract.js: npm install tesseract.js',
        extractionTime: Date.now() - startTime,
      };
    }

    try {
      const initialized = await this.initialize();
      if (!initialized) {
        throw new Error('Failed to initialize OCR engine');
      }

      // Perform OCR
      const { data } = await this.worker.recognize(image);

      const result = {
        success: true,
        extractionTime: Date.now() - startTime,
        confidence: data.confidence,
        rawText: data.text,
      };

      // Parse the extracted text into structured data
      const parsed = this._parseExtractedText(data.text, options);
      
      return {
        ...result,
        ...parsed,
      };

    } catch (error) {
      return {
        success: false,
        error: error.message,
        extractionTime: Date.now() - startTime,
      };
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // TEXT PARSING
  // ─────────────────────────────────────────────────────────────────────────

  _parseExtractedText(text, options) {
    const lines = text.split('\n')
      .map(l => l.trim())
      .filter(l => l.length > 0);

    if (lines.length === 0) {
      return { headers: [], data: [], tableDetected: false };
    }

    // Try to detect table structure
    const tableData = this._detectTable(lines);
    
    if (tableData.isTable) {
      return {
        tableDetected: true,
        headers: tableData.headers,
        data: tableData.data,
        warnings: tableData.warnings,
      };
    }

    // Fallback: treat as list
    return {
      tableDetected: false,
      headers: ['No', 'Content'],
      data: lines.map((line, i) => ({
        _rowIndex: i + 2,
        No: i + 1,
        Content: line,
      })),
    };
  }

  _detectTable(lines) {
    const warnings = [];
    
    // Look for common delimiters
    const delimiterCounts = lines.map(line => ({
      tabs: (line.match(/\t/g) || []).length,
      pipes: (line.match(/\|/g) || []).length,
      multiSpaces: (line.match(/\s{2,}/g) || []).length,
    }));

    // Determine most likely delimiter
    const avgTabs = delimiterCounts.reduce((s, d) => s + d.tabs, 0) / lines.length;
    const avgPipes = delimiterCounts.reduce((s, d) => s + d.pipes, 0) / lines.length;
    const avgSpaces = delimiterCounts.reduce((s, d) => s + d.multiSpaces, 0) / lines.length;

    let delimiter = /\s{2,}/; // Default: multiple spaces
    let isTable = false;

    if (avgTabs >= 2) {
      delimiter = /\t/;
      isTable = true;
    } else if (avgPipes >= 2) {
      delimiter = /\|/;
      isTable = true;
    } else if (avgSpaces >= 2) {
      delimiter = /\s{2,}/;
      isTable = true;
    }

    if (!isTable) {
      return { isTable: false };
    }

    // Parse rows
    const rows = lines.map(line => 
      line.split(delimiter)
        .map(cell => cell.trim())
        .filter(cell => cell.length > 0)
    );

    // First row as headers
    const headers = rows[0].map((h, i) => h || `Column_${i + 1}`);
    
    // Rest as data
    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const row = { _rowIndex: i + 1 };
      headers.forEach((h, idx) => {
        row[h] = rows[i][idx] || '';
      });
      data.push(row);
    }

    // Check for low-confidence cells
    data.forEach((row, i) => {
      headers.forEach(h => {
        const val = row[h];
        // OCR often confuses similar characters
        if (/[0O]{2,}/.test(val) || /[Il1]{3,}/.test(val)) {
          warnings.push({
            row: i + 2,
            column: h,
            value: val,
            issue: 'Possible OCR error (similar characters)',
          });
        }
      });
    });

    return {
      isTable: true,
      headers,
      data,
      warnings,
    };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // UTILITY METHODS
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Check if OCR is available
   */
  async isAvailable() {
    return await this.checkAvailability();
  }

  /**
   * Validate OCR result quality
   */
  validateQuality(result) {
    const issues = [];
    
    if (!result.success) {
      return { isGoodQuality: false, issues: [{ type: 'error', message: result.error }] };
    }

    if (result.confidence < 80) {
      issues.push({
        type: 'low_confidence',
        message: `OCR confidence is low (${result.confidence.toFixed(1)}%). Results may be inaccurate.`,
        suggestion: 'Try with a clearer image or better lighting.',
      });
    }

    if (result.warnings?.length > 0) {
      issues.push({
        type: 'ocr_warnings',
        message: `${result.warnings.length} cells may have OCR errors.`,
        details: result.warnings.slice(0, 5),
      });
    }

    return {
      isGoodQuality: issues.length === 0,
      issues,
    };
  }
}

module.exports = new ImageToExcel();
