// ═══════════════════════════════════════════════════════════════════════════
// ENGINE ORCHESTRATOR - Main entry point for all engine operations
// ═══════════════════════════════════════════════════════════════════════════

const fileParser = require('../utils/fileParser');
const analyzer = require('./analyzer');
const cleaner = require('./cleaner');
const formatter = require('./formatter');
const reporter = require('./reporter');
const converter = require('./converter');
const textToExcel = require('./generators/textToExcel');
const instructionParser = require('./generators/instructionParser');
const templateEngine = require('./generators/templateEngine');
const imageToExcel = require('./generators/imageToExcel');
const helpers = require('../utils/helpers');

class ExcelEngine {
  constructor() {
    this.fileParser = fileParser;
    this.analyzer = analyzer;
    this.cleaner = cleaner;
    this.formatter = formatter;
    this.reporter = reporter;
    this.converter = converter;
    this.textToExcel = textToExcel;
    this.instructionParser = instructionParser;
    this.templateEngine = templateEngine;
    this.imageToExcel = imageToExcel;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // FULL ANALYSIS & CLEANING PIPELINE
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Full intelligent analysis and cleaning pipeline
   * @param {Buffer|string} input - File buffer or path
   * @param {Object} options - Processing options
   * @returns {Object} Complete result with cleaned data and report
   */
  async process(input, options = {}) {
    const startTime = Date.now();
    const result = {
      success: false,
      stages: {},
    };

    try {
      // ─────────────────────────────────────────────────────────────────────
      // Stage 1: Parse file
      // ─────────────────────────────────────────────────────────────────────
      console.log('[Engine] Stage 1: Parsing file...');
      const parsed = this.fileParser.parse(input, options);
      
      if (!parsed.success) {
        throw new Error(parsed.error || 'Failed to parse file');
      }
      
      result.stages.parse = {
        success: true,
        time: parsed.parseTime,
        rows: parsed.rowCount,
        columns: parsed.columnCount,
      };

      // ─────────────────────────────────────────────────────────────────────
      // Stage 2: Analyze
      // ─────────────────────────────────────────────────────────────────────
      console.log('[Engine] Stage 2: Analyzing data...');
      const analysis = this.analyzer.analyze(parsed, options);
      
      result.stages.analyze = {
        success: analysis.success,
        time: analysis.analysisTime,
        qualityScore: analysis.qualityScore,
        issuesFound: analysis.totalIssues,
        autoFixable: analysis.autoFixCount,
        needsReview: analysis.needsReviewCount,
        critical: analysis.criticalCount,
      };

      // ─────────────────────────────────────────────────────────────────────
      // Stage 3: Clean (Auto-fix)
      // ─────────────────────────────────────────────────────────────────────
      console.log('[Engine] Stage 3: Cleaning data...');
      const cleaned = this.cleaner.clean(parsed, analysis, options);
      
      result.stages.clean = {
        success: cleaned.success,
        time: cleaned.cleanTime,
        changes: cleaned.stats.totalChanges,
        rowsRemoved: cleaned.stats.rowsRemoved,
        cellsModified: cleaned.stats.cellsModified,
      };

      // ─────────────────────────────────────────────────────────────────────
      // Stage 4: Format (Professional styling)
      // ─────────────────────────────────────────────────────────────────────
      console.log('[Engine] Stage 4: Formatting output...');
      
      // Apply instruction-based formatting if provided
      let formatOptions = options.formatOptions || {};
      if (options.instructions) {
        const instructionResult = this.instructionParser.parse(options.instructions);
        formatOptions = {
          ...formatOptions,
          ...this.instructionParser.toFormatOptions(instructionResult),
        };
      }

      // ─────────────────────────────────────────────────────────────────────
      // Stage 5: Generate report
      // ─────────────────────────────────────────────────────────────────────
      console.log('[Engine] Stage 5: Generating report...');
      const workbook = await this.reporter.generateReport({
        originalData: parsed,
        cleanedData: cleaned,
        analysisResult: analysis,
        cleaningResult: cleaned,
        options: formatOptions,
      });

      const buffer = await this.reporter.toBuffer();

      result.stages.report = {
        success: true,
        sheets: workbook.worksheets.length,
      };

      // ─────────────────────────────────────────────────────────────────────
      // Final result
      // ─────────────────────────────────────────────────────────────────────
      result.success = true;
      result.totalTime = Date.now() - startTime;
      result.totalTimeFormatted = helpers.formatDuration(Date.now() - startTime);
      result.output = {
        buffer,
        filename: `analyzed_${parsed.fileName || 'data'}.xlsx`,
      };

      // Summary for quick access
      result.summary = {
        originalRows: parsed.rowCount,
        cleanedRows: cleaned.data.length,
        rowsRemoved: cleaned.stats.rowsRemoved,
        issuesFixed: analysis.autoFixCount,
        issuesNeedReview: analysis.needsReviewCount,
        qualityBefore: analysis.qualityScore.score,
        qualityAfter: this._calculateAfterQuality(analysis.qualityScore.score, cleaned),
      };

      // Include data for further processing
      result.data = {
        headers: cleaned.headers,
        data: cleaned.data,
      };

      result.analysis = analysis;
      result.changes = cleaned.changes;

      return result;

    } catch (error) {
      result.success = false;
      result.error = error.message;
      result.totalTime = Date.now() - startTime;
      return result;
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // QUICK OPERATIONS (Individual functions)
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Quick analysis only (no cleaning)
   */
  async analyze(input, options = {}) {
    const parsed = this.fileParser.parse(input, options);
    if (!parsed.success) throw new Error(parsed.error);
    
    return this.analyzer.analyze(parsed, options);
  }

  /**
   * Quick clean (basic cleaning without full analysis)
   */
  async quickClean(input, options = {}) {
    const parsed = this.fileParser.parse(input, options);
    if (!parsed.success) throw new Error(parsed.error);
    
    const cleaned = this.cleaner.basicClean(parsed.headers, parsed.data, options);
    
    const workbook = await this.formatter.format({
      headers: cleaned.headers,
      data: cleaned.data,
      columnTypes: parsed.columnTypes,
    }, options);

    return {
      success: true,
      buffer: await this.formatter.toBuffer(workbook),
      stats: cleaned.stats,
      changes: cleaned.changes,
    };
  }

  /**
   * Convert format only
   */
  async convert(input, targetFormat, options = {}) {
    const parsed = this.fileParser.parse(input, options);
    if (!parsed.success) throw new Error(parsed.error);

    let output;
    const { headers, data } = parsed;

    switch (targetFormat.toLowerCase()) {
      case 'csv':
        output = this.converter.toCSV(headers, data, options);
        break;
      case 'json':
        output = this.converter.toJSON(headers, data, options);
        break;
      case 'html':
        output = this.converter.toHTML(headers, data, options);
        break;
      case 'markdown':
      case 'md':
        output = this.converter.toMarkdown(headers, data, options);
        break;
      case 'sql':
        output = this.converter.toSQL(headers, data, options);
        break;
      case 'xml':
        output = this.converter.toXML(headers, data, options);
        break;
      default:
        throw new Error(`Unsupported format: ${targetFormat}`);
    }

    return {
      success: true,
      format: targetFormat,
      output,
    };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // GENERATION OPERATIONS
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Create Excel from text input
   */
  async createFromText(text, options = {}) {
    const parsed = this.textToExcel.parse(text, options);
    if (!parsed.success) throw new Error(parsed.error);

    const workbook = await this.formatter.format({
      headers: parsed.headers,
      data: parsed.data,
      columnTypes: {},
    }, {
      zebraStripes: true,
      borders: true,
      addTotals: options.addTotals,
      ...options,
    });

    return {
      success: true,
      format: parsed.format,
      headers: parsed.headers,
      rowCount: parsed.data.length,
      buffer: await this.formatter.toBuffer(workbook),
    };
  }

  /**
   * Create Excel from image (OCR)
   */
  async createFromImage(imageBuffer, options = {}) {
    const extracted = await this.imageToExcel.extract(imageBuffer, options);
    
    if (!extracted.success) {
      throw new Error(extracted.error);
    }

    const quality = this.imageToExcel.validateQuality(extracted);

    const workbook = await this.formatter.format({
      headers: extracted.headers,
      data: extracted.data,
      columnTypes: {},
    }, options);

    return {
      success: true,
      confidence: extracted.confidence,
      tableDetected: extracted.tableDetected,
      headers: extracted.headers,
      rowCount: extracted.data.length,
      warnings: extracted.warnings,
      quality,
      buffer: await this.formatter.toBuffer(workbook),
    };
  }

  /**
   * Apply custom formatting based on instructions
   */
  async applyFormat(input, instructions, options = {}) {
    const parsed = this.fileParser.parse(input, options);
    if (!parsed.success) throw new Error(parsed.error);

    const instructionResult = this.instructionParser.parse(instructions);
    const formatOptions = this.instructionParser.toFormatOptions(instructionResult);

    const workbook = await this.formatter.format({
      headers: parsed.headers,
      data: parsed.data,
      columnTypes: parsed.columnTypes,
    }, formatOptions);

    return {
      success: true,
      instructionsApplied: instructionResult.instructionCount,
      instructions: instructionResult.instructions,
      buffer: await this.formatter.toBuffer(workbook),
    };
  }

  /**
   * Generate template
   */
  async generateTemplate(templateName, options = {}) {
    const workbook = await this.templateEngine.generate(templateName, options);
    
    return {
      success: true,
      template: templateName,
      buffer: await workbook.xlsx.writeBuffer(),
    };
  }

  /**
   * List available templates
   */
  listTemplates() {
    return this.templateEngine.listTemplates();
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HELPER METHODS
  // ─────────────────────────────────────────────────────────────────────────

  _calculateAfterQuality(beforeScore, cleanedResult) {
    const fixes = cleanedResult.stats?.totalChanges || 0;
    const improvement = Math.min(50, fixes * 0.5);
    return Math.min(100, Math.round(beforeScore + improvement));
  }

  /**
   * Cleanup resources
   */
  async cleanup() {
    await this.imageToExcel.terminate();
  }
}

// Export singleton instance
module.exports = new ExcelEngine();
