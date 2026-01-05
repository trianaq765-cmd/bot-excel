// ═══════════════════════════════════════════════════════════════════════════
// ENGINE ORCHESTRATOR - Main entry point for all engine operations
// ═══════════════════════════════════════════════════════════════════════════

const fileParser = require('../utils/fileParser');
const helpers = require('../utils/helpers');

// Lazy load modules to prevent startup errors
let analyzer = null;
let cleaner = null;
let formatter = null;
let reporter = null;
let converter = null;
let textToExcel = null;
let instructionParser = null;
let templateEngine = null;
let imageToExcel = null;

// Safe require function
function safeRequire(modulePath, moduleName) {
  try {
    return require(modulePath);
  } catch (error) {
    console.error(`⚠️ Failed to load ${moduleName}:`, error.message);
    return null;
  }
}

// Initialize modules
function initModules() {
  if (!analyzer) {
    analyzer = safeRequire('./analyzer', 'analyzer');
    cleaner = safeRequire('./cleaner', 'cleaner');
    formatter = safeRequire('./formatter', 'formatter');
    reporter = safeRequire('./reporter', 'reporter');
    converter = safeRequire('./converter', 'converter');
    textToExcel = safeRequire('./generators/textToExcel', 'textToExcel');
    instructionParser = safeRequire('./generators/instructionParser', 'instructionParser');
    templateEngine = safeRequire('./generators/templateEngine', 'templateEngine');
    imageToExcel = safeRequire('./generators/imageToExcel', 'imageToExcel');
  }
}

class ExcelEngine {
  constructor() {
    this.fileParser = fileParser;
    initModules();
  }

  // ─────────────────────────────────────────────────────────────────────────
  // GETTERS (Lazy load)
  // ─────────────────────────────────────────────────────────────────────────

  get analyzer() { initModules(); return analyzer; }
  get cleaner() { initModules(); return cleaner; }
  get formatter() { initModules(); return formatter; }
  get reporter() { initModules(); return reporter; }
  get converter() { initModules(); return converter; }
  get textToExcel() { initModules(); return textToExcel; }
  get instructionParser() { initModules(); return instructionParser; }
  get templateEngine() { initModules(); return templateEngine; }
  get imageToExcel() { initModules(); return imageToExcel; }

  // ─────────────────────────────────────────────────────────────────────────
  // FULL ANALYSIS & CLEANING PIPELINE
  // ─────────────────────────────────────────────────────────────────────────

  async process(input, options = {}) {
    const startTime = Date.now();
    const result = {
      success: false,
      stages: {},
    };

    try {
      console.log('[Engine] Starting process...');

      // Stage 1: Parse file
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

      console.log(`[Engine] Parsed: ${parsed.rowCount} rows, ${parsed.columnCount} columns`);

      // Stage 2: Analyze
      console.log('[Engine] Stage 2: Analyzing data...');
      if (!this.analyzer) {
        throw new Error('Analyzer module not available');
      }
      
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

      console.log(`[Engine] Analysis complete: ${analysis.totalIssues} issues found`);

      // Stage 3: Clean (Auto-fix)
      console.log('[Engine] Stage 3: Cleaning data...');
      if (!this.cleaner) {
        throw new Error('Cleaner module not available');
      }
      
      const cleaned = this.cleaner.clean(parsed, analysis, options);
      
      result.stages.clean = {
        success: true,
        time: cleaned.cleanTime,
        changes: cleaned.stats.totalChanges,
        rowsRemoved: cleaned.stats.rowsRemoved,
        cellsModified: cleaned.stats.cellsModified,
      };

      console.log(`[Engine] Cleaning complete: ${cleaned.stats.totalChanges} changes`);

      // Stage 4: Format
      console.log('[Engine] Stage 4: Formatting output...');
      
      let formatOptions = options.formatOptions || {};
      if (options.instructions && this.instructionParser) {
        const instructionResult = this.instructionParser.parse(options.instructions);
        formatOptions = {
          ...formatOptions,
          ...this.instructionParser.toFormatOptions(instructionResult),
        };
      }

      // Stage 5: Generate report
      console.log('[Engine] Stage 5: Generating report...');
      
      let buffer;
      if (this.reporter) {
        const workbook = await this.reporter.generateReport({
          originalData: parsed,
          cleanedData: cleaned,
          analysisResult: analysis,
          cleaningResult: cleaned,
          options: formatOptions,
        });
        buffer = await this.reporter.toBuffer();
      } else if (this.formatter) {
        const workbook = await this.formatter.format({
          headers: cleaned.headers,
          data: cleaned.data,
          columnTypes: parsed.columnTypes,
        }, formatOptions);
        buffer = await this.formatter.toBuffer(workbook);
      } else {
        throw new Error('No formatter or reporter available');
      }

      result.stages.report = {
        success: true,
      };

      // Final result
      result.success = true;
      result.totalTime = Date.now() - startTime;
      result.totalTimeFormatted = helpers.formatDuration(Date.now() - startTime);
      result.output = {
        buffer,
        filename: `analyzed_${parsed.fileName || 'data'}.xlsx`,
      };

      // Summary
      result.summary = {
        originalRows: parsed.rowCount,
        cleanedRows: cleaned.data.length,
        rowsRemoved: cleaned.stats.rowsRemoved,
        issuesFixed: analysis.autoFixCount,
        issuesNeedReview: analysis.needsReviewCount,
        qualityBefore: analysis.qualityScore?.score || 0,
        qualityAfter: this._calculateAfterQuality(analysis.qualityScore?.score || 0, cleaned),
      };

      result.data = {
        headers: cleaned.headers,
        data: cleaned.data,
      };

      result.analysis = analysis;
      result.changes = cleaned.changes;

      console.log(`[Engine] Process complete in ${result.totalTimeFormatted}`);
      return result;

    } catch (error) {
      console.error('[Engine] Process error:', error);
      result.success = false;
      result.error = error.message;
      result.totalTime = Date.now() - startTime;
      return result;
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // QUICK OPERATIONS
  // ─────────────────────────────────────────────────────────────────────────

  async analyze(input, options = {}) {
    try {
      const parsed = this.fileParser.parse(input, options);
      if (!parsed.success) throw new Error(parsed.error);
      
      if (!this.analyzer) throw new Error('Analyzer not available');
      return this.analyzer.analyze(parsed, options);
    } catch (error) {
      console.error('[Engine] Analyze error:', error);
      throw error;
    }
  }

  async quickClean(input, options = {}) {
    try {
      const parsed = this.fileParser.parse(input, options);
      if (!parsed.success) throw new Error(parsed.error);
      
      if (!this.cleaner) throw new Error('Cleaner not available');
      const cleaned = this.cleaner.basicClean(parsed.headers, parsed.data, options);
      
      if (!this.formatter) throw new Error('Formatter not available');
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
    } catch (error) {
      console.error('[Engine] QuickClean error:', error);
      throw error;
    }
  }

  async convert(input, targetFormat, options = {}) {
    try {
      const parsed = this.fileParser.parse(input, options);
      if (!parsed.success) throw new Error(parsed.error);

      if (!this.converter) throw new Error('Converter not available');

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

      return { success: true, format: targetFormat, output };
    } catch (error) {
      console.error('[Engine] Convert error:', error);
      throw error;
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // GENERATION OPERATIONS
  // ─────────────────────────────────────────────────────────────────────────

  async createFromText(text, options = {}) {
    try {
      if (!this.textToExcel) throw new Error('TextToExcel not available');
      
      const parsed = this.textToExcel.parse(text, options);
      if (!parsed.success) throw new Error(parsed.error || 'Failed to parse text');

      if (!this.formatter) throw new Error('Formatter not available');
      
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
    } catch (error) {
      console.error('[Engine] CreateFromText error:', error);
      throw error;
    }
  }

  async createFromImage(imageBuffer, options = {}) {
    try {
      if (!this.imageToExcel) {
        throw new Error('Image extraction not available. OCR module not loaded.');
      }

      const available = await this.imageToExcel.isAvailable();
      if (!available) {
        throw new Error('OCR feature not available. Please install tesseract.js');
      }

      const extracted = await this.imageToExcel.extract(imageBuffer, options);
      
      if (!extracted.success) {
        throw new Error(extracted.error);
      }

      const quality = this.imageToExcel.validateQuality(extracted);

      if (!this.formatter) throw new Error('Formatter not available');
      
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
    } catch (error) {
      console.error('[Engine] CreateFromImage error:', error);
      throw error;
    }
  }

  async applyFormat(input, instructions, options = {}) {
    try {
      const parsed = this.fileParser.parse(input, options);
      if (!parsed.success) throw new Error(parsed.error);

      if (!this.instructionParser) throw new Error('InstructionParser not available');
      if (!this.formatter) throw new Error('Formatter not available');

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
    } catch (error) {
      console.error('[Engine] ApplyFormat error:', error);
      throw error;
    }
  }

  async generateTemplate(templateName, options = {}) {
    try {
      if (!this.templateEngine) throw new Error('TemplateEngine not available');
      
      const workbook = await this.templateEngine.generate(templateName, options);
      
      return {
        success: true,
        template: templateName,
        buffer: await workbook.xlsx.writeBuffer(),
      };
    } catch (error) {
      console.error('[Engine] GenerateTemplate error:', error);
      throw error;
    }
  }

  listTemplates() {
    if (!this.templateEngine) return [];
    return this.templateEngine.listTemplates();
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HELPER METHODS
  // ─────────────────────────────────────────────────────────────────────────

  _calculateAfterQuality(beforeScore, cleanedResult) {
    const fixes = cleanedResult?.stats?.totalChanges || 0;
    const improvement = Math.min(50, fixes * 0.5);
    return Math.min(100, Math.round(beforeScore + improvement));
  }

  async cleanup() {
    if (this.imageToExcel) {
      await this.imageToExcel.terminate();
    }
  }
}

module.exports = new ExcelEngine();
