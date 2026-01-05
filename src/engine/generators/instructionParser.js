// ═══════════════════════════════════════════════════════════════════════════
// INSTRUCTION PARSER - Parse natural language formatting instructions
// ═══════════════════════════════════════════════════════════════════════════

const { FORMATTING } = require('../../utils/constants');

class InstructionParser {
  constructor() {
    // ─────────────────────────────────────────────────────────────────────
    // INSTRUCTION PATTERNS (Indonesian + English)
    // ─────────────────────────────────────────────────────────────────────
    this.patterns = {
      // Header styling
      headerColor: [
        /header\s*(?:warna|color|background)?\s*[:=]?\s*(biru|merah|hijau|kuning|hitam|putih|blue|red|green|yellow|black|white|#[0-9a-fA-F]{6})/i,
        /(?:warna|color)\s*header\s*[:=]?\s*(biru|merah|hijau|kuning|hitam|putih|blue|red|green|yellow|black|white|#[0-9a-fA-F]{6})/i,
      ],
      
      // Column formatting
      columnCurrency: [
        /(?:kolom|column)\s*[""']?([^""']+)[""']?\s*(?:format|sebagai|as)?\s*(?:currency|rupiah|mata\s*uang|uang)/i,
        /format\s*(?:kolom|column)?\s*[""']?([^""']+)[""']?\s*(?:jadi|ke|to|as)?\s*(?:currency|rupiah)/i,
      ],
      
      columnDate: [
        /(?:kolom|column)\s*[""']?([^""']+)[""']?\s*(?:format|sebagai|as)?\s*(?:tanggal|date)/i,
        /format\s*(?:kolom|column)?\s*[""']?([^""']+)[""']?\s*(?:jadi|ke|to|as)?\s*(?:tanggal|date)/i,
      ],
      
      columnNumber: [
        /(?:kolom|column)\s*[""']?([^""']+)[""']?\s*(?:format|sebagai|as)?\s*(?:angka|number|numeric)/i,
      ],
      
      columnPercentage: [
        /(?:kolom|column)\s*[""']?([^""']+)[""']?\s*(?:format|sebagai|as)?\s*(?:persen|percentage|percent|%)/i,
      ],
      
      // Date format specification
      dateFormat: [
        /(?:format\s*tanggal|date\s*format)\s*[:=]?\s*(DD[-\/]MM[-\/]YYYY|DD[-\/]MMM[-\/]YYYY|YYYY[-\/]MM[-\/]DD)/i,
      ],
      
      // Add totals
      addTotals: [
        /(?:tambah|add|buat|create)\s*(?:baris|row)?\s*(?:total|sum|jumlah)/i,
        /(?:total|sum|jumlah)\s*(?:row|baris)?\s*(?:di\s*bawah|at\s*bottom)/i,
      ],
      
      // Freeze panes
      freezeHeader: [
        /(?:freeze|bekukan|kunci)\s*(?:header|baris\s*pertama|first\s*row|row\s*1)/i,
        /(?:header|baris\s*pertama)\s*(?:freeze|bekukan|kunci)/i,
      ],
      
      freezeColumn: [
        /(?:freeze|bekukan|kunci)\s*(?:kolom|column)\s*(?:pertama|first|1|A)/i,
      ],
      
      // Zebra stripes
      zebraStripes: [
        /(?:zebra|alternating|selang[-\s]?seling)\s*(?:stripes|warna|color|rows|baris)/i,
        /(?:warna|color)\s*(?:baris|row)\s*(?:selang[-\s]?seling|alternating|bergantian)/i,
      ],
      
      // Borders
      borders: [
        /(?:tambah|add|buat)\s*(?:border|garis|bingkai)/i,
        /(?:border|garis|bingkai)\s*(?:semua|all|full)/i,
      ],
      
      // Auto-fit columns
      autoFit: [
        /(?:auto[-\s]?fit|sesuaikan)\s*(?:kolom|column|lebar|width)/i,
        /(?:lebar|width)\s*(?:kolom|column)\s*(?:otomatis|auto)/i,
      ],
      
      // Text case
      textCase: [
        /(?:text|teks|huruf)\s*(?:case)?\s*[:=]?\s*(uppercase|lowercase|titlecase|title\s*case|capital|kapital)/i,
        /(?:kolom|column)\s*[""']?([^""']+)[""']?\s*(?:jadi|to|as)?\s*(uppercase|lowercase|titlecase)/i,
      ],
      
      // Conditional formatting
      highlightDuplicates: [
        /(?:highlight|sorot|tandai)\s*(?:duplikat|duplicate)/i,
      ],
      
      highlightEmpty: [
        /(?:highlight|sorot|tandai)\s*(?:kosong|empty|blank)/i,
      ],
      
      highlightNegative: [
        /(?:highlight|sorot|tandai)\s*(?:negatif|negative|minus)/i,
      ],
      
      // Summary/statistics
      addSummary: [
        /(?:tambah|add|buat)\s*(?:sheet|lembar)?\s*(?:summary|ringkasan|statistik)/i,
        /(?:summary|ringkasan)\s*(?:sheet|lembar)/i,
      ],
      
      // Custom title
      title: [
        /(?:judul|title)\s*[:=]?\s*[""']?([^""'\n]+)[""']?/i,
        /(?:report|laporan)\s*(?:title|judul)\s*[:=]?\s*[""']?([^""'\n]+)[""']?/i,
      ],
      
      // Print settings
      printLandscape: [
        /(?:print|cetak)\s*(?:landscape|horizontal)/i,
        /(?:orientasi|orientation)\s*[:=]?\s*(?:landscape|horizontal)/i,
      ],
      
      printPortrait: [
        /(?:print|cetak)\s*(?:portrait|vertical)/i,
        /(?:orientasi|orientation)\s*[:=]?\s*(?:portrait|vertical)/i,
      ],
    };

    // Color mapping
    this.colorMap = {
      // Indonesian
      'biru': 'FF2B579A',
      'merah': 'FFFF0000',
      'hijau': 'FF00B050',
      'kuning': 'FFFFC000',
      'hitam': 'FF000000',
      'putih': 'FFFFFFFF',
      'abu': 'FF808080',
      'orange': 'FFFF6600',
      'ungu': 'FF7030A0',
      
      // English
      'blue': 'FF2B579A',
      'red': 'FFFF0000',
      'green': 'FF00B050',
      'yellow': 'FFFFC000',
      'black': 'FF000000',
      'white': 'FFFFFFFF',
      'gray': 'FF808080',
      'grey': 'FF808080',
      'purple': 'FF7030A0',
      
      // Professional colors
      'navy': 'FF1F4E79',
      'teal': 'FF008080',
      'maroon': 'FF800000',
      'olive': 'FF808000',
    };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // MAIN PARSE METHOD
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Parse instruction text into structured commands
   * @param {string} instructionText - Natural language instructions
   * @returns {Object} Parsed instructions
   */
  parse(instructionText) {
    const text = instructionText.toLowerCase();
    const instructions = [];
    const parsed = {
      formatting: {},
      columns: {},
      options: {},
    };

    // ─────────────────────────────────────────────────────────────────────
    // Parse each pattern type
    // ─────────────────────────────────────────────────────────────────────

    // Header color
    for (const pattern of this.patterns.headerColor) {
      const match = instructionText.match(pattern);
      if (match) {
        const color = this._parseColor(match[1]);
        parsed.formatting.headerColor = color;
        instructions.push({
          type: 'headerColor',
          color: color,
          original: match[0],
        });
        break;
      }
    }

    // Column currency format
    for (const pattern of this.patterns.columnCurrency) {
      const matches = instructionText.matchAll(new RegExp(pattern, 'gi'));
      for (const match of matches) {
        const columnName = match[1].trim();
        if (!parsed.columns[columnName]) parsed.columns[columnName] = {};
        parsed.columns[columnName].format = 'currency';
        instructions.push({
          type: 'columnFormat',
          column: columnName,
          format: 'currency',
          original: match[0],
        });
      }
    }

    // Column date format
    for (const pattern of this.patterns.columnDate) {
      const matches = instructionText.matchAll(new RegExp(pattern, 'gi'));
      for (const match of matches) {
        const columnName = match[1].trim();
        if (!parsed.columns[columnName]) parsed.columns[columnName] = {};
        parsed.columns[columnName].format = 'date';
        instructions.push({
          type: 'columnFormat',
          column: columnName,
          format: 'date',
          original: match[0],
        });
      }
    }

    // Column number format
    for (const pattern of this.patterns.columnNumber) {
      const matches = instructionText.matchAll(new RegExp(pattern, 'gi'));
      for (const match of matches) {
        const columnName = match[1].trim();
        if (!parsed.columns[columnName]) parsed.columns[columnName] = {};
        parsed.columns[columnName].format = 'number';
        instructions.push({
          type: 'columnFormat',
          column: columnName,
          format: 'number',
          original: match[0],
        });
      }
    }

    // Column percentage format
    for (const pattern of this.patterns.columnPercentage) {
      const matches = instructionText.matchAll(new RegExp(pattern, 'gi'));
      for (const match of matches) {
        const columnName = match[1].trim();
        if (!parsed.columns[columnName]) parsed.columns[columnName] = {};
        parsed.columns[columnName].format = 'percentage';
        instructions.push({
          type: 'columnFormat',
          column: columnName,
          format: 'percentage',
          original: match[0],
        });
      }
    }

    // Date format specification
    for (const pattern of this.patterns.dateFormat) {
      const match = instructionText.match(pattern);
      if (match) {
        parsed.options.dateFormat = match[1].toUpperCase();
        instructions.push({
          type: 'dateFormat',
          format: match[1].toUpperCase(),
          original: match[0],
        });
        break;
      }
    }

    // Add totals
    for (const pattern of this.patterns.addTotals) {
      if (pattern.test(instructionText)) {
        parsed.options.addTotals = true;
        instructions.push({
          type: 'addTotals',
          original: instructionText.match(pattern)[0],
        });
        break;
      }
    }

    // Freeze header
    for (const pattern of this.patterns.freezeHeader) {
      if (pattern.test(instructionText)) {
        parsed.options.freezeHeader = true;
        instructions.push({
          type: 'freezeHeader',
          original: instructionText.match(pattern)[0],
        });
        break;
      }
    }

    // Freeze column
    for (const pattern of this.patterns.freezeColumn) {
      if (pattern.test(instructionText)) {
        parsed.options.freezeColumn = true;
        instructions.push({
          type: 'freezeColumn',
          original: instructionText.match(pattern)[0],
        });
        break;
      }
    }

    // Zebra stripes
    for (const pattern of this.patterns.zebraStripes) {
      if (pattern.test(instructionText)) {
        parsed.options.zebraStripes = true;
        instructions.push({
          type: 'zebraStripes',
          original: instructionText.match(pattern)[0],
        });
        break;
      }
    }

    // Borders
    for (const pattern of this.patterns.borders) {
      if (pattern.test(instructionText)) {
        parsed.options.borders = true;
        instructions.push({
          type: 'borders',
          original: instructionText.match(pattern)[0],
        });
        break;
      }
    }

    // Auto-fit
    for (const pattern of this.patterns.autoFit) {
      if (pattern.test(instructionText)) {
        parsed.options.autoFit = true;
        instructions.push({
          type: 'autoFit',
          original: instructionText.match(pattern)[0],
        });
        break;
      }
    }

    // Add summary sheet
    for (const pattern of this.patterns.addSummary) {
      if (pattern.test(instructionText)) {
        parsed.options.addSummary = true;
        instructions.push({
          type: 'addSummary',
          original: instructionText.match(pattern)[0],
        });
        break;
      }
    }

    // Highlight duplicates
    for (const pattern of this.patterns.highlightDuplicates) {
      if (pattern.test(instructionText)) {
        parsed.options.highlightDuplicates = true;
        instructions.push({
          type: 'highlightDuplicates',
          original: instructionText.match(pattern)[0],
        });
        break;
      }
    }

    // Highlight empty
    for (const pattern of this.patterns.highlightEmpty) {
      if (pattern.test(instructionText)) {
        parsed.options.highlightEmpty = true;
        instructions.push({
          type: 'highlightEmpty',
          original: instructionText.match(pattern)[0],
        });
        break;
      }
    }

    // Title
    for (const pattern of this.patterns.title) {
      const match = instructionText.match(pattern);
      if (match) {
        parsed.options.title = match[1].trim();
        instructions.push({
          type: 'title',
          value: match[1].trim(),
          original: match[0],
        });
        break;
      }
    }

    // Print orientation
    for (const pattern of this.patterns.printLandscape) {
      if (pattern.test(instructionText)) {
        parsed.options.orientation = 'landscape';
        instructions.push({ type: 'orientation', value: 'landscape' });
        break;
      }
    }

    for (const pattern of this.patterns.printPortrait) {
      if (pattern.test(instructionText)) {
        parsed.options.orientation = 'portrait';
        instructions.push({ type: 'orientation', value: 'portrait' });
        break;
      }
    }

    return {
      success: true,
      instructionCount: instructions.length,
      instructions,
      parsed,
    };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HELPER METHODS
  // ─────────────────────────────────────────────────────────────────────────

  _parseColor(colorInput) {
    const color = colorInput.toLowerCase().trim();
    
    // Check if it's a hex color
    if (color.startsWith('#')) {
      return 'FF' + color.slice(1).toUpperCase();
    }
    
    // Look up in color map
    return this.colorMap[color] || FORMATTING.HEADER_COLOR;
  }

  /**
   * Generate format options from parsed instructions
   */
  toFormatOptions(parsedResult) {
    const { parsed } = parsedResult;
    
    return {
      headerColor: parsed.formatting.headerColor,
      columnFormats: parsed.columns,
      addTotals: parsed.options.addTotals || false,
      addSummary: parsed.options.addSummary || false,
      zebraStripes: parsed.options.zebraStripes !== false,
      borders: parsed.options.borders !== false,
      autoFit: parsed.options.autoFit !== false,
      freezeHeader: parsed.options.freezeHeader || false,
      freezeColumn: parsed.options.freezeColumn || false,
      dateFormat: parsed.options.dateFormat || 'DD-MMM-YYYY',
      title: parsed.options.title,
      orientation: parsed.options.orientation || 'portrait',
      highlightDuplicates: parsed.options.highlightDuplicates || false,
      highlightEmpty: parsed.options.highlightEmpty || false,
    };
  }

  /**
   * Get human-readable summary of parsed instructions
   */
  getSummary(parsedResult) {
    const { instructions } = parsedResult;
    
    if (instructions.length === 0) {
      return 'No formatting instructions detected.';
    }

    const summaryLines = ['Detected formatting instructions:'];
    
    instructions.forEach((inst, i) => {
      switch (inst.type) {
        case 'headerColor':
          summaryLines.push(`${i + 1}. Header color: ${inst.color}`);
          break;
        case 'columnFormat':
          summaryLines.push(`${i + 1}. Column "${inst.column}": ${inst.format} format`);
          break;
        case 'addTotals':
          summaryLines.push(`${i + 1}. Add total row`);
          break;
        case 'freezeHeader':
          summaryLines.push(`${i + 1}. Freeze header row`);
          break;
        case 'zebraStripes':
          summaryLines.push(`${i + 1}. Apply zebra stripes`);
          break;
        case 'addSummary':
          summaryLines.push(`${i + 1}. Add summary sheet`);
          break;
        case 'title':
          summaryLines.push(`${i + 1}. Report title: "${inst.value}"`);
          break;
        default:
          summaryLines.push(`${i + 1}. ${inst.type}`);
      }
    });

    return summaryLines.join('\n');
  }
}

module.exports = new InstructionParser();
