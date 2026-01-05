// ═══════════════════════════════════════════════════════════════════════════
// TEXT TO EXCEL GENERATOR - Parse text data into Excel
// ═══════════════════════════════════════════════════════════════════════════

const helpers = require('../../utils/helpers');

class TextToExcel {
  constructor() {
    this.patterns = {
      // CSV-like: value1, value2, value3
      csv: /^(.+?)(?:,|\t)(.+)/,
      
      // Key-value: key: value or key = value
      keyValue: /^([^:=]+)[:=]\s*(.+)$/,
      
      // Table-like with pipes: | col1 | col2 |
      pipeTable: /^\|(.+)\|$/,
      
      // List item: - item or * item or 1. item
      listItem: /^[-*•]\s+(.+)$|^(\d+)[.)]\s+(.+)$/,
      
      // Indonesian data patterns
      invoicePattern: /(?:invoice|inv|faktur)[:\s#]*([^\s,]+)/i,
      datePattern: /(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}|\d{1,2}\s+\w+\s+\d{4})/,
      currencyPattern: /(?:Rp\.?\s?|IDR\s?)([\d.,]+)/i,
      phonePattern: /(?:\+62|62|0)[\d\s.-]{8,14}/,
    };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // MAIN PARSE METHOD
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Parse text input into structured data
   * @param {string} text - Input text
   * @param {Object} options - Parse options
   * @returns {Object} { headers, data, format }
   */
  parse(text, options = {}) {
    const lines = text.trim().split(/\r?\n/).filter(l => l.trim());
    
    if (lines.length === 0) {
      return { success: false, error: 'No data found in text' };
    }

    // Try to detect format
    const format = this._detectFormat(lines);
    
    let result;
    switch (format) {
      case 'csv':
        result = this._parseCSV(lines, options);
        break;
      case 'pipe-table':
        result = this._parsePipeTable(lines);
        break;
      case 'key-value':
        result = this._parseKeyValue(lines);
        break;
      case 'list':
        result = this._parseList(lines, options);
        break;
      case 'natural':
        result = this._parseNaturalText(lines, options);
        break;
      default:
        result = this._parseGeneric(lines, options);
    }

    return {
      success: true,
      format,
      ...result,
    };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // FORMAT DETECTION
  // ─────────────────────────────────────────────────────────────────────────

  _detectFormat(lines) {
    const firstLine = lines[0];
    
    // Check for pipe table format
    if (firstLine.startsWith('|') && firstLine.endsWith('|')) {
      return 'pipe-table';
    }
    
    // Check for CSV/TSV format
    const commaCount = (firstLine.match(/,/g) || []).length;
    const tabCount = (firstLine.match(/\t/g) || []).length;
    
    if (commaCount >= 2 || tabCount >= 2) {
      return 'csv';
    }
    
    // Check for key-value pairs
    const kvMatches = lines.filter(l => this.patterns.keyValue.test(l)).length;
    if (kvMatches > lines.length * 0.7) {
      return 'key-value';
    }
    
    // Check for list format
    const listMatches = lines.filter(l => this.patterns.listItem.test(l)).length;
    if (listMatches > lines.length * 0.5) {
      return 'list';
    }
    
    // Natural text (try to extract structured data)
    return 'natural';
  }

  // ─────────────────────────────────────────────────────────────────────────
  // CSV PARSING
  // ─────────────────────────────────────────────────────────────────────────

  _parseCSV(lines, options) {
    const delimiter = options.delimiter || this._detectDelimiter(lines[0]);
    
    // First line as headers
    const headers = this._splitLine(lines[0], delimiter)
      .map(h => h.trim())
      .map((h, i) => h || `Column_${i + 1}`);
    
    // Rest as data
    const data = [];
    for (let i = 1; i < lines.length; i++) {
      const values = this._splitLine(lines[i], delimiter);
      const row = { _rowIndex: i + 1 };
      
      headers.forEach((h, idx) => {
        row[h] = values[idx]?.trim() || '';
      });
      
      data.push(row);
    }

    // Auto-detect and convert types
    this._autoConvertTypes(headers, data);

    return { headers, data };
  }

  _detectDelimiter(line) {
    const delimiters = [',', '\t', ';', '|'];
    let maxCount = 0;
    let detected = ',';

    delimiters.forEach(d => {
      const count = (line.match(new RegExp(d === '|' ? '\\|' : d, 'g')) || []).length;
      if (count > maxCount) {
        maxCount = count;
        detected = d;
      }
    });

    return detected;
  }

  _splitLine(line, delimiter) {
    const result = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];

      if (inQuotes) {
        if (char === '"') {
          if (line[i + 1] === '"') {
            current += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          current += char;
        }
      } else {
        if (char === '"') {
          inQuotes = true;
        } else if (char === delimiter) {
          result.push(current);
          current = '';
        } else {
          current += char;
        }
      }
    }

    result.push(current);
    return result;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // PIPE TABLE PARSING
  // ─────────────────────────────────────────────────────────────────────────

  _parsePipeTable(lines) {
    // Filter out separator lines (like |---|---|)
    const dataLines = lines.filter(l => !l.match(/^\|[\s-]+\|$/));
    
    if (dataLines.length === 0) {
      return { headers: [], data: [] };
    }

    // Parse headers
    const headers = dataLines[0]
      .split('|')
      .filter(c => c.trim())
      .map(c => c.trim());

    // Parse data
    const data = [];
    for (let i = 1; i < dataLines.length; i++) {
      const values = dataLines[i]
        .split('|')
        .filter(c => c !== '');
      
      const row = { _rowIndex: i + 1 };
      headers.forEach((h, idx) => {
        row[h] = values[idx]?.trim() || '';
      });
      
      data.push(row);
    }

    return { headers, data };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // KEY-VALUE PARSING
  // ─────────────────────────────────────────────────────────────────────────

  _parseKeyValue(lines) {
    const headers = ['Key', 'Value'];
    const data = [];

    lines.forEach((line, index) => {
      const match = line.match(this.patterns.keyValue);
      if (match) {
        data.push({
          _rowIndex: index + 2,
          Key: match[1].trim(),
          Value: match[2].trim(),
        });
      }
    });

    return { headers, data };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // LIST PARSING
  // ─────────────────────────────────────────────────────────────────────────

  _parseList(lines, options) {
    const headers = options.listHeader 
      ? [options.listHeader] 
      : ['No', 'Item'];
    
    const data = [];
    let counter = 1;

    lines.forEach((line, index) => {
      const match = line.match(this.patterns.listItem);
      if (match) {
        const item = match[1] || match[3];
        
        if (headers.length === 1) {
          data.push({
            _rowIndex: index + 2,
            [headers[0]]: item.trim(),
          });
        } else {
          data.push({
            _rowIndex: index + 2,
            No: counter++,
            Item: item.trim(),
          });
        }
      }
    });

    return { headers, data };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // NATURAL TEXT PARSING
  // ─────────────────────────────────────────────────────────────────────────

  _parseNaturalText(lines, options) {
    const text = lines.join('\n');
    const extracted = [];

    // Try to extract structured data from natural text
    // Pattern: "item description, quantity, price, date"
    
    lines.forEach((line, index) => {
      const entry = { _rowIndex: index + 2 };
      
      // Extract currency
      const currencyMatch = line.match(this.patterns.currencyPattern);
      if (currencyMatch) {
        entry['Amount'] = helpers.parseNumber(currencyMatch[1]);
      }

      // Extract date
      const dateMatch = line.match(this.patterns.datePattern);
      if (dateMatch) {
        entry['Date'] = dateMatch[1];
      }

      // Extract phone
      const phoneMatch = line.match(this.patterns.phonePattern);
      if (phoneMatch) {
        entry['Phone'] = phoneMatch[0];
      }

      // Extract remaining text as description
      let description = line
        .replace(this.patterns.currencyPattern, '')
        .replace(this.patterns.datePattern, '')
        .replace(this.patterns.phonePattern, '')
        .replace(/[,;]+/g, ' ')
        .trim();

      // Try to split by common patterns
      const parts = description.split(/[,;]\s*/);
      
      if (parts.length >= 2) {
        entry['Item'] = parts[0].trim();
        
        // Look for quantity pattern
        const qtyPart = parts.find(p => /\d+\s*(unit|pcs|qty|buah|item)/i.test(p));
        if (qtyPart) {
          const qtyMatch = qtyPart.match(/(\d+)/);
          if (qtyMatch) entry['Quantity'] = parseInt(qtyMatch[1]);
        }
        
        // Remaining as description
        const remaining = parts.filter(p => p !== parts[0] && p !== qtyPart);
        if (remaining.length > 0) {
          entry['Description'] = remaining.join(', ');
        }
      } else if (description) {
        entry['Description'] = description;
      }

      // Only add if has meaningful data
      if (Object.keys(entry).length > 1) {
        extracted.push(entry);
      }
    });

    // Determine headers from extracted data
    const allKeys = new Set();
    extracted.forEach(row => {
      Object.keys(row).forEach(k => {
        if (k !== '_rowIndex') allKeys.add(k);
      });
    });

    const headers = Array.from(allKeys);
    
    // Normalize data to have all headers
    const data = extracted.map(row => {
      headers.forEach(h => {
        if (!(h in row)) row[h] = '';
      });
      return row;
    });

    return { headers, data };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // GENERIC PARSING (Fallback)
  // ─────────────────────────────────────────────────────────────────────────

  _parseGeneric(lines, options) {
    // Treat each line as a single column entry
    const headers = ['No', 'Content'];
    const data = lines.map((line, index) => ({
      _rowIndex: index + 2,
      No: index + 1,
      Content: line.trim(),
    }));

    return { headers, data };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // TYPE CONVERSION
  // ─────────────────────────────────────────────────────────────────────────

  _autoConvertTypes(headers, data) {
    headers.forEach(header => {
      const values = data.map(row => row[header]).filter(v => v !== '');
      
      if (values.length === 0) return;

      // Check if all values are numbers
      const allNumbers = values.every(v => !isNaN(helpers.parseNumber(v)));
      if (allNumbers) {
        data.forEach(row => {
          if (row[header] !== '') {
            row[header] = helpers.parseNumber(row[header]);
          }
        });
        return;
      }

      // Check if all values are dates
      const allDates = values.every(v => helpers.parseDate(v) !== null);
      if (allDates) {
        data.forEach(row => {
          if (row[header] !== '') {
            row[header] = helpers.formatDate(row[header]);
          }
        });
      }
    });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // STRUCTURED TEXT PARSING (for specific formats)
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Parse structured text like:
   * "Produk A, 10 unit, harga 50000, tanggal 15 Januari 2024"
   */
  parseStructuredLines(text, options = {}) {
    const lines = text.trim().split(/\r?\n/).filter(l => l.trim());
    
    if (lines.length === 0) {
      return { success: false, error: 'No data found' };
    }

    const data = [];
    const foundHeaders = new Set();

    lines.forEach((line, index) => {
      const entry = { _rowIndex: index + 2 };
      
      // Split by comma
      const parts = line.split(/[,;]+/).map(p => p.trim());
      
      parts.forEach(part => {
        // Try to identify each part
        
        // Quantity pattern: "10 unit", "5 pcs", "20 buah"
        const qtyMatch = part.match(/^(\d+)\s*(unit|pcs|qty|buah|item|box|pack)?$/i);
        if (qtyMatch) {
          entry['Quantity'] = parseInt(qtyMatch[1]);
          foundHeaders.add('Quantity');
          return;
        }

        // Currency pattern: "harga 50000", "Rp 50.000"
        const priceMatch = part.match(/(?:harga|price)?\s*(?:Rp\.?\s?)?([\d.,]+)/i);
        if (priceMatch && !entry['Price']) {
          const num = helpers.parseNumber(priceMatch[1]);
          if (num > 0) {
            entry['Price'] = num;
            foundHeaders.add('Price');
            return;
          }
        }

        // Date pattern
        const dateMatch = part.match(/(?:tanggal|date|tgl)?\s*(\d{1,2}[-\/\s]\w+[-\/\s]\d{2,4}|\d{1,2}\s+\w+\s+\d{4})/i);
        if (dateMatch) {
          entry['Date'] = helpers.formatDate(dateMatch[1]) || dateMatch[1];
          foundHeaders.add('Date');
          return;
        }

        // Otherwise it's probably a name/description
        if (!entry['Item'] && part.length > 0) {
          entry['Item'] = part;
          foundHeaders.add('Item');
        } else if (!entry['Description'] && part.length > 0) {
          entry['Description'] = part;
          foundHeaders.add('Description');
        }
      });

      data.push(entry);
    });

    // Order headers logically
    const headerOrder = ['Item', 'Description', 'Quantity', 'Price', 'Date'];
    const headers = headerOrder.filter(h => foundHeaders.has(h));
    
    // Add any other found headers
    foundHeaders.forEach(h => {
      if (!headers.includes(h)) headers.push(h);
    });

    // Calculate totals if we have quantity and price
    if (foundHeaders.has('Quantity') && foundHeaders.has('Price')) {
      headers.push('Total');
      data.forEach(row => {
        if (row.Quantity && row.Price) {
          row.Total = row.Quantity * row.Price;
        }
      });
    }

    // Normalize data
    data.forEach(row => {
      headers.forEach(h => {
        if (!(h in row)) row[h] = '';
      });
    });

    return {
      success: true,
      headers,
      data,
      format: 'structured',
    };
  }
}

module.exports = new TextToExcel();
