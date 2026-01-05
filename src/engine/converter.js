// ═══════════════════════════════════════════════════════════════════════════
// CONVERTER ENGINE - Format conversion utilities
// ═══════════════════════════════════════════════════════════════════════════

const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const helpers = require('../utils/helpers');

class Converter {
  // ─────────────────────────────────────────────────────────────────────────
  // EXCEL TO OTHER FORMATS
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Convert data to CSV string
   */
  toCSV(headers, data, options = {}) {
    const delimiter = options.delimiter || ',';
    const lineEnding = options.lineEnding || '\n';
    
    const rows = [headers.join(delimiter)];
    
    data.forEach(row => {
      const values = headers.map(h => {
        let val = row[h];
        if (val === null || val === undefined) return '';
        val = String(val);
        // Escape quotes and wrap if contains delimiter
        if (val.includes(delimiter) || val.includes('"') || val.includes('\n')) {
          val = '"' + val.replace(/"/g, '""') + '"';
        }
        return val;
      });
      rows.push(values.join(delimiter));
    });
    
    return rows.join(lineEnding);
  }

  /**
   * Convert data to JSON
   */
  toJSON(headers, data, options = {}) {
    const formatted = data.map(row => {
      const obj = {};
      headers.forEach(h => {
        obj[h] = row[h];
      });
      // Remove internal properties
      delete obj._rowIndex;
      return obj;
    });
    
    return options.pretty 
      ? JSON.stringify(formatted, null, 2)
      : JSON.stringify(formatted);
  }

  /**
   * Convert data to HTML table
   */
  toHTML(headers, data, options = {}) {
    const title = options.title || 'Data Export';
    const theme = options.theme || 'default';
    
    const styles = this._getHTMLStyles(theme);
    
    let html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${title}</title>
  <style>${styles}</style>
</head>
<body>
  <h1>${title}</h1>
  <p>Generated: ${new Date().toLocaleString()}</p>
  <p>Total Rows: ${data.length}</p>
  
  <table>
    <thead>
      <tr>
        ${headers.map(h => `<th>${this._escapeHTML(h)}</th>`).join('')}
      </tr>
    </thead>
    <tbody>
      ${data.map((row, i) => `
      <tr class="${i % 2 === 0 ? 'even' : 'odd'}">
        ${headers.map(h => `<td>${this._escapeHTML(String(row[h] || ''))}</td>`).join('')}
      </tr>`).join('')}
    </tbody>
  </table>
</body>
</html>`;
    
    return html;
  }

  /**
   * Convert data to Markdown table
   */
  toMarkdown(headers, data, options = {}) {
    let md = '';
    
    if (options.title) {
      md += `# ${options.title}\n\n`;
    }
    
    // Header row
    md += '| ' + headers.join(' | ') + ' |\n';
    
    // Separator
    md += '| ' + headers.map(() => '---').join(' | ') + ' |\n';
    
    // Data rows
    data.forEach(row => {
      const values = headers.map(h => {
        const val = row[h];
        if (val === null || val === undefined) return '';
        return String(val).replace(/\|/g, '\\|');
      });
      md += '| ' + values.join(' | ') + ' |\n';
    });
    
    return md;
  }

  /**
   * Convert data to SQL INSERT statements
   */
  toSQL(headers, data, options = {}) {
    const tableName = options.tableName || 'data_table';
    const batchSize = options.batchSize || 100;
    
    const statements = [];
    
    // Create table statement (optional)
    if (options.includeCreate) {
      const columns = headers.map(h => {
        const safeName = h.replace(/[^a-zA-Z0-9_]/g, '_');
        return `  ${safeName} VARCHAR(255)`;
      }).join(',\n');
      
      statements.push(`CREATE TABLE IF NOT EXISTS ${tableName} (\n${columns}\n);`);
    }
    
    // Insert statements
    const safeHeaders = headers.map(h => h.replace(/[^a-zA-Z0-9_]/g, '_'));
    const columnList = safeHeaders.join(', ');
    
    for (let i = 0; i < data.length; i += batchSize) {
      const batch = data.slice(i, i + batchSize);
      
      const values = batch.map(row => {
        const rowValues = headers.map(h => {
          const val = row[h];
          if (val === null || val === undefined) return 'NULL';
          if (typeof val === 'number') return val;
          return `'${String(val).replace(/'/g, "''")}'`;
        });
        return `(${rowValues.join(', ')})`;
      }).join(',\n  ');
      
      statements.push(`INSERT INTO ${tableName} (${columnList})\nVALUES\n  ${values};`);
    }
    
    return statements.join('\n\n');
  }

  /**
   * Convert data to XML
   */
  toXML(headers, data, options = {}) {
    const rootElement = options.rootElement || 'data';
    const rowElement = options.rowElement || 'row';
    
    let xml = '<?xml version="1.0" encoding="UTF-8"?>\n';
    xml += `<${rootElement}>\n`;
    
    data.forEach(row => {
      xml += `  <${rowElement}>\n`;
      headers.forEach(h => {
        const safeName = h.replace(/[^a-zA-Z0-9_]/g, '_');
        const val = this._escapeXML(String(row[h] || ''));
        xml += `    <${safeName}>${val}</${safeName}>\n`;
      });
      xml += `  </${rowElement}>\n`;
    });
    
    xml += `</${rootElement}>`;
    
    return xml;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // OTHER FORMATS TO EXCEL
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Convert CSV string to data object
   */
  fromCSV(csvString, options = {}) {
    const delimiter = options.delimiter || ',';
    const lines = csvString.split(/\r?\n/).filter(line => line.trim());
    
    if (lines.length === 0) {
      return { headers: [], data: [] };
    }
    
    const headers = this._parseCSVLine(lines[0], delimiter);
    const data = [];
    
    for (let i = 1; i < lines.length; i++) {
      const values = this._parseCSVLine(lines[i], delimiter);
      const row = { _rowIndex: i + 1 };
      headers.forEach((h, idx) => {
        row[h] = values[idx] || '';
      });
      data.push(row);
    }
    
    return { headers, data };
  }

  /**
   * Convert JSON string/array to data object
   */
  fromJSON(jsonData) {
    let parsed;
    
    if (typeof jsonData === 'string') {
      parsed = JSON.parse(jsonData);
    } else {
      parsed = jsonData;
    }
    
    if (!Array.isArray(parsed)) {
      parsed = [parsed];
    }
    
    if (parsed.length === 0) {
      return { headers: [], data: [] };
    }
    
    // Extract headers from first object
    const headers = Object.keys(parsed[0]);
    
    // Map data
    const data = parsed.map((item, index) => {
      const row = { _rowIndex: index + 2 };
      headers.forEach(h => {
        row[h] = item[h];
      });
      return row;
    });
    
    return { headers, data };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HELPER METHODS
  // ─────────────────────────────────────────────────────────────────────────

  _parseCSVLine(line, delimiter) {
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

  _escapeHTML(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  _escapeXML(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  _getHTMLStyles(theme) {
    const themes = {
      default: `
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        h1 { color: #2B579A; }
        table { border-collapse: collapse; width: 100%; background: white; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        th { background: #2B579A; color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr.even { background: #f9f9f9; }
        tr:hover { background: #e8f0fe; }
      `,
      dark: `
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #1a1a2e; color: #eee; }
        h1 { color: #00d2ff; }
        table { border-collapse: collapse; width: 100%; background: #16213e; }
        th { background: #0f3460; color: #00d2ff; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #0f3460; }
        tr.even { background: #1a1a2e; }
        tr:hover { background: #0f3460; }
      `,
      minimal: `
        body { font-family: Arial, sans-serif; margin: 20px; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background: #f2f2f2; }
      `,
    };
    
    return themes[theme] || themes.default;
  }
}

module.exports = new Converter();
