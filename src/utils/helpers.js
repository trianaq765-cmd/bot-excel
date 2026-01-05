// ═══════════════════════════════════════════════════════════════════════════
// HELPERS - Utility functions used across the application
// ═══════════════════════════════════════════════════════════════════════════

const { PATTERNS, INDONESIA } = require('./constants');

module.exports = {
  // ─────────────────────────────────────────────────────────────────────────
  // STRING HELPERS
  // ─────────────────────────────────────────────────────────────────────────
  
  /**
   * Trim dan remove multiple spaces
   */
  cleanString(str) {
    if (typeof str !== 'string') return str;
    return str.trim().replace(/\s+/g, ' ');
  },

  /**
   * Convert to Title Case
   */
  toTitleCase(str) {
    if (typeof str !== 'string') return str;
    return str.toLowerCase().replace(/(?:^|\s)\w/g, (match) => match.toUpperCase());
  },

  /**
   * Convert to Sentence case
   */
  toSentenceCase(str) {
    if (typeof str !== 'string') return str;
    return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
  },

  /**
   * Remove special characters
   */
  removeSpecialChars(str, keepSpaces = true) {
    if (typeof str !== 'string') return str;
    const pattern = keepSpaces ? /[^a-zA-Z0-9\s]/g : /[^a-zA-Z0-9]/g;
    return str.replace(pattern, '');
  },

  // ─────────────────────────────────────────────────────────────────────────
  // NUMBER HELPERS
  // ─────────────────────────────────────────────────────────────────────────
  
  /**
   * Parse number from various formats
   */
  parseNumber(value) {
    if (typeof value === 'number') return value;
    if (typeof value !== 'string') return NaN;
    
    // Remove currency symbols and spaces
    let cleaned = value.replace(/[Rr][Pp]\.?\s?/g, '')
                       .replace(/[IDR\s]/g, '')
                       .trim();
    
    // Handle Indonesian format (1.234.567,89)
    if (/^\d{1,3}(\.\d{3})*,\d+$/.test(cleaned)) {
      cleaned = cleaned.replace(/\./g, '').replace(',', '.');
    }
    // Handle US format (1,234,567.89)
    else if (/^\d{1,3}(,\d{3})*\.\d+$/.test(cleaned)) {
      cleaned = cleaned.replace(/,/g, '');
    }
    // Handle simple comma as decimal
    else if (/^\d+,\d+$/.test(cleaned)) {
      cleaned = cleaned.replace(',', '.');
    }
    // Remove any remaining commas/dots except last one for decimal
    else {
      cleaned = cleaned.replace(/,/g, '');
    }
    
    return parseFloat(cleaned);
  },

  /**
   * Format number as Indonesian currency
   */
  formatCurrency(value, withSymbol = true) {
    const num = typeof value === 'number' ? value : this.parseNumber(value);
    if (isNaN(num)) return value;
    
    const formatted = new Intl.NumberFormat('id-ID').format(Math.round(num));
    return withSymbol ? `Rp ${formatted}` : formatted;
  },

  /**
   * Format number with thousand separator
   */
  formatNumber(value, decimals = 0) {
    const num = typeof value === 'number' ? value : this.parseNumber(value);
    if (isNaN(num)) return value;
    
    return new Intl.NumberFormat('id-ID', {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    }).format(num);
  },

  /**
   * Check if number is outlier using IQR method
   */
  isOutlier(value, values) {
    const sorted = values.filter(v => !isNaN(v)).sort((a, b) => a - b);
    const q1 = sorted[Math.floor(sorted.length * 0.25)];
    const q3 = sorted[Math.floor(sorted.length * 0.75)];
    const iqr = q3 - q1;
    const lowerBound = q1 - 1.5 * iqr;
    const upperBound = q3 + 1.5 * iqr;
    
    return value < lowerBound || value > upperBound;
  },

  // ─────────────────────────────────────────────────────────────────────────
  // DATE HELPERS
  // ─────────────────────────────────────────────────────────────────────────
  
  /**
   * Parse date from various formats
   */
  parseDate(value) {
    if (value instanceof Date) return value;
    if (typeof value === 'number') {
      // Excel serial date
      return new Date((value - 25569) * 86400 * 1000);
    }
    if (typeof value !== 'string') return null;
    
    const str = value.trim();
    
    // Try various formats
    const formats = [
      // ISO
      { regex: /^(\d{4})-(\d{2})-(\d{2})$/, parse: (m) => new Date(m[1], m[2] - 1, m[3]) },
      // DD/MM/YYYY or DD-MM-YYYY
      { regex: /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/, parse: (m) => new Date(m[3], m[2] - 1, m[1]) },
      // DD/MM/YY
      { regex: /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})$/, parse: (m) => new Date(2000 + parseInt(m[3]), m[2] - 1, m[1]) },
      // DD MMM YYYY
      { regex: /^(\d{1,2})\s+(\w{3,})\s+(\d{4})$/i, parse: (m) => new Date(`${m[2]} ${m[1]}, ${m[3]}`) },
    ];
    
    for (const fmt of formats) {
      const match = str.match(fmt.regex);
      if (match) {
        const date = fmt.parse(match);
        if (!isNaN(date.getTime())) return date;
      }
    }
    
    // Fallback to native parsing
    const date = new Date(str);
    return isNaN(date.getTime()) ? null : date;
  },

  /**
   * Format date to standard format
   */
  formatDate(value, format = 'DD-MMM-YYYY') {
    const date = this.parseDate(value);
    if (!date) return value;
    
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    
    const dd = String(date.getDate()).padStart(2, '0');
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const mmm = months[date.getMonth()];
    const yyyy = date.getFullYear();
    
    switch (format) {
      case 'DD-MMM-YYYY': return `${dd}-${mmm}-${yyyy}`;
      case 'DD/MM/YYYY': return `${dd}/${mm}/${yyyy}`;
      case 'YYYY-MM-DD': return `${yyyy}-${mm}-${dd}`;
      case 'DD MMMM YYYY': 
        const fullMonths = ['January', 'February', 'March', 'April', 'May', 'June',
                           'July', 'August', 'September', 'October', 'November', 'December'];
        return `${dd} ${fullMonths[date.getMonth()]} ${yyyy}`;
      default: return `${dd}-${mmm}-${yyyy}`;
    }
  },

  /**
   * Check if date is in future
   */
  isFutureDate(value) {
    const date = this.parseDate(value);
    if (!date) return false;
    return date > new Date();
  },

  // ─────────────────────────────────────────────────────────────────────────
  // PHONE HELPERS
  // ─────────────────────────────────────────────────────────────────────────
  
  /**
   * Format phone number to Indonesian standard
   */
  formatPhoneID(value) {
    if (typeof value !== 'string') value = String(value);
    
    // Remove all non-digits
    let digits = value.replace(/\D/g, '');
    
    // Handle different prefixes
    if (digits.startsWith('62')) {
      digits = '0' + digits.slice(2);
    } else if (digits.startsWith('8')) {
      digits = '0' + digits;
    }
    
    // Format: 0812-3456-7890
    if (digits.length >= 10 && digits.length <= 13) {
      const formatted = digits.replace(/(\d{4})(\d{4})(\d+)/, '$1-$2-$3');
      return '+62 ' + formatted.slice(1);
    }
    
    return value;
  },

  /**
   * Validate Indonesian phone number
   */
  isValidPhoneID(value) {
    if (typeof value !== 'string') value = String(value);
    const digits = value.replace(/\D/g, '');
    
    // Check length (10-13 digits for Indonesian mobile)
    if (digits.length < 10 || digits.length > 13) return false;
    
    // Check prefix (valid Indonesian mobile prefixes)
    const validPrefixes = ['08', '628', '8'];
    return validPrefixes.some(prefix => digits.startsWith(prefix) || 
                             digits.startsWith(prefix.replace('0', '')));
  },

  // ─────────────────────────────────────────────────────────────────────────
  // EMAIL HELPERS
  // ─────────────────────────────────────────────────────────────────────────
  
  /**
   * Validate email format
   */
  isValidEmail(value) {
    if (typeof value !== 'string') return false;
    return PATTERNS.EMAIL.test(value.trim().toLowerCase());
  },

  /**
   * Fix common email issues
   */
  fixEmail(value) {
    if (typeof value !== 'string') return value;
    
    let email = value.trim().toLowerCase();
    
    // Remove trailing dots
    email = email.replace(/\.+$/, '');
    
    // Remove double @
    email = email.replace(/@@+/g, '@');
    
    // Remove spaces
    email = email.replace(/\s+/g, '');
    
    return email;
  },

  // ─────────────────────────────────────────────────────────────────────────
  // INDONESIA SPECIFIC HELPERS
  // ─────────────────────────────────────────────────────────────────────────
  
  /**
   * Validate NIK (KTP Indonesia)
   */
  validateNIK(nik) {
    if (typeof nik !== 'string') nik = String(nik);
    const cleaned = nik.replace(/\D/g, '');
    
    const result = {
      isValid: false,
      errors: [],
      parsed: null,
    };
    
    // Check length
    if (cleaned.length !== 16) {
      result.errors.push(`Invalid length: ${cleaned.length} (should be 16)`);
      return result;
    }
    
    // Parse components
    const provinceCode = cleaned.substring(0, 2);
    const regencyCode = cleaned.substring(2, 4);
    const districtCode = cleaned.substring(4, 6);
    const birthDate = cleaned.substring(6, 12);
    const sequence = cleaned.substring(12, 16);
    
    // Validate province code
    if (!INDONESIA.PROVINCE_CODES[provinceCode]) {
      result.errors.push(`Invalid province code: ${provinceCode}`);
    }
    
    // Parse birth date
    let day = parseInt(birthDate.substring(0, 2));
    const month = parseInt(birthDate.substring(2, 4));
    const year = parseInt(birthDate.substring(4, 6));
    
    // For females, day is added by 40
    const isFemale = day > 40;
    if (isFemale) day -= 40;
    
    // Validate date
    if (month < 1 || month > 12 || day < 1 || day > 31) {
      result.errors.push(`Invalid birth date in NIK`);
    }
    
    if (result.errors.length === 0) {
      result.isValid = true;
      result.parsed = {
        province: INDONESIA.PROVINCE_CODES[provinceCode] || 'Unknown',
        provinceCode,
        regencyCode,
        districtCode,
        birthDate: `${day.toString().padStart(2, '0')}/${month.toString().padStart(2, '0')}/${year < 50 ? 2000 + year : 1900 + year}`,
        gender: isFemale ? 'Female' : 'Male',
        sequence,
      };
    }
    
    return result;
  },

  /**
   * Format NIK with proper spacing
   */
  formatNIK(nik) {
    if (typeof nik !== 'string') nik = String(nik);
    const cleaned = nik.replace(/\D/g, '');
    if (cleaned.length !== 16) return nik;
    
    return cleaned.replace(/(\d{6})(\d{6})(\d{4})/, '$1.$2.$3');
  },

  /**
   * Validate NPWP
   */
  validateNPWP(npwp) {
    if (typeof npwp !== 'string') npwp = String(npwp);
    const cleaned = npwp.replace(/\D/g, '');
    
    const result = {
      isValid: false,
      errors: [],
      formatted: null,
    };
    
    if (cleaned.length !== 15) {
      result.errors.push(`Invalid length: ${cleaned.length} (should be 15)`);
      return result;
    }
    
    result.isValid = true;
    result.formatted = cleaned.replace(/(\d{2})(\d{3})(\d{3})(\d{1})(\d{3})(\d{3})/, 
                                        '$1.$2.$3.$4-$5.$6');
    
    return result;
  },

  /**
   * Format NPWP to standard format
   */
  formatNPWP(npwp) {
    if (typeof npwp !== 'string') npwp = String(npwp);
    const cleaned = npwp.replace(/\D/g, '');
    if (cleaned.length !== 15) return npwp;
    
    return cleaned.replace(/(\d{2})(\d{3})(\d{3})(\d{1})(\d{3})(\d{3})/, 
                          '$1.$2.$3.$4-$5.$6');
  },

  // ─────────────────────────────────────────────────────────────────────────
  // STATISTICS HELPERS
  // ─────────────────────────────────────────────────────────────────────────
  
  /**
   * Calculate basic statistics for an array of numbers
   */
  calculateStats(values) {
    const numbers = values.map(v => this.parseNumber(v)).filter(v => !isNaN(v));
    
    if (numbers.length === 0) {
      return {
        count: 0,
        sum: 0,
        average: 0,
        min: 0,
        max: 0,
        median: 0,
        stdDev: 0,
      };
    }
    
    const sorted = [...numbers].sort((a, b) => a - b);
    const sum = numbers.reduce((a, b) => a + b, 0);
    const avg = sum / numbers.length;
    
    // Standard deviation
    const squaredDiffs = numbers.map(v => Math.pow(v - avg, 2));
    const avgSquaredDiff = squaredDiffs.reduce((a, b) => a + b, 0) / numbers.length;
    const stdDev = Math.sqrt(avgSquaredDiff);
    
    // Median
    const mid = Math.floor(sorted.length / 2);
    const median = sorted.length % 2 !== 0
      ? sorted[mid]
      : (sorted[mid - 1] + sorted[mid]) / 2;
    
    return {
      count: numbers.length,
      sum,
      average: avg,
      min: sorted[0],
      max: sorted[sorted.length - 1],
      median,
      stdDev,
    };
  },

  // ─────────────────────────────────────────────────────────────────────────
  // SIMILARITY HELPERS (for fuzzy duplicate detection)
  // ─────────────────────────────────────────────────────────────────────────
  
  /**
   * Calculate Levenshtein distance between two strings
   */
  levenshteinDistance(str1, str2) {
    const m = str1.length;
    const n = str2.length;
    const dp = Array(m + 1).fill(null).map(() => Array(n + 1).fill(0));
    
    for (let i = 0; i <= m; i++) dp[i][0] = i;
    for (let j = 0; j <= n; j++) dp[0][j] = j;
    
    for (let i = 1; i <= m; i++) {
      for (let j = 1; j <= n; j++) {
        if (str1[i - 1] === str2[j - 1]) {
          dp[i][j] = dp[i - 1][j - 1];
        } else {
          dp[i][j] = 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
        }
      }
    }
    
    return dp[m][n];
  },

  /**
   * Calculate similarity score between two strings (0-1)
   */
  stringSimilarity(str1, str2) {
    if (!str1 || !str2) return 0;
    
    const s1 = String(str1).toLowerCase().trim();
    const s2 = String(str2).toLowerCase().trim();
    
    if (s1 === s2) return 1;
    
    const distance = this.levenshteinDistance(s1, s2);
    const maxLength = Math.max(s1.length, s2.length);
    
    return 1 - (distance / maxLength);
  },

  // ─────────────────────────────────────────────────────────────────────────
  // UTILITY HELPERS
  // ─────────────────────────────────────────────────────────────────────────
  
  /**
   * Generate unique ID
   */
  generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
  },

  /**
   * Deep clone object
   */
  deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
  },

  /**
   * Check if value is empty
   */
  isEmpty(value) {
    if (value === null || value === undefined) return true;
    if (typeof value === 'string') return value.trim() === '';
    if (Array.isArray(value)) return value.length === 0;
    if (typeof value === 'object') return Object.keys(value).length === 0;
    return false;
  },

  /**
   * Format bytes to human readable
   */
  formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  },

  /**
   * Format duration in milliseconds to human readable
   */
  formatDuration(ms) {
    if (ms < 1000) return `${ms}ms`;
    if (ms < 60000) return `${(ms / 1000).toFixed(1)}s`;
    return `${Math.floor(ms / 60000)}m ${Math.floor((ms % 60000) / 1000)}s`;
  },
};
