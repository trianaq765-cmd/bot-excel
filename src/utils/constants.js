// ═══════════════════════════════════════════════════════════════════════════
// CONSTANTS - Semua konstanta yang digunakan di seluruh aplikasi
// ═══════════════════════════════════════════════════════════════════════════

module.exports = {
  // ─────────────────────────────────────────────────────────────────────────
  // FILE SETTINGS
  // ─────────────────────────────────────────────────────────────────────────
  FILE: {
    MAX_SIZE: 10 * 1024 * 1024, // 10MB
    ALLOWED_EXTENSIONS: ['.xlsx', '.xls', '.csv'],
    TEMP_TTL: 60 * 60 * 1000, // 1 hour
  },

  // ─────────────────────────────────────────────────────────────────────────
  // ANALYSIS MODES
  // ─────────────────────────────────────────────────────────────────────────
  ANALYSIS_MODES: {
    AUTO: 'auto',           // Detect tipe data otomatis
    FINANCE: 'finance',     // Fokus ke financial rules
    SALES: 'sales',         // Fokus ke sales/marketing rules
    DATA: 'data',           // General data analysis
    STRICT: 'strict',       // Maximum detection
  },

  // ─────────────────────────────────────────────────────────────────────────
  // ISSUE SEVERITY
  // ─────────────────────────────────────────────────────────────────────────
  SEVERITY: {
    AUTO_FIX: 'auto_fix',       // 🟢 Bisa diperbaiki otomatis
    NEEDS_REVIEW: 'needs_review', // 🟡 Butuh review manual
    CRITICAL: 'critical',        // 🔴 Tidak bisa auto-fix
  },

  // ─────────────────────────────────────────────────────────────────────────
  // ISSUE TYPES
  // ─────────────────────────────────────────────────────────────────────────
  ISSUE_TYPES: {
    // Format Issues
    DATE_INCONSISTENT: 'date_inconsistent',
    NUMBER_FORMAT: 'number_format',
    CURRENCY_FORMAT: 'currency_format',
    PHONE_FORMAT: 'phone_format',
    EMAIL_FORMAT: 'email_format',
    TEXT_CASE: 'text_case',
    WHITESPACE: 'whitespace',

    // Data Quality Issues
    DUPLICATE_ROW: 'duplicate_row',
    DUPLICATE_FUZZY: 'duplicate_fuzzy',
    EMPTY_ROW: 'empty_row',
    EMPTY_CELL: 'empty_cell',
    MISSING_REQUIRED: 'missing_required',
    
    // Outliers & Anomalies
    NUMERIC_OUTLIER: 'numeric_outlier',
    NEGATIVE_INVALID: 'negative_invalid',
    FUTURE_DATE: 'future_date',
    PAST_DATE_INVALID: 'past_date_invalid',
    
    // Logic Issues
    CALCULATION_ERROR: 'calculation_error',
    SEQUENCE_GAP: 'sequence_gap',
    SEQUENCE_ORDER: 'sequence_order',
    
    // Indonesia Specific
    NIK_INVALID: 'nik_invalid',
    NPWP_INVALID: 'npwp_invalid',
    TAX_CALCULATION: 'tax_calculation',
    
    // Structure Issues
    NO_HEADER: 'no_header',
    DUPLICATE_HEADER: 'duplicate_header',
    MIXED_DATA_TYPE: 'mixed_data_type',
  },

  // ─────────────────────────────────────────────────────────────────────────
  // INDONESIA SPECIFIC
  // ─────────────────────────────────────────────────────────────────────────
  INDONESIA: {
    // Kode Provinsi (sample)
    PROVINCE_CODES: {
      '11': 'Aceh', '12': 'Sumatera Utara', '13': 'Sumatera Barat',
      '14': 'Riau', '15': 'Jambi', '16': 'Sumatera Selatan',
      '17': 'Bengkulu', '18': 'Lampung', '19': 'Bangka Belitung',
      '21': 'Kepulauan Riau', '31': 'DKI Jakarta', '32': 'Jawa Barat',
      '33': 'Jawa Tengah', '34': 'DI Yogyakarta', '35': 'Jawa Timur',
      '36': 'Banten', '51': 'Bali', '52': 'NTB', '53': 'NTT',
      '61': 'Kalimantan Barat', '62': 'Kalimantan Tengah',
      '63': 'Kalimantan Selatan', '64': 'Kalimantan Timur',
      '65': 'Kalimantan Utara', '71': 'Sulawesi Utara',
      '72': 'Sulawesi Tengah', '73': 'Sulawesi Selatan',
      '74': 'Sulawesi Tenggara', '75': 'Gorontalo',
      '76': 'Sulawesi Barat', '81': 'Maluku', '82': 'Maluku Utara',
      '91': 'Papua', '92': 'Papua Barat',
    },
    
    // Tax rates
    PPN_RATE: 0.11,           // 11%
    PPH21_BRACKETS: [
      { min: 0, max: 60000000, rate: 0.05 },
      { min: 60000000, max: 250000000, rate: 0.15 },
      { min: 250000000, max: 500000000, rate: 0.25 },
      { min: 500000000, max: 5000000000, rate: 0.30 },
      { min: 5000000000, max: Infinity, rate: 0.35 },
    ],
    
    // PTKP 2024
    PTKP: {
      'TK/0': 54000000, 'TK/1': 58500000, 'TK/2': 63000000, 'TK/3': 67500000,
      'K/0': 58500000, 'K/1': 63000000, 'K/2': 67500000, 'K/3': 72000000,
    },
    
    // Bank Account Length
    BANK_ACCOUNT_LENGTH: {
      'BCA': 10, 'MANDIRI': 13, 'BNI': 10, 'BRI': 15, 'CIMB': 13,
    },
  },

  // ─────────────────────────────────────────────────────────────────────────
  // FORMAT PATTERNS (Regex)
  // ─────────────────────────────────────────────────────────────────────────
  PATTERNS: {
    EMAIL: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
    PHONE_ID: /^(\+62|62|0)[\s.-]?(\d{2,4})[\s.-]?(\d{3,4})[\s.-]?(\d{3,4})$/,
    NIK: /^\d{16}$/,
    NPWP: /^\d{15}$/,
    CURRENCY_ID: /^[Rr][Pp]\.?\s?[\d.,]+$/,
    DATE_ISO: /^\d{4}-\d{2}-\d{2}$/,
    DATE_DMY: /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/,
    DATE_MDY: /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/,
  },

  // ─────────────────────────────────────────────────────────────────────────
  // PROFESSIONAL FORMATTING
  // ─────────────────────────────────────────────────────────────────────────
  FORMATTING: {
    HEADER_COLOR: 'FF2B579A',     // Dark Blue
    HEADER_FONT_COLOR: 'FFFFFFFF', // White
    ALT_ROW_COLOR: 'FFF5F5F5',    // Light Gray
    ERROR_COLOR: 'FFFFC7CE',       // Light Red
    WARNING_COLOR: 'FFFFEB9C',     // Light Yellow
    SUCCESS_COLOR: 'FFC6EFCE',     // Light Green
    BORDER_COLOR: 'FFD4D4D4',      // Gray
    FONT_NAME: 'Calibri',
    FONT_SIZE: 11,
    HEADER_FONT_SIZE: 12,
  },

  // ─────────────────────────────────────────────────────────────────────────
  // COLUMN TYPE DETECTION
  // ─────────────────────────────────────────────────────────────────────────
  COLUMN_TYPES: {
    STRING: 'string',
    NUMBER: 'number',
    CURRENCY: 'currency',
    DATE: 'date',
    EMAIL: 'email',
    PHONE: 'phone',
    NIK: 'nik',
    NPWP: 'npwp',
    PERCENTAGE: 'percentage',
    BOOLEAN: 'boolean',
    MIXED: 'mixed',
    EMPTY: 'empty',
  },
};
