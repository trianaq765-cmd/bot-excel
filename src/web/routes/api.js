// ═══════════════════════════════════════════════════════════════════════════
// API ROUTES - Express API endpoints with proper error handling
// ═══════════════════════════════════════════════════════════════════════════

const express = require('express');
const multer = require('multer');
const path = require('path');
const { v4: uuidv4 } = require('uuid');

const router = express.Router();

// Lazy load engine to prevent startup errors
let engine = null;
function getEngine() {
  if (!engine) {
    try {
      engine = require('../../engine');
      console.log('✅ Engine loaded successfully');
    } catch (error) {
      console.error('❌ Failed to load engine:', error.message);
      throw error;
    }
  }
  return engine;
}

// ─────────────────────────────────────────────────────────────────────────
// MULTER CONFIGURATION
// ─────────────────────────────────────────────────────────────────────────

const storage = multer.memoryStorage();

const fileFilter = (req, file, cb) => {
  const allowedTypes = ['.xlsx', '.xls', '.csv'];
  const ext = path.extname(file.originalname).toLowerCase();
  
  if (allowedTypes.includes(ext)) {
    cb(null, true);
  } else {
    cb(new Error(`Invalid file type. Allowed: ${allowedTypes.join(', ')}`), false);
  }
};

const upload = multer({
  storage,
  fileFilter,
  limits: { fileSize: 10 * 1024 * 1024 },
});

// ─────────────────────────────────────────────────────────────────────────
// FILE STORAGE
// ─────────────────────────────────────────────────────────────────────────

const fileStore = new Map();

setInterval(() => {
  const now = Date.now();
  const maxAge = 60 * 60 * 1000;
  
  for (const [id, data] of fileStore.entries()) {
    if (now - data.timestamp > maxAge) {
      fileStore.delete(id);
    }
  }
}, 60 * 60 * 1000);

// ─────────────────────────────────────────────────────────────────────────
// HELPER: Wrap async route handlers
// ─────────────────────────────────────────────────────────────────────────

const asyncHandler = (fn) => (req, res, next) => {
  Promise.resolve(fn(req, res, next)).catch(next);
};

// ─────────────────────────────────────────────────────────────────────────
// ROUTES
// ─────────────────────────────────────────────────────────────────────────

/**
 * GET /api/health
 */
router.get('/health', (req, res) => {
  res.json({
    success: true,
    status: 'healthy',
    timestamp: new Date().toISOString(),
    filesInMemory: fileStore.size,
  });
});

/**
 * POST /api/upload
 */
router.post('/upload', upload.single('file'), asyncHandler(async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ success: false, error: 'No file uploaded' });
  }

  console.log(`[API] Upload: ${req.file.originalname} (${req.file.size} bytes)`);

  const eng = getEngine();
  const fileId = uuidv4();
  const parsed = eng.fileParser.parse(req.file.buffer, {
    fileName: req.file.originalname,
  });

  if (!parsed.success) {
    return res.status(400).json({ success: false, error: parsed.error });
  }

  fileStore.set(fileId, {
    buffer: req.file.buffer,
    fileName: req.file.originalname,
    parsed,
    timestamp: Date.now(),
  });

  console.log(`[API] File stored: ${fileId}`);

  res.json({
    success: true,
    fileId,
    fileName: req.file.originalname,
    fileSize: req.file.size,
    preview: {
      headers: parsed.headers,
      data: parsed.data.slice(0, 10),
      rowCount: parsed.rowCount,
      columnCount: parsed.columnCount,
    },
    columnTypes: parsed.columnTypes,
  });
}));

/**
 * POST /api/analyze
 */
router.post('/analyze', asyncHandler(async (req, res) => {
  const { fileId, mode = 'auto' } = req.body;

  if (!fileId || !fileStore.has(fileId)) {
    return res.status(400).json({ success: false, error: 'File not found. Please upload again.' });
  }

  console.log(`[API] Analyze: ${fileId}, mode: ${mode}`);

  const fileData = fileStore.get(fileId);
  const eng = getEngine();
  
  const result = await eng.process(fileData.buffer, {
    fileName: fileData.fileName,
    mode,
  });

  if (!result.success) {
    console.error(`[API] Analyze failed:`, result.error);
    return res.status(400).json({ success: false, error: result.error });
  }

  const resultId = uuidv4();
  fileStore.set(resultId, {
    buffer: result.output.buffer,
    fileName: result.output.filename,
    timestamp: Date.now(),
  });

  console.log(`[API] Analyze complete: ${resultId}`);

  res.json({
    success: true,
    resultId,
    fileName: result.output.filename,
    totalTime: result.totalTimeFormatted,
    summary: result.summary,
    stages: result.stages,
    analysis: {
      qualityScore: result.analysis?.qualityScore,
      totalIssues: result.analysis?.totalIssues || 0,
      autoFixCount: result.analysis?.autoFixCount || 0,
      needsReviewCount: result.analysis?.needsReviewCount || 0,
      criticalCount: result.analysis?.criticalCount || 0,
    },
    changes: (result.changes || []).filter(c => c.type === 'SUMMARY'),
  });
}));

/**
 * POST /api/clean
 */
router.post('/clean', asyncHandler(async (req, res) => {
  const { fileId, options = {} } = req.body;

  if (!fileId || !fileStore.has(fileId)) {
    return res.status(400).json({ success: false, error: 'File not found' });
  }

  console.log(`[API] Clean: ${fileId}`);

  const fileData = fileStore.get(fileId);
  const eng = getEngine();
  
  const result = await eng.quickClean(fileData.buffer, {
    fileName: fileData.fileName,
    ...options,
  });

  const resultId = uuidv4();
  fileStore.set(resultId, {
    buffer: result.buffer,
    fileName: `cleaned_${fileData.fileName}`,
    timestamp: Date.now(),
  });

  console.log(`[API] Clean complete: ${resultId}`);

  res.json({
    success: true,
    resultId,
    fileName: `cleaned_${fileData.fileName}`,
    stats: result.stats,
    changes: result.changes,
  });
}));

/**
 * POST /api/convert
 */
router.post('/convert', asyncHandler(async (req, res) => {
  const { fileId, format } = req.body;

  if (!fileId || !fileStore.has(fileId)) {
    return res.status(400).json({ success: false, error: 'File not found' });
  }

  console.log(`[API] Convert: ${fileId} to ${format}`);

  const fileData = fileStore.get(fileId);
  const eng = getEngine();
  
  const result = await eng.convert(fileData.buffer, format, {
    fileName: fileData.fileName,
    pretty: true,
  });

  const extensions = {
    csv: '.csv', json: '.json', html: '.html',
    markdown: '.md', sql: '.sql', xml: '.xml',
  };

  const baseName = fileData.fileName.replace(/\.[^/.]+$/, '');
  const outputFileName = `${baseName}${extensions[format]}`;

  const resultId = uuidv4();
  fileStore.set(resultId, {
    buffer: Buffer.from(result.output, 'utf-8'),
    fileName: outputFileName,
    timestamp: Date.now(),
  });

  console.log(`[API] Convert complete: ${resultId}`);

  res.json({
    success: true,
    resultId,
    fileName: outputFileName,
    format,
    preview: result.output.substring(0, 500),
  });
}));

/**
 * POST /api/create
 */
router.post('/create', asyncHandler(async (req, res) => {
  const { text, options = {} } = req.body;

  if (!text) {
    return res.status(400).json({ success: false, error: 'No text data provided' });
  }

  console.log(`[API] Create from text (${text.length} chars)`);

  const eng = getEngine();
  const result = await eng.createFromText(text, options);

  const resultId = uuidv4();
  fileStore.set(resultId, {
    buffer: result.buffer,
    fileName: 'created_data.xlsx',
    timestamp: Date.now(),
  });

  console.log(`[API] Create complete: ${resultId}`);

  res.json({
    success: true,
    resultId,
    fileName: 'created_data.xlsx',
    format: result.format,
    headers: result.headers,
    rowCount: result.rowCount,
  });
}));

/**
 * POST /api/format
 */
router.post('/format', asyncHandler(async (req, res) => {
  const { fileId, instructions } = req.body;

  if (!fileId || !fileStore.has(fileId)) {
    return res.status(400).json({ success: false, error: 'File not found' });
  }

  if (!instructions) {
    return res.status(400).json({ success: false, error: 'No instructions provided' });
  }

  console.log(`[API] Format: ${fileId}`);

  const fileData = fileStore.get(fileId);
  const eng = getEngine();
  
  const result = await eng.applyFormat(fileData.buffer, instructions, {
    fileName: fileData.fileName,
  });

  const resultId = uuidv4();
  fileStore.set(resultId, {
    buffer: result.buffer,
    fileName: `formatted_${fileData.fileName}`,
    timestamp: Date.now(),
  });

  console.log(`[API] Format complete: ${resultId}`);

  res.json({
    success: true,
    resultId,
    fileName: `formatted_${fileData.fileName}`,
    instructionsApplied: result.instructionsApplied,
    instructions: result.instructions,
  });
}));

/**
 * POST /api/template
 */
router.post('/template', asyncHandler(async (req, res) => {
  const { type, options = {} } = req.body;

  if (!type) {
    return res.status(400).json({ success: false, error: 'Template type required' });
  }

  console.log(`[API] Template: ${type}`);

  const eng = getEngine();
  const result = await eng.generateTemplate(type, options);

  const resultId = uuidv4();
  fileStore.set(resultId, {
    buffer: result.buffer,
    fileName: `${type}_template.xlsx`,
    timestamp: Date.now(),
  });

  console.log(`[API] Template complete: ${resultId}`);

  res.json({
    success: true,
    resultId,
    fileName: `${type}_template.xlsx`,
    template: type,
  });
}));

/**
 * GET /api/templates
 */
router.get('/templates', (req, res) => {
  try {
    const eng = getEngine();
    const templates = eng.listTemplates();
    res.json({ success: true, templates });
  } catch (error) {
    res.json({ success: true, templates: [] });
  }
});

/**
 * GET /api/download/:id
 */
router.get('/download/:id', (req, res) => {
  const { id } = req.params;

  if (!fileStore.has(id)) {
    return res.status(404).json({ success: false, error: 'File not found or expired' });
  }

  const fileData = fileStore.get(id);
  
  res.setHeader('Content-Type', 'application/octet-stream');
  res.setHeader('Content-Disposition', `attachment; filename="${fileData.fileName}"`);
  res.send(fileData.buffer);
});

// ─────────────────────────────────────────────────────────────────────────
// ERROR HANDLER
// ─────────────────────────────────────────────────────────────────────────

router.use((err, req, res, next) => {
  console.error('[API] Error:', err.message);
  
  if (err instanceof multer.MulterError) {
    if (err.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ success: false, error: 'File too large. Maximum 10MB.' });
    }
  }
  
  res.status(500).json({ 
    success: false, 
    error: err.message || 'Internal server error' 
  });
});

module.exports = router;
