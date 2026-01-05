// ═══════════════════════════════════════════════════════════════════════════
// API ROUTES - Express API endpoints
// ═══════════════════════════════════════════════════════════════════════════

const express = require('express');
const multer = require('multer');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const engine = require('../../engine');

const router = express.Router();

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
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB
});

// ─────────────────────────────────────────────────────────────────────────
// FILE STORAGE (In-memory for simplicity)
// ─────────────────────────────────────────────────────────────────────────

const fileStore = new Map();

// Cleanup old files every hour
setInterval(() => {
  const now = Date.now();
  const maxAge = 60 * 60 * 1000; // 1 hour
  
  for (const [id, data] of fileStore.entries()) {
    if (now - data.timestamp > maxAge) {
      fileStore.delete(id);
    }
  }
}, 60 * 60 * 1000);

// ─────────────────────────────────────────────────────────────────────────
// ROUTES
// ─────────────────────────────────────────────────────────────────────────

/**
 * POST /api/upload
 * Upload file for processing
 */
router.post('/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No file uploaded' });
    }

    const fileId = uuidv4();
    const parsed = engine.fileParser.parse(req.file.buffer, {
      fileName: req.file.originalname,
    });

    if (!parsed.success) {
      return res.status(400).json({ success: false, error: parsed.error });
    }

    // Store file data
    fileStore.set(fileId, {
      buffer: req.file.buffer,
      fileName: req.file.originalname,
      parsed,
      timestamp: Date.now(),
    });

    res.json({
      success: true,
      fileId,
      fileName: req.file.originalname,
      fileSize: req.file.size,
      preview: {
        headers: parsed.headers,
        data: parsed.data.slice(0, 10), // First 10 rows
        rowCount: parsed.rowCount,
        columnCount: parsed.columnCount,
      },
      columnTypes: parsed.columnTypes,
    });

  } catch (error) {
    console.error('Upload error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * POST /api/analyze
 * Full intelligent analysis
 */
router.post('/analyze', async (req, res) => {
  try {
    const { fileId, mode = 'auto' } = req.body;

    if (!fileId || !fileStore.has(fileId)) {
      return res.status(400).json({ success: false, error: 'File not found. Please upload again.' });
    }

    const fileData = fileStore.get(fileId);
    
    const result = await engine.process(fileData.buffer, {
      fileName: fileData.fileName,
      mode,
    });

    if (!result.success) {
      return res.status(400).json({ success: false, error: result.error });
    }

    // Store result for download
    const resultId = uuidv4();
    fileStore.set(resultId, {
      buffer: result.output.buffer,
      fileName: result.output.filename,
      timestamp: Date.now(),
    });

    res.json({
      success: true,
      resultId,
      fileName: result.output.filename,
      totalTime: result.totalTimeFormatted,
      summary: result.summary,
      stages: result.stages,
      analysis: {
        qualityScore: result.analysis.qualityScore,
        totalIssues: result.analysis.totalIssues,
        autoFixCount: result.analysis.autoFixCount,
        needsReviewCount: result.analysis.needsReviewCount,
        criticalCount: result.analysis.criticalCount,
      },
      changes: result.changes.filter(c => c.type === 'SUMMARY'),
    });

  } catch (error) {
    console.error('Analyze error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * POST /api/clean
 * Quick cleaning
 */
router.post('/clean', async (req, res) => {
  try {
    const { fileId, options = {} } = req.body;

    if (!fileId || !fileStore.has(fileId)) {
      return res.status(400).json({ success: false, error: 'File not found' });
    }

    const fileData = fileStore.get(fileId);
    
    const result = await engine.quickClean(fileData.buffer, {
      fileName: fileData.fileName,
      ...options,
    });

    if (!result.success) {
      return res.status(400).json({ success: false, error: result.error });
    }

    const resultId = uuidv4();
    fileStore.set(resultId, {
      buffer: result.buffer,
      fileName: `cleaned_${fileData.fileName}`,
      timestamp: Date.now(),
    });

    res.json({
      success: true,
      resultId,
      fileName: `cleaned_${fileData.fileName}`,
      stats: result.stats,
      changes: result.changes,
    });

  } catch (error) {
    console.error('Clean error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * POST /api/convert
 * Format conversion
 */
router.post('/convert', async (req, res) => {
  try {
    const { fileId, format } = req.body;

    if (!fileId || !fileStore.has(fileId)) {
      return res.status(400).json({ success: false, error: 'File not found' });
    }

    const fileData = fileStore.get(fileId);
    
    const result = await engine.convert(fileData.buffer, format, {
      fileName: fileData.fileName,
      pretty: true,
    });

    if (!result.success) {
      return res.status(400).json({ success: false, error: result.error });
    }

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

    res.json({
      success: true,
      resultId,
      fileName: outputFileName,
      format,
      preview: result.output.substring(0, 500),
    });

  } catch (error) {
    console.error('Convert error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * POST /api/create
 * Create Excel from text
 */
router.post('/create', async (req, res) => {
  try {
    const { text, options = {} } = req.body;

    if (!text) {
      return res.status(400).json({ success: false, error: 'No text data provided' });
    }

    const result = await engine.createFromText(text, options);

    if (!result.success) {
      return res.status(400).json({ success: false, error: result.error });
    }

    const resultId = uuidv4();
    fileStore.set(resultId, {
      buffer: result.buffer,
      fileName: 'created_data.xlsx',
      timestamp: Date.now(),
    });

    res.json({
      success: true,
      resultId,
      fileName: 'created_data.xlsx',
      format: result.format,
      headers: result.headers,
      rowCount: result.rowCount,
    });

  } catch (error) {
    console.error('Create error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * POST /api/format
 * Apply custom formatting
 */
router.post('/format', async (req, res) => {
  try {
    const { fileId, instructions } = req.body;

    if (!fileId || !fileStore.has(fileId)) {
      return res.status(400).json({ success: false, error: 'File not found' });
    }

    if (!instructions) {
      return res.status(400).json({ success: false, error: 'No instructions provided' });
    }

    const fileData = fileStore.get(fileId);
    
    const result = await engine.applyFormat(fileData.buffer, instructions, {
      fileName: fileData.fileName,
    });

    if (!result.success) {
      return res.status(400).json({ success: false, error: result.error });
    }

    const resultId = uuidv4();
    fileStore.set(resultId, {
      buffer: result.buffer,
      fileName: `formatted_${fileData.fileName}`,
      timestamp: Date.now(),
    });

    res.json({
      success: true,
      resultId,
      fileName: `formatted_${fileData.fileName}`,
      instructionsApplied: result.instructionsApplied,
      instructions: result.instructions,
    });

  } catch (error) {
    console.error('Format error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * POST /api/template
 * Generate template
 */
router.post('/template', async (req, res) => {
  try {
    const { type, options = {} } = req.body;

    if (!type) {
      return res.status(400).json({ success: false, error: 'Template type required' });
    }

    const result = await engine.generateTemplate(type, options);

    if (!result.success) {
      return res.status(400).json({ success: false, error: result.error });
    }

    const resultId = uuidv4();
    fileStore.set(resultId, {
      buffer: result.buffer,
      fileName: `${type}_template.xlsx`,
      timestamp: Date.now(),
    });

    res.json({
      success: true,
      resultId,
      fileName: `${type}_template.xlsx`,
      template: type,
    });

  } catch (error) {
    console.error('Template error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * GET /api/templates
 * List available templates
 */
router.get('/templates', (req, res) => {
  const templates = engine.listTemplates();
  res.json({ success: true, templates });
});

/**
 * GET /api/download/:id
 * Download processed file
 */
router.get('/download/:id', (req, res) => {
  try {
    const { id } = req.params;

    if (!fileStore.has(id)) {
      return res.status(404).json({ success: false, error: 'File not found or expired' });
    }

    const fileData = fileStore.get(id);
    
    res.setHeader('Content-Type', 'application/octet-stream');
    res.setHeader('Content-Disposition', `attachment; filename="${fileData.fileName}"`);
    res.send(fileData.buffer);

  } catch (error) {
    console.error('Download error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * GET /api/health
 * Health check
 */
router.get('/health', (req, res) => {
  res.json({
    success: true,
    status: 'healthy',
    timestamp: new Date().toISOString(),
    filesInMemory: fileStore.size,
  });
});

module.exports = router;
