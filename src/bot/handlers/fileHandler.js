// ═══════════════════════════════════════════════════════════════════════════
// FILE HANDLER - Handle file uploads from Discord
// ═══════════════════════════════════════════════════════════════════════════

const https = require('https');
const http = require('http');
const path = require('path');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid');

class FileHandler {
  constructor() {
    this.tempDir = path.join(process.cwd(), 'temp');
    this.maxFileSize = 10 * 1024 * 1024; // 10MB
    this.allowedExtensions = ['.xlsx', '.xls', '.csv'];
    
    // Ensure temp directory exists
    if (!fs.existsSync(this.tempDir)) {
      fs.mkdirSync(this.tempDir, { recursive: true });
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // DOWNLOAD FILE FROM DISCORD
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Download attachment from Discord
   * @param {Attachment} attachment - Discord attachment object
   * @returns {Promise<Object>} File info with buffer
   */
  async downloadAttachment(attachment) {
    // Validate file
    const validation = this.validateFile(attachment);
    if (!validation.valid) {
      throw new Error(validation.error);
    }

    return new Promise((resolve, reject) => {
      const url = attachment.url;
      const protocol = url.startsWith('https') ? https : http;

      const chunks = [];
      let totalSize = 0;

      protocol.get(url, (response) => {
        if (response.statusCode !== 200) {
          reject(new Error(`Failed to download: HTTP ${response.statusCode}`));
          return;
        }

        response.on('data', (chunk) => {
          totalSize += chunk.length;
          
          if (totalSize > this.maxFileSize) {
            response.destroy();
            reject(new Error(`File too large (max ${this.maxFileSize / 1024 / 1024}MB)`));
            return;
          }
          
          chunks.push(chunk);
        });

        response.on('end', () => {
          const buffer = Buffer.concat(chunks);
          
          resolve({
            buffer,
            fileName: attachment.name,
            fileSize: buffer.length,
            extension: path.extname(attachment.name).toLowerCase(),
          });
        });

        response.on('error', reject);
      }).on('error', reject);
    });
  }

  /**
   * Validate file before processing
   */
  validateFile(attachment) {
    // Check size
    if (attachment.size > this.maxFileSize) {
      return {
        valid: false,
        error: `File too large. Maximum size: ${this.maxFileSize / 1024 / 1024}MB`,
      };
    }

    // Check extension
    const ext = path.extname(attachment.name).toLowerCase();
    if (!this.allowedExtensions.includes(ext)) {
      return {
        valid: false,
        error: `Invalid file type. Allowed: ${this.allowedExtensions.join(', ')}`,
      };
    }

    return { valid: true };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // SAVE & CLEANUP
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Save buffer to temp file
   */
  async saveToTemp(buffer, fileName) {
    const id = uuidv4();
    const ext = path.extname(fileName);
    const tempFileName = `${id}${ext}`;
    const tempPath = path.join(this.tempDir, tempFileName);

    fs.writeFileSync(tempPath, buffer);

    return {
      id,
      path: tempPath,
      fileName: tempFileName,
    };
  }

  /**
   * Delete temp file
   */
  deleteTemp(filePath) {
    try {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    } catch (error) {
      console.error('Failed to delete temp file:', error);
    }
  }

  /**
   * Cleanup old temp files
   */
  cleanupOldFiles(maxAgeMs = 3600000) { // 1 hour default
    const files = fs.readdirSync(this.tempDir);
    const now = Date.now();

    files.forEach(file => {
      const filePath = path.join(this.tempDir, file);
      const stats = fs.statSync(filePath);
      
      if (now - stats.mtimeMs > maxAgeMs) {
        fs.unlinkSync(filePath);
      }
    });
  }
}

module.exports = new FileHandler();
