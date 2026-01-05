// ═══════════════════════════════════════════════════════════════════════════
// EXCEL INTELLIGENT TOOL - Frontend Application
// ═══════════════════════════════════════════════════════════════════════════

const API_BASE = '/api';

// ─────────────────────────────────────────────────────────────────────────
// STATE MANAGEMENT
// ─────────────────────────────────────────────────────────────────────────

const state = {
  currentTab: 'analyze',
  files: {
    analyze: null,
    clean: null,
    convert: null,
  },
  fileIds: {
    analyze: null,
    clean: null,
    convert: null,
  },
  selectedTemplate: null,
  selectedFormat: null,
};

// ─────────────────────────────────────────────────────────────────────────
// INITIALIZATION
// ─────────────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  initTabs();
  initUploadZones();
  initButtons();
  initFormatButtons();
  loadTemplates();
});

// ─────────────────────────────────────────────────────────────────────────
// TAB NAVIGATION
// ─────────────────────────────────────────────────────────────────────────

function initTabs() {
  const tabs = document.querySelectorAll('.tab');
  
  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const tabName = tab.dataset.tab;
      
      // Update active tab
      tabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      
      // Show corresponding content
      document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
      });
      document.getElementById(`tab-${tabName}`).classList.add('active');
      
      state.currentTab = tabName;
      
      // Hide preview when switching tabs
      document.getElementById('previewSection').style.display = 'none';
    });
  });
}

// ─────────────────────────────────────────────────────────────────────────
// FILE UPLOAD HANDLING
// ─────────────────────────────────────────────────────────────────────────

function initUploadZones() {
  const zones = ['analyze', 'clean', 'convert'];
  
  zones.forEach(zone => {
    const uploadZone = document.getElementById(`${zone}Upload`);
    const fileInput = document.getElementById(`${zone}File`);
    
    if (!uploadZone || !fileInput) return;
    
    // Click to upload
    uploadZone.addEventListener('click', () => fileInput.click());
    
    // Drag & drop
    uploadZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      uploadZone.classList.add('dragover');
    });
    
    uploadZone.addEventListener('dragleave', () => {
      uploadZone.classList.remove('dragover');
    });
    
    uploadZone.addEventListener('drop', (e) => {
      e.preventDefault();
      uploadZone.classList.remove('dragover');
      
      const file = e.dataTransfer.files[0];
      if (file) handleFileSelect(zone, file);
    });
    
    // File input change
    fileInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) handleFileSelect(zone, file);
    });
  });
}

async function handleFileSelect(zone, file) {
  // Validate file type
  const validTypes = ['.xlsx', '.xls', '.csv'];
  const ext = '.' + file.name.split('.').pop().toLowerCase();
  
  if (!validTypes.includes(ext)) {
    showToast(`Invalid file type. Allowed: ${validTypes.join(', ')}`, 'error');
    return;
  }
  
  // Validate file size
  if (file.size > 10 * 1024 * 1024) {
    showToast('File too large. Maximum: 10MB', 'error');
    return;
  }
  
  state.files[zone] = file;
  
  // Update UI
  const uploadZone = document.getElementById(`${zone}Upload`);
  uploadZone.classList.add('has-file');
  uploadZone.innerHTML = `
    <div class="upload-icon">✅</div>
    <p>File selected</p>
    <div class="file-info">
      <span class="file-name">${file.name}</span>
      <span>(${formatBytes(file.size)})</span>
    </div>
  `;
  
  // Enable button
  const btnId = zone === 'analyze' ? 'btnAnalyze' : 
                zone === 'clean' ? 'btnClean' : null;
  if (btnId) {
    document.getElementById(btnId).disabled = false;
  }
  
  // Upload file to server
  await uploadFile(zone, file);
}

async function uploadFile(zone, file) {
  showLoading('Uploading file...');
  
  try {
    const formData = new FormData();
    formData.append('file', file);
    
    const response = await fetch(`${API_BASE}/upload`, {
      method: 'POST',
      body: formData,
    });
    
    const result = await response.json();
    
    if (!result.success) {
      throw new Error(result.error);
    }
    
    state.fileIds[zone] = result.fileId;
    
    // Show preview
    showPreview(result.preview);
    showToast('File uploaded successfully', 'success');
    
    // Enable format buttons for convert tab
    if (zone === 'convert') {
      document.querySelectorAll('.format-btn').forEach(btn => {
        btn.disabled = false;
      });
    }
    
  } catch (error) {
    showToast(error.message, 'error');
  } finally {
    hideLoading();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// BUTTONS INITIALIZATION
// ─────────────────────────────────────────────────────────────────────────

function initButtons() {
  // Analyze button
  document.getElementById('btnAnalyze')?.addEventListener('click', handleAnalyze);
  
  // Clean button
  document.getElementById('btnClean')?.addEventListener('click', handleClean);
  
  // Create button
  document.getElementById('btnCreate')?.addEventListener('click', handleCreate);
  
  // Template button
  document.getElementById('btnTemplate')?.addEventListener('click', handleTemplate);
}

// ─────────────────────────────────────────────────────────────────────────
// ANALYZE HANDLER
// ─────────────────────────────────────────────────────────────────────────

async function handleAnalyze() {
  const fileId = state.fileIds.analyze;
  if (!fileId) {
    showToast('Please upload a file first', 'warning');
    return;
  }
  
  const mode = document.getElementById('analyzeMode').value;
  
  showLoading('Analyzing data...');
  
  try {
    const response = await fetch(`${API_BASE}/analyze`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileId, mode }),
    });
    
    const result = await response.json();
    
    if (!result.success) {
      throw new Error(result.error);
    }
    
    // Display results
    displayAnalyzeResult(result);
    showToast('Analysis complete!', 'success');
    
  } catch (error) {
    showToast(error.message, 'error');
  } finally {
    hideLoading();
  }
}

function displayAnalyzeResult(result) {
  const container = document.getElementById('analyzeResult');
  container.style.display = 'block';
  
  const qualityBefore = result.summary.qualityBefore;
  const qualityAfter = result.summary.qualityAfter;
  
  container.innerHTML = `
    <div class="result-header">
      <h3>🧠 Analysis Complete</h3>
      <span>Processed in ${result.totalTime}</span>
    </div>
    
    <div class="quality-score">
      <div class="score-value">${qualityAfter}%</div>
      <div class="score-label">Data Quality Score</div>
      <div class="progress-bar">
        <div class="progress-fill" style="width: ${qualityAfter}%"></div>
      </div>
      <small>Improved from ${qualityBefore}% (+${qualityAfter - qualityBefore}%)</small>
    </div>
    
    <div class="stats-grid">
      <div class="stat-box">
        <div class="stat-value">${result.summary.originalRows.toLocaleString()}</div>
        <div class="stat-label">Original Rows</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${result.summary.cleanedRows.toLocaleString()}</div>
        <div class="stat-label">Cleaned Rows</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${result.summary.rowsRemoved}</div>
        <div class="stat-label">Rows Removed</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${result.analysis.autoFixCount}</div>
        <div class="stat-label">Issues Fixed</div>
      </div>
    </div>
    
    ${result.changes.length > 0 ? `
      <h4 style="margin-top: 20px; color: var(--success);">✅ Changes Applied</h4>
      <ul class="changes-list">
        ${result.changes.map(c => `<li>✓ ${c.operation}: ${c.count} changes</li>`).join('')}
      </ul>
    ` : ''}
    
    ${result.analysis.needsReviewCount > 0 ? `
      <div class="issue-category needs-review">
        <h4>⚠️ ${result.analysis.needsReviewCount} items need your review</h4>
      </div>
    ` : ''}
    
    <a href="${API_BASE}/download/${result.resultId}" 
       class="btn btn-download" 
       style="margin-top: 20px; display: inline-flex;">
      📥 Download Result
    </a>
  `;
}

// ─────────────────────────────────────────────────────────────────────────
// CLEAN HANDLER
// ─────────────────────────────────────────────────────────────────────────

async function handleClean() {
  const fileId = state.fileIds.clean;
  if (!fileId) {
    showToast('Please upload a file first', 'warning');
    return;
  }
  
  const options = {
    removeDuplicates: document.getElementById('optDuplicates').checked,
    removeEmpty: document.getElementById('optEmpty').checked,
    trimWhitespace: document.getElementById('optTrim').checked,
    textCase: document.getElementById('optTextCase').value || undefined,
  };
  
  showLoading('Cleaning data...');
  
  try {
    const response = await fetch(`${API_BASE}/clean`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileId, options }),
    });
    
    const result = await response.json();
    
    if (!result.success) {
      throw new Error(result.error);
    }
    
    displayCleanResult(result);
    showToast('Cleaning complete!', 'success');
    
  } catch (error) {
    showToast(error.message, 'error');
  } finally {
    hideLoading();
  }
}

function displayCleanResult(result) {
  const container = document.getElementById('cleanResult');
  container.style.display = 'block';
  
  container.innerHTML = `
    <div class="result-header">
      <h3>🧹 Cleaning Complete</h3>
    </div>
    
    <div class="stats-grid">
      <div class="stat-box">
        <div class="stat-value">${result.stats.rowsRemoved}</div>
        <div class="stat-label">Rows Removed</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${result.stats.cellsModified}</div>
        <div class="stat-label">Cells Modified</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${result.stats.totalChanges}</div>
        <div class="stat-label">Total Changes</div>
      </div>
    </div>
    
    <a href="${API_BASE}/download/${result.resultId}" 
       class="btn btn-download">
      📥 Download Cleaned File
    </a>
  `;
}

// ─────────────────────────────────────────────────────────────────────────
// CONVERT HANDLER
// ─────────────────────────────────────────────────────────────────────────

function initFormatButtons() {
  document.querySelectorAll('.format-btn').forEach(btn => {
    btn.addEventListener('click', () => handleConvert(btn.dataset.format));
  });
}

async function handleConvert(format) {
  const fileId = state.fileIds.convert;
  if (!fileId) {
    showToast('Please upload a file first', 'warning');
    return;
  }
  
  showLoading(`Converting to ${format.toUpperCase()}...`);
  
  try {
    const response = await fetch(`${API_BASE}/convert`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileId, format }),
    });
    
    const result = await response.json();
    
    if (!result.success) {
      throw new Error(result.error);
    }
    
    displayConvertResult(result);
    showToast('Conversion complete!', 'success');
    
  } catch (error) {
    showToast(error.message, 'error');
  } finally {
    hideLoading();
  }
}

function displayConvertResult(result) {
  const container = document.getElementById('convertResult');
  container.style.display = 'block';
  
  container.innerHTML = `
    <div class="result-header">
      <h3>🔄 Conversion Complete</h3>
    </div>
    
    <p style="color: var(--text-secondary); margin-bottom: 15px;">
      Converted to: <strong>${result.format.toUpperCase()}</strong>
    </p>
    
    <a href="${API_BASE}/download/${result.resultId}" 
       class="btn btn-download">
      📥 Download ${result.fileName}
    </a>
    
    ${result.preview ? `
      <div style="margin-top: 20px;">
        <h4>Preview:</h4>
        <pre style="background: rgba(0,0,0,0.3); padding: 15px; border-radius: 8px; overflow-x: auto; font-size: 0.85rem; color: var(--text-secondary);">${escapeHtml(result.preview)}...</pre>
      </div>
    ` : ''}
  `;
}

// ─────────────────────────────────────────────────────────────────────────
// CREATE HANDLER
// ─────────────────────────────────────────────────────────────────────────

async function handleCreate() {
  const text = document.getElementById('createText').value.trim();
  
  if (!text) {
    showToast('Please enter some data', 'warning');
    return;
  }
  
  const addTotals = document.getElementById('optAddTotals').checked;
  
  showLoading('Creating Excel...');
  
  try {
    const response = await fetch(`${API_BASE}/create`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text, options: { addTotals } }),
    });
    
    const result = await response.json();
    
    if (!result.success) {
      throw new Error(result.error);
    }
    
    displayCreateResult(result);
    showToast('Excel created!', 'success');
    
  } catch (error) {
    showToast(error.message, 'error');
  } finally {
    hideLoading();
  }
}

function displayCreateResult(result) {
  const container = document.getElementById('createResult');
  container.style.display = 'block';
  
  container.innerHTML = `
    <div class="result-header">
      <h3>📊 Excel Created</h3>
    </div>
    
    <div class="stats-grid">
      <div class="stat-box">
        <div class="stat-value">${result.headers.length}</div>
        <div class="stat-label">Columns</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${result.rowCount}</div>
        <div class="stat-label">Rows</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${result.format}</div>
        <div class="stat-label">Format Detected</div>
      </div>
    </div>
    
    <p style="color: var(--text-secondary); margin: 15px 0;">
      Columns: ${result.headers.join(', ')}
    </p>
    
    <a href="${API_BASE}/download/${result.resultId}" 
       class="btn btn-download">
      📥 Download Excel
    </a>
  `;
}

// ─────────────────────────────────────────────────────────────────────────
// TEMPLATE HANDLER
// ─────────────────────────────────────────────────────────────────────────

async function loadTemplates() {
  try {
    const response = await fetch(`${API_BASE}/templates`);
    const result = await response.json();
    
    if (result.success) {
      displayTemplates(result.templates);
    }
  } catch (error) {
    console.error('Failed to load templates:', error);
  }
}

function displayTemplates(templates) {
  const container = document.getElementById('templateGrid');
  
  const icons = {
    'invoice': '🧾',
    'payroll': '💵',
    'inventory': '📦',
    'sales-report': '📈',
    'budget': '💰',
    'attendance': '📅',
    'expense': '🧾',
  };
  
  container.innerHTML = templates.map(t => `
    <div class="template-card" data-template="${t.name}">
      <div class="icon">${icons[t.name] || '📄'}</div>
      <div class="name">${t.name}</div>
      <div class="desc">${t.description}</div>
    </div>
  `).join('');
  
  // Add click handlers
  container.querySelectorAll('.template-card').forEach(card => {
    card.addEventListener('click', () => {
      container.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
      card.classList.add('selected');
      state.selectedTemplate = card.dataset.template;
      document.getElementById('btnTemplate').disabled = false;
      document.getElementById('templateOptions').style.display = 'flex';
    });
  });
}

async function handleTemplate() {
  if (!state.selectedTemplate) {
    showToast('Please select a template', 'warning');
    return;
  }
  
  const companyName = document.getElementById('templateCompany').value;
  
  showLoading('Generating template...');
  
  try {
    const response = await fetch(`${API_BASE}/template`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ 
        type: state.selectedTemplate,
        options: { companyName }
      }),
    });
    
    const result = await response.json();
    
    if (!result.success) {
      throw new Error(result.error);
    }
    
    displayTemplateResult(result);
    showToast('Template generated!', 'success');
    
  } catch (error) {
    showToast(error.message, 'error');
  } finally {
    hideLoading();
  }
}

function displayTemplateResult(result) {
  const container = document.getElementById('templateResult');
  container.style.display = 'block';
  
  container.innerHTML = `
    <div class="result-header">
      <h3>📋 Template Ready</h3>
    </div>
    
    <p style="color: var(--text-secondary); margin-bottom: 15px;">
      Template: <strong>${result.template}</strong>
    </p>
    
    <a href="${API_BASE}/download/${result.resultId}" 
       class="btn btn-download">
      📥 Download ${result.fileName}
    </a>
  `;
}

// ─────────────────────────────────────────────────────────────────────────
// PREVIEW
// ─────────────────────────────────────────────────────────────────────────

function showPreview(preview) {
  const section = document.getElementById('previewSection');
  const table = document.getElementById('previewTable');
  const info = document.getElementById('previewInfo');
  
  section.style.display = 'block';
  info.textContent = `${preview.rowCount} rows × ${preview.columnCount} columns`;
  
  let html = '<thead><tr>';
  preview.headers.forEach(h => {
    html += `<th>${escapeHtml(h)}</th>`;
  });
  html += '</tr></thead><tbody>';
  
  preview.data.forEach(row => {
    html += '<tr>';
    preview.headers.forEach(h => {
      const val = row[h] !== undefined ? row[h] : '';
      html += `<td>${escapeHtml(String(val))}</td>`;
    });
    html += '</tr>';
  });
  
  if (preview.rowCount > preview.data.length) {
    html += `<tr><td colspan="${preview.headers.length}" style="text-align: center; color: var(--text-muted);">... and ${preview.rowCount - preview.data.length} more rows</td></tr>`;
  }
  
  html += '</tbody>';
  table.innerHTML = html;
}

// ─────────────────────────────────────────────────────────────────────────
// UTILITIES
// ─────────────────────────────────────────────────────────────────────────

function showLoading(text = 'Processing...') {
  document.getElementById('loadingText').textContent = text;
  document.getElementById('loadingOverlay').classList.add('active');
}

function hideLoading() {
  document.getElementById('loadingOverlay').classList.remove('active');
}

function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.textContent = message;
  container.appendChild(toast);
  
  setTimeout(() => {
    toast.style.animation = 'slideIn 0.3s ease reverse';
    setTimeout(() => toast.remove(), 300);
  }, 3000);
}

function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
              }
