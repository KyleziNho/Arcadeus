/**
 * FileDropZone.js - Enhanced drag & drop interface with multi-format support
 * Supports: PDF, CSV, Excel, PNG, JPG with proper validation and preview
 */

class FileDropZone {
  constructor() {
    this.supportedFormats = {
      'application/pdf': { ext: '.pdf', icon: 'file-text', processor: 'pdf' },
      'text/csv': { ext: '.csv', icon: 'file-spreadsheet', processor: 'csv' },
      'application/vnd.ms-excel': { ext: '.xls', icon: 'file-spreadsheet', processor: 'excel' },
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': { ext: '.xlsx', icon: 'file-spreadsheet', processor: 'excel' },
      'image/png': { ext: '.png', icon: 'file-image', processor: 'ocr' },
      'image/jpeg': { ext: '.jpg', icon: 'file-image', processor: 'ocr' },
      'image/jpg': { ext: '.jpg', icon: 'file-image', processor: 'ocr' }
    };
    
    this.maxFileSize = 25 * 1024 * 1024; // 25MB
    this.maxFiles = 10;
    this.uploadedFiles = [];
    this.onFilesProcessed = null;
    this.isProcessing = false;
  }

  initialize() {
    this.setupDOM();
    this.attachEventListeners();
    console.log('✅ FileDropZone initialized with enhanced multi-format support');
  }

  setupDOM() {
    const container = document.getElementById('mainUploadZone');
    if (!container) return;

    // Enhance the upload zone with better visual feedback
    container.innerHTML = `
      <div class="upload-content" id="dropZoneContent">
        <svg class="upload-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
          <path d="M7 10l5-5 5 5"></path>
          <line x1="12" y1="15" x2="12" y2="3"></line>
        </svg>
        <h2>Drop your files here</h2>
        <p>Upload M&A documents, financial reports, or spreadsheets</p>
        <div class="upload-specs">
          <span>PDF • Excel • CSV • Images (PNG/JPG)</span>
          <span>Up to 10 files • 25MB each</span>
        </div>
        <button class="browse-files-btn" id="browseFilesBtn">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14,2 14,8 20,8"></polyline>
          </svg>
          Browse Files
        </button>
        <div class="processing-indicator" id="processingIndicator" style="display: none;">
          <div class="spinner"></div>
          <span>Processing files...</span>
        </div>
      </div>
      <input type="file" id="mainFileInput" multiple accept=".pdf,.csv,.xls,.xlsx,.png,.jpg,.jpeg" style="display: none;">
    `;

    // Add file preview area
    const previewArea = document.getElementById('uploadedFilesDisplay');
    if (!previewArea) {
      const preview = document.createElement('div');
      preview.id = 'uploadedFilesDisplay';
      preview.className = 'uploaded-files-display';
      preview.style.display = 'none';
      container.parentNode.insertBefore(preview, container.nextSibling);
    }
  }

  attachEventListeners() {
    const dropZone = document.getElementById('mainUploadZone');
    const fileInput = document.getElementById('mainFileInput');
    const browseBtn = document.getElementById('browseFilesBtn');

    if (!dropZone || !fileInput) return;

    // Drag and drop events
    dropZone.addEventListener('dragover', this.handleDragOver.bind(this));
    dropZone.addEventListener('dragleave', this.handleDragLeave.bind(this));
    dropZone.addEventListener('drop', this.handleDrop.bind(this));

    // File input events
    fileInput.addEventListener('change', this.handleFileSelect.bind(this));
    
    // Browse button
    if (browseBtn) {
      browseBtn.addEventListener('click', () => fileInput.click());
    }

    // Click on drop zone
    dropZone.addEventListener('click', (e) => {
      if (e.target === browseBtn || browseBtn.contains(e.target)) return;
      fileInput.click();
    });
  }

  handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    const dropZone = document.getElementById('mainUploadZone');
    dropZone.classList.add('dragover');
  }

  handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    const dropZone = document.getElementById('mainUploadZone');
    if (!dropZone.contains(e.relatedTarget)) {
      dropZone.classList.remove('dragover');
    }
  }

  handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    const dropZone = document.getElementById('mainUploadZone');
    dropZone.classList.remove('dragover');
    
    const files = Array.from(e.dataTransfer.files);
    this.processFiles(files);
  }

  handleFileSelect(e) {
    const files = Array.from(e.target.files);
    this.processFiles(files);
    e.target.value = ''; // Reset input
  }

  async processFiles(files) {
    if (this.isProcessing) {
      this.showMessage('Please wait for current processing to complete', 'warning');
      return;
    }

    // Validate files
    const validFiles = this.validateFiles(files);
    if (validFiles.length === 0) return;

    // Check total file limit
    if (this.uploadedFiles.length + validFiles.length > this.maxFiles) {
      this.showMessage(`Maximum ${this.maxFiles} files allowed`, 'error');
      return;
    }

    this.isProcessing = true;
    this.showProcessingIndicator(true);

    try {
      // Process each file based on type
      const processedFiles = [];
      
      for (const file of validFiles) {
        console.log(`📄 Processing ${file.name} (${file.type})`);
        
        const fileData = {
          name: file.name,
          type: file.type,
          size: file.size,
          file: file,
          content: null,
          extractedData: null,
          processingStatus: 'pending',
          processor: this.getProcessorType(file)
        };

        try {
          // Read file content based on type
          if (fileData.processor === 'csv' || fileData.processor === 'excel') {
            fileData.content = await this.readTextFile(file);
          } else if (fileData.processor === 'pdf') {
            fileData.content = await this.processPDFFile(file);
          } else if (fileData.processor === 'ocr') {
            fileData.content = await this.processImageFile(file);
          }
          
          fileData.processingStatus = 'ready';
          processedFiles.push(fileData);
          
        } catch (error) {
          console.error(`Error processing ${file.name}:`, error);
          fileData.processingStatus = 'error';
          fileData.error = error.message;
          processedFiles.push(fileData);
        }
      }

      // Add to uploaded files
      this.uploadedFiles.push(...processedFiles);
      this.updateFileDisplay();
      
      // Trigger callback if set
      if (this.onFilesProcessed) {
        this.onFilesProcessed(processedFiles);
      }

      this.showMessage(`Successfully processed ${processedFiles.length} file(s)`, 'success');
      
    } catch (error) {
      console.error('Error processing files:', error);
      this.showMessage('Error processing files. Please try again.', 'error');
    } finally {
      this.isProcessing = false;
      this.showProcessingIndicator(false);
    }
  }

  validateFiles(files) {
    const validFiles = [];
    const errors = [];

    for (const file of files) {
      // Check file size
      if (file.size > this.maxFileSize) {
        errors.push(`${file.name} exceeds 25MB limit`);
        continue;
      }

      // Check file type
      const isValidType = Object.keys(this.supportedFormats).some(type => 
        file.type === type || file.name.toLowerCase().endsWith(this.supportedFormats[type]?.ext)
      );

      if (!isValidType) {
        errors.push(`${file.name} is not a supported format`);
        continue;
      }

      validFiles.push(file);
    }

    if (errors.length > 0) {
      this.showMessage(errors.join('\n'), 'error');
    }

    return validFiles;
  }

  getProcessorType(file) {
    const format = this.supportedFormats[file.type];
    if (format) return format.processor;
    
    // Fallback based on extension
    const ext = file.name.toLowerCase().split('.').pop();
    if (ext === 'csv') return 'csv';
    if (ext === 'xls' || ext === 'xlsx') return 'excel';
    if (ext === 'pdf') return 'pdf';
    if (ext === 'png' || ext === 'jpg' || ext === 'jpeg') return 'ocr';
    
    return 'unknown';
  }

  async readTextFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = (e) => reject(new Error('Failed to read file'));
      reader.readAsText(file);
    });
  }

  async processPDFFile(file) {
    // For now, return placeholder - will integrate PDF.js
    console.log('📑 PDF processing will be implemented with PDF.js');
    return `PDF: ${file.name} - PDF parsing will extract text content here`;
  }

  async processImageFile(file) {
    // For now, return placeholder - will integrate OCR
    console.log('🖼️ Image OCR will be implemented with Tesseract.js or API');
    return `Image: ${file.name} - OCR will extract text from image here`;
  }

  updateFileDisplay() {
    const display = document.getElementById('uploadedFilesDisplay');
    const filesGrid = document.getElementById('filesGrid');
    
    if (!display) return;
    
    if (this.uploadedFiles.length === 0) {
      display.style.display = 'none';
      return;
    }

    display.style.display = 'block';
    display.innerHTML = `
      <h3>Uploaded Files (${this.uploadedFiles.length})</h3>
      <div class="files-grid" id="filesGrid">
        ${this.uploadedFiles.map((file, index) => this.createFileCard(file, index)).join('')}
      </div>
      <div class="autofill-section">
        <button class="autofill-btn" id="autoFillBtn" ${this.uploadedFiles.length === 0 ? 'disabled' : ''}>
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M9 11l3 3L22 4"></path>
            <path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"></path>
          </svg>
          Extract & Auto-Fill Data
        </button>
        <small class="autofill-help">AI will analyze all files and populate relevant fields</small>
      </div>
    `;

    // Attach remove handlers
    this.uploadedFiles.forEach((file, index) => {
      const removeBtn = document.querySelector(`[data-remove-index="${index}"]`);
      if (removeBtn) {
        removeBtn.addEventListener('click', () => this.removeFile(index));
      }
    });
  }

  createFileCard(file, index) {
    const formatInfo = Object.values(this.supportedFormats).find(f => f.processor === file.processor);
    const icon = formatInfo?.icon || 'file';
    
    const statusClass = file.processingStatus === 'error' ? 'file-card-error' : 
                       file.processingStatus === 'ready' ? 'file-card-ready' : 'file-card-pending';

    return `
      <div class="file-card ${statusClass}">
        <svg class="file-card-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          ${this.getFileIcon(icon)}
        </svg>
        <div class="file-card-info">
          <div class="file-card-name">${file.name}</div>
          <div class="file-card-size">${this.formatFileSize(file.size)}</div>
          ${file.processingStatus === 'error' ? 
            `<div class="file-card-error-msg">${file.error || 'Processing failed'}</div>` : ''}
        </div>
        <button class="file-card-remove" data-remove-index="${index}" title="Remove file">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <line x1="18" y1="6" x2="6" y2="18"></line>
            <line x1="6" y1="6" x2="18" y2="18"></line>
          </svg>
        </button>
      </div>
    `;
  }

  getFileIcon(type) {
    const icons = {
      'file-text': '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line>',
      'file-spreadsheet': '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="8" y1="13" x2="16" y2="13"></line><line x1="8" y1="17" x2="16" y2="17"></line><line x1="12" y1="11" x2="12" y2="19"></line>',
      'file-image': '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><circle cx="10" cy="13" r="2"></circle><path d="m20 17-3.5-3.5a2 2 0 0 0-2.8 0L8 19"></path>',
      'file': '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline>'
    };
    return icons[type] || icons['file'];
  }

  removeFile(index) {
    this.uploadedFiles.splice(index, 1);
    this.updateFileDisplay();
    this.showMessage('File removed', 'info');
  }

  formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  showProcessingIndicator(show) {
    const indicator = document.getElementById('processingIndicator');
    if (indicator) {
      indicator.style.display = show ? 'flex' : 'none';
    }
  }

  showMessage(message, type = 'info') {
    // Create or update message element
    let messageEl = document.getElementById('fileDropMessage');
    if (!messageEl) {
      messageEl = document.createElement('div');
      messageEl.id = 'fileDropMessage';
      messageEl.className = 'file-drop-message';
      document.getElementById('mainUploadZone').parentNode.appendChild(messageEl);
    }

    messageEl.className = `file-drop-message file-drop-message-${type}`;
    messageEl.textContent = message;
    messageEl.style.display = 'block';

    // Auto-hide after 5 seconds
    setTimeout(() => {
      messageEl.style.display = 'none';
    }, 5000);
  }

  getUploadedFiles() {
    return this.uploadedFiles;
  }

  clearAllFiles() {
    this.uploadedFiles = [];
    this.updateFileDisplay();
  }
}

// Export for use
window.FileDropZone = FileDropZone;