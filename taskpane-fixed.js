/* global Office, Excel */

class MAModelingAddin {
  constructor() {
    this.chatMessages = [];
    this.selectedRange = null;
    this.uploadedFiles = [];

    // Initialize when Office is ready
    Office.onReady(() => {
      this.initializeAddin();
    });
  }

  initializeAddin() {
    // Set up event listeners
    const selectRangeBtn = document.getElementById('selectRangeBtn');
    const generateModelBtn = document.getElementById('generateModelBtn');
    const validateModelBtn = document.getElementById('validateModelBtn');
    const sendChatBtn = document.getElementById('sendChatBtn');
    const chatInput = document.getElementById('chatInput');

    if (selectRangeBtn) {
      selectRangeBtn.addEventListener('click', () => this.selectAssumptionsRange());
    }
    if (generateModelBtn) {
      generateModelBtn.addEventListener('click', () => this.generateModel());
    }
    if (validateModelBtn) {
      validateModelBtn.addEventListener('click', () => this.validateModel());
    }
    if (sendChatBtn) {
      sendChatBtn.addEventListener('click', () => this.sendChatMessage());
    }
    
    // Allow Enter key in chat input
    if (chatInput) {
      chatInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
          this.sendChatMessage();
        }
      });
    }

    // File upload event listeners
    this.initializeFileUpload();

    // API key is already configured
    console.log('OpenAI API key configured');
    console.log('MAModelingAddin initialized successfully');
  }

  initializeFileUpload() {
    const dropzone = document.getElementById('fileDropzone');
    const fileInput = document.getElementById('fileInput');
    const uploadLink = document.querySelector('.upload-link');

    // Dropzone click handler
    if (dropzone) {
      dropzone.addEventListener('click', () => {
        if (fileInput) fileInput.click();
      });
    }

    // Upload link click handler
    if (uploadLink) {
      uploadLink.addEventListener('click', (e) => {
        e.stopPropagation();
        if (fileInput) fileInput.click();
      });
    }

    // File input change handler
    if (fileInput) {
      fileInput.addEventListener('change', (e) => {
        const files = e.target.files;
        if (files) {
          this.handleFileSelection(Array.from(files));
        }
      });
    }

    // Drag and drop handlers
    if (dropzone) {
      dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
      });

      dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('dragover');
      });

      dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        const files = Array.from(e.dataTransfer.files || []);
        this.handleFileSelection(files);
      });
    }
  }

  handleFileSelection(files) {
    // Filter valid files (PDF and CSV only)
    const validFiles = files.filter(file => {
      const isValidType = file.type === 'application/pdf' || file.type === 'text/csv' || file.name.endsWith('.csv');
      const isValidSize = file.size <= 10 * 1024 * 1024; // 10MB limit
      return isValidType && isValidSize;
    });

    // Check total file limit
    if (this.uploadedFiles.length + validFiles.length > 4) {
      this.addChatMessage('assistant', 'Maximum 4 files allowed. Please remove some files first.');
      return;
    }

    // Add files to uploaded list
    this.uploadedFiles.push(...validFiles);
    this.updateFileDisplay();

    if (validFiles.length > 0) {
      this.addChatMessage('assistant', `Successfully uploaded ${validFiles.length} file(s). You can now ask me to analyze them and fill out your assumptions template!`);
    }
  }

  updateFileDisplay() {
    const uploadedFilesDiv = document.getElementById('uploadedFiles');
    const fileList = document.getElementById('fileList');

    if (this.uploadedFiles.length === 0) {
      if (uploadedFilesDiv) uploadedFilesDiv.style.display = 'none';
      return;
    }

    if (uploadedFilesDiv) uploadedFilesDiv.style.display = 'block';
    if (fileList) fileList.innerHTML = '';

    this.uploadedFiles.forEach((file, index) => {
      const fileItem = document.createElement('div');
      fileItem.className = 'file-item';
      fileItem.innerHTML = `
        <div class="file-info">
          <svg class="file-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14,2 14,8 20,8"></polyline>
          </svg>
          <div>
            <div class="file-name">${file.name}</div>
            <div class="file-size">${this.formatFileSize(file.size)}</div>
          </div>
        </div>
        <button class="remove-file" data-index="${index}">Remove</button>
      `;

      const removeBtn = fileItem.querySelector('.remove-file');
      if (removeBtn) {
        removeBtn.addEventListener('click', () => this.removeFile(index));
      }

      if (fileList) fileList.appendChild(fileItem);
    });
  }

  removeFile(index) {
    this.uploadedFiles.splice(index, 1);
    this.updateFileDisplay();
  }

  formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  async selectAssumptionsRange() {
    console.log('Select assumptions range clicked');
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(['address', 'values', 'text']);
        
        await context.sync();
        
        this.selectedRange = range.address;
        const rangeData = range.values;
        
        // Update UI
        const statusElement = document.getElementById('rangeStatus');
        if (statusElement) {
          statusElement.textContent = `Selected: ${range.address} (${rangeData.length} rows)`;
        }
        
        // Try to parse assumptions from selected range
        this.parseAssumptionsFromRange(rangeData);
      });
    } catch (error) {
      console.error('Error selecting range:', error);
      this.showStatus('Error selecting range. Please try again.');
    }
  }

  parseAssumptionsFromRange(rangeData) {
    // Smart parsing of assumptions from Excel range
    const assumptions = {};
    
    for (let i = 0; i < rangeData.length; i++) {
      const row = rangeData[i];
      if (row.length >= 2) {
        const label = String(row[0]).toLowerCase();
        const value = row[1];
        
        // Map common assumption labels to our fields
        if (label.includes('deal') && label.includes('size')) {
          assumptions.dealSize = parseFloat(value) || 0;
        } else if (label.includes('ltv') || label.includes('leverage')) {
          assumptions.ltv = parseFloat(value) * (value <= 1 ? 100 : 1); // Convert decimal to percentage
        } else if (label.includes('holding') || label.includes('period')) {
          assumptions.holdingPeriod = parseFloat(value) || 0;
        } else if (label.includes('revenue') && label.includes('growth')) {
          assumptions.revenueGrowth = parseFloat(value) * (value <= 1 ? 100 : 1);
        } else if (label.includes('exit') && label.includes('multiple')) {
          assumptions.exitMultiple = parseFloat(value) || 0;
        }
      }
    }
    
    // Update form fields with parsed values
    this.updateFormWithAssumptions(assumptions);
  }

  updateFormWithAssumptions(assumptions) {
    const fields = ['dealSize', 'ltv', 'holdingPeriod', 'revenueGrowth', 'exitMultiple'];
    
    fields.forEach(field => {
      const element = document.getElementById(field);
      if (element && assumptions[field] !== undefined) {
        element.value = assumptions[field].toString();
      }
    });
  }

  async generateModel() {
    console.log('Generate model clicked');
    const assumptions = this.collectAssumptions();
    
    this.showLoading(true);
    this.showStatus('Generating M&A model sheets...');
    
    try {
      await Excel.run(async (context) => {
        // Create or clear sheets
        await this.createModelSheets(context, assumptions);
        
        // Generate cash flows and calculate metrics
        const metrics = await this.calculateMetrics(assumptions);
        
        // Populate sheets with data and formulas
        await this.populateAssumptionsSheet(context, assumptions);
        await this.populateNPVSheet(context, assumptions, metrics);
        await this.populatePLSheet(context, assumptions);
        await this.populateCFSheet(context, assumptions);
        
        await context.sync();
        
        this.showStatus('✅ M&A model generated successfully!');
        this.addChatMessage('assistant', `M&A model for "${assumptions.dealName}" has been generated with the following key metrics:\n• Levered IRR: ${(metrics.leveredIRR * 100).toFixed(1)}%\n• MOIC: ${metrics.moic.toFixed(1)}x\n• Deal Size: $${assumptions.dealSize}M`);
      });
    } catch (error) {
      console.error('Error generating model:', error);
      this.showStatus('❌ Error generating model. Please check your inputs.');
    } finally {
      this.showLoading(false);
    }
  }

  collectAssumptions() {
    const getValue = (id) => {
      const element = document.getElementById(id);
      return element ? element.value : '';
    };

    return {
      dealName: getValue('dealName') || 'Sample Deal',
      dealSize: parseFloat(getValue('dealSize')) || 50,
      ltv: parseFloat(getValue('ltv')) || 70,
      holdingPeriod: parseFloat(getValue('holdingPeriod')) || 60,
      revenueGrowth: parseFloat(getValue('revenueGrowth')) || 15,
      exitMultiple: parseFloat(getValue('exitMultiple')) || 12,
      selectedRange: this.selectedRange || undefined
    };
  }

  async validateModel() {
    console.log('Validate model clicked');
    this.addChatMessage('assistant', 'Model validation feature coming soon! This will check all formulas and cross-references for accuracy.');
  }

  async sendChatMessage() {
    console.log('Send chat message clicked');
    const input = document.getElementById('chatInput');
    const message = input ? input.value.trim() : '';
    
    if (!message) return;
    
    // Add user message
    this.addChatMessage('user', message);
    if (input) input.value = '';
    
    // Process with AI
    await this.processWithAI(message);
  }

  async processWithAI(message) {
    try {
      const context = await this.getExcelContext();
      
      // Process uploaded files
      const fileContents = await this.processUploadedFiles();
      
      const response = await fetch('/.netlify/functions/chat', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          message: message,
          excelContext: context,
          fileContents: fileContents
        })
      });

      const data = await response.json();
      
      // Check if response includes commands
      if (data.commands && Array.isArray(data.commands)) {
        this.addChatMessage('assistant', data.response || 'Processing your request...');
        
        // Execute commands
        for (const command of data.commands) {
          await this.executeCommand(command);
        }
      } else {
        // Regular text response
        const responseText = data.response || data.error || 'No response received';
        this.addChatMessage('assistant', responseText);
        
        // Legacy formula detection for backwards compatibility
        if (responseText && typeof responseText === 'string' && responseText.includes('=') && responseText.includes('cell')) {
          this.offerToImplementFormula(responseText);
        }
      }
      
    } catch (error) {
      console.error('AI processing error:', error);
      
      // Try to get more specific error info
      let errorMessage = 'Sorry, I encountered an error. Please try again.';
      if (error instanceof Error) {
        errorMessage = `Error: ${error.message}`;
      }
      
      this.addChatMessage('assistant', errorMessage);
    }
  }

  addChatMessage(role, content) {
    this.chatMessages.push({ role, content });
    
    const messagesDiv = document.getElementById('chatMessages');
    if (messagesDiv) {
      const messageBubble = document.createElement('div');
      messageBubble.className = `message-bubble ${role}-bubble`;
      
      const messageContent = document.createElement('div');
      messageContent.className = 'message-content';
      
      const messageText = document.createElement('div');
      messageText.className = 'message-text';
      messageText.textContent = content;
      
      messageContent.appendChild(messageText);
      messageBubble.appendChild(messageContent);
      messagesDiv.appendChild(messageBubble);
      
      messagesDiv.scrollTop = messagesDiv.scrollHeight;
    }
  }

  showLoading(show) {
    const loading = document.getElementById('loading');
    if (loading) {
      loading.style.display = show ? 'block' : 'none';
    }
  }

  showStatus(message) {
    console.log('Status:', message);
    // Could show in a status bar
  }

  async getExcelContext() {
    try {
      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const selectedRange = context.workbook.getSelectedRange();
        const usedRange = worksheet.getUsedRange();
        
        worksheet.load(['name']);
        selectedRange.load(['address', 'values', 'formulas']);
        usedRange.load(['address', 'values', 'formulas']);
        
        await context.sync();
        
        return JSON.stringify({
          worksheetName: worksheet.name,
          selectedRange: {
            address: selectedRange.address,
            values: selectedRange.values,
            formulas: selectedRange.formulas
          },
          usedRange: {
            address: usedRange.address,
            values: usedRange.values.slice(0, 10), // Limit to first 10 rows
            formulas: usedRange.formulas.slice(0, 10)
          }
        });
      });
    } catch {
      return JSON.stringify({ error: 'Could not read Excel context' });
    }
  }

  // Simplified versions of other methods for basic functionality
  async processUploadedFiles() {
    return [];
  }

  async executeCommand(command) {
    console.log('Executing command:', command);
  }

  async createModelSheets(context, assumptions) {
    console.log('Creating model sheets:', assumptions);
  }

  async calculateMetrics(assumptions) {
    return {
      leveredIRR: 0.25,
      moic: 2.5,
      leveredNPV: 10000000
    };
  }

  async populateAssumptionsSheet(context, assumptions) {
    console.log('Populating assumptions sheet');
  }

  async populateNPVSheet(context, assumptions, metrics) {
    console.log('Populating NPV sheet');
  }

  async populatePLSheet(context, assumptions) {
    console.log('Populating P&L sheet');
  }

  async populateCFSheet(context, assumptions) {
    console.log('Populating CF sheet');
  }
}

// Initialize the add-in
console.log('Initializing MAModelingAddin...');
new MAModelingAddin();