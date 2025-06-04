/* global Office, Excel */

class MAModelingAddin {
  constructor() {
    this.chatMessages = [];
    this.selectedRange = null;
    this.uploadedFiles = [];
    this.isInitialized = false;

    // Initialize when Office is ready
    console.log('MAModelingAddin constructor called');
    
    // Check if Office is already available
    if (typeof Office !== 'undefined' && Office.onReady) {
      Office.onReady((info) => {
        console.log('Office.onReady fired:', info);
        this.initializeAddin();
      });
    } else {
      console.log('Office not available, trying fallback initialization');
      // Fallback - try to initialize after a delay
      setTimeout(() => {
        this.initializeAddin();
      }, 2000);
    }
  }

  initializeAddin() {
    if (this.isInitialized) {
      console.log('Add-in already initialized, skipping');
      return;
    }
    
    console.log('Initializing add-in...');
    
    // Wait for DOM to be ready
    if (document.readyState === 'loading') {
      console.log('DOM still loading, waiting...');
      document.addEventListener('DOMContentLoaded', () => {
        this.initializeAddin();
      });
      return;
    }
    
    // Set up event listeners
    const selectRangeBtn = document.getElementById('selectRangeBtn');
    const generateModelBtn = document.getElementById('generateModelBtn');
    const validateModelBtn = document.getElementById('validateModelBtn');
    const sendChatBtn = document.getElementById('sendChatBtn');
    const chatInput = document.getElementById('chatInput');
    
    console.log('DOM elements found:', {
      selectRangeBtn: !!selectRangeBtn,
      generateModelBtn: !!generateModelBtn,
      validateModelBtn: !!validateModelBtn,
      sendChatBtn: !!sendChatBtn,
      chatInput: !!chatInput
    });

    if (selectRangeBtn) {
      selectRangeBtn.addEventListener('click', () => this.selectAssumptionsRange());
      console.log('Select range button listener added');
    }
    if (generateModelBtn) {
      generateModelBtn.addEventListener('click', () => this.generateModel());
      console.log('Generate model button listener added');
    }
    if (validateModelBtn) {
      validateModelBtn.addEventListener('click', () => this.validateModel());
      console.log('Validate model button listener added');
    }
    if (sendChatBtn) {
      sendChatBtn.addEventListener('click', () => this.sendChatMessage());
      console.log('Send chat button listener added');
    }
    
    // Allow Enter key in chat input
    if (chatInput) {
      chatInput.addEventListener('keypress', (e) => {
        console.log('Key pressed in chat input:', e.key);
        if (e.key === 'Enter') {
          e.preventDefault();
          console.log('Enter pressed, sending message');
          this.sendChatMessage();
        }
      });
      
      // Also add keydown for better compatibility
      chatInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          console.log('Enter keydown, sending message');
          this.sendChatMessage();
        }
      });
      console.log('Chat input listeners added');
    }

    // File upload event listeners
    this.initializeFileUpload();

    // Collapsible sections
    this.initializeCollapsibleSections();

    this.isInitialized = true;
    console.log('MAModelingAddin initialized successfully');
    
    // Add a test message to verify chat is working
    setTimeout(() => {
      this.addChatMessage('assistant', '✅ Add-in loaded successfully! Try typing a message or clicking "browse files".');
      
      // Also test if Office.js is working
      if (typeof Office !== 'undefined' && Office.context) {
        this.addChatMessage('assistant', '📊 Excel integration ready! You can use all features.');
      } else {
        this.addChatMessage('assistant', '⚠️ Excel integration limited - some features may not work.');
      }
    }, 1500);
  }

  initializeFileUpload() {
    console.log('Initializing file upload...');
    const dropzone = document.getElementById('fileDropzone');
    const fileInput = document.getElementById('fileInput');
    const uploadLink = document.querySelector('.upload-link');

    console.log('Elements found:', {
      dropzone: !!dropzone,
      fileInput: !!fileInput,
      uploadLink: !!uploadLink
    });

    // Dropzone click handler
    if (dropzone) {
      dropzone.addEventListener('click', (e) => {
        console.log('Dropzone clicked');
        e.preventDefault();
        if (fileInput) {
          console.log('Triggering file input click');
          fileInput.click();
        } else {
          console.error('File input not found');
        }
      });
      console.log('Dropzone click listener added');
    }

    // Upload link click handler
    if (uploadLink) {
      uploadLink.addEventListener('click', (e) => {
        console.log('Upload link clicked');
        e.preventDefault();
        e.stopPropagation();
        if (fileInput) {
          console.log('Triggering file input click from link');
          fileInput.click();
        } else {
          console.error('File input not found');
        }
      });
      console.log('Upload link click listener added');
    }

    // File input change handler
    if (fileInput) {
      fileInput.addEventListener('change', (e) => {
        console.log('File input changed');
        const files = e.target.files;
        console.log('Files selected:', files ? files.length : 0);
        if (files && files.length > 0) {
          console.log('Processing files:', Array.from(files).map(f => f.name));
          this.handleFileSelection(Array.from(files));
        } else {
          console.log('No files selected');
        }
      });
      console.log('File input change listener added');
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
    console.log('Handling file selection:', files.length, 'files');
    
    // Filter valid files (PDF and CSV only)
    const validFiles = files.filter(file => {
      const isValidType = file.type === 'application/pdf' || file.type === 'text/csv' || file.name.endsWith('.csv');
      const isValidSize = file.size <= 10 * 1024 * 1024; // 10MB limit
      console.log(`File ${file.name}: type=${file.type}, size=${file.size}, valid=${isValidType && isValidSize}`);
      return isValidType && isValidSize;
    });

    console.log('Valid files:', validFiles.length);

    // Check total file limit
    if (this.uploadedFiles.length + validFiles.length > 4) {
      console.log('Too many files uploaded');
      this.addChatMessage('assistant', 'Maximum 4 files allowed. Please remove some files first.');
      return;
    }

    // Add files to uploaded list
    this.uploadedFiles.push(...validFiles);
    console.log('Total uploaded files:', this.uploadedFiles.length);
    this.updateFileDisplay();

    if (validFiles.length > 0) {
      console.log('Files uploaded successfully');
      this.addChatMessage('assistant', `Successfully uploaded ${validFiles.length} file(s). You can now ask me to analyze them and fill out your assumptions template!`);
    } else {
      console.log('No valid files to upload');
      this.addChatMessage('assistant', 'Please upload PDF or CSV files only (max 10MB each).');
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
    
    // Check if Excel is available
    if (typeof Excel === 'undefined') {
      console.error('Excel API not available');
      this.addChatMessage('assistant', '❌ Excel API not available. Please make sure you are running this in Excel.');
      return;
    }
    
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
        this.addChatMessage('assistant', `✅ Selected range ${range.address} with ${rangeData.length} rows.`);
      });
    } catch (error) {
      console.error('Error selecting range:', error);
      this.addChatMessage('assistant', `❌ Error selecting range: ${error.message || error}`);
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
    this.addChatMessage('assistant', 'Model generation feature is being loaded. Please use the chat to ask me to "create a blank assumptions template" for now.');
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
    console.log('Send chat message function called');
    const input = document.getElementById('chatInput');
    console.log('Chat input element:', input);
    
    if (!input) {
      console.error('Chat input element not found');
      return;
    }
    
    const message = input.value.trim();
    console.log('Message to send:', message);
    
    if (!message) {
      console.log('No message to send (empty)');
      return;
    }
    
    // Add user message
    console.log('Adding user message to chat');
    this.addChatMessage('user', message);
    input.value = '';
    
    // Process with AI
    console.log('Processing message with AI');
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

  async processUploadedFiles() {
    const fileContents = [];
    
    for (const file of this.uploadedFiles) {
      try {
        let content = '';
        
        if (file.type === 'text/csv' || file.name.endsWith('.csv')) {
          content = await this.readTextFile(file);
          // Limit CSV content to first 5000 characters to stay within token limits
          content = content.substring(0, 5000);
        } else if (file.type === 'application/pdf') {
          // For PDF files, we'll send the filename and indicate it needs server-side processing
          content = `[PDF FILE: ${file.name} - ${this.formatFileSize(file.size)}]`;
        }
        
        fileContents.push(`File: ${file.name}\nContent: ${content}`);
      } catch (error) {
        console.error(`Error processing file ${file.name}:`, error);
        fileContents.push(`File: ${file.name}\nError: Could not read file`);
      }
    }
    
    return fileContents;
  }

  readTextFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = reject;
      reader.readAsText(file);
    });
  }

  async executeCommand(command) {
    try {
      await Excel.run(async (context) => {
        switch (command.action) {
          case 'setValue':
            await this.setValueCommand(context, command.cell, command.value);
            break;
          case 'addToCell':
            await this.addToCellCommand(context, command.cell, command.value);
            break;
          case 'setFormula':
            await this.setFormulaCommand(context, command.cell, command.formula);
            break;
          case 'formatCell':
            await this.formatCellCommand(context, command.cell, command.format);
            break;
          case 'generateAssumptionsTemplate':
            await this.generateAssumptionsTemplate(context);
            break;
          case 'fillAssumptionsData':
            await this.fillAssumptionsData(context, command.data);
            break;
          default:
            console.warn('Unknown command action:', command.action);
        }
        
        await context.sync();
        
        if (command.action === 'generateAssumptionsTemplate') {
          this.addChatMessage('assistant', `✅ Generated M&A assumptions template`);
        } else if (command.action === 'fillAssumptionsData') {
          this.addChatMessage('assistant', `✅ Filled assumptions template with sample data`);
        } else {
          this.addChatMessage('assistant', `✅ Executed: ${command.action} on ${command.cell}`);
        }
      });
    } catch (error) {
      console.error('Command execution error:', error);
      this.addChatMessage('assistant', `❌ Error executing ${command.action}: ${error}`);
    }
  }

  async setValueCommand(context, cellAddress, value) {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(cellAddress);
    range.values = [[value]];
  }

  async setFormulaCommand(context, cellAddress, formula) {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(cellAddress);
    range.formulas = [[formula]];
  }

  async generateAssumptionsTemplate(context) {
    try {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Step 1: Add basic content first
      const templateData = [
        ["Sample Company Ltd. - Key Assumptions", ""],
        ["Deal type", ""],
        ["Sector/Industry", ""],
        ["Geography", ""],
        ["Business Model", ""],
        ["Ownership Structure", ""],
        ["Purchase Price ($M)", ""],
        ["", ""],
        ["Acquisition Assumptions", ""],
        ["Acquisition date", ""],
        ["Holding Period (Months)", ""],
        ["Currency", ""],
        ["Transaction Fees", ""],
        ["Acquisition LTV", ""],
        ["Equity Contribution", "=B7*(1-B14)"],
        ["Debt Financing", "=B7*B14"],
        ["Debt Issuance Fees", ""],
        ["Interest Rate Margin", ""],
        ["", ""],
        ["Cost Items", ""],
        ["Staff expenses", ""],
        ["Salary Growth (p.a.)", ""],
        ["Cost Item 1", ""],
        ["Cost Item 2", ""],
        ["Cost Item 3", ""],
        ["Cost Item 4", ""],
        ["Cost Item 5", ""],
        ["Cost Item 6", ""],
        ["", ""],
        ["Exit Assumptions", ""],
        ["Disposal Costs", ""],
        ["Terminal Equity Multiple", ""],
        ["Terminal EBITDA", ""],
        ["Sale Price", ""]
      ];
      
      // Add all data at once
      worksheet.getRange("A1:B34").values = templateData;
      await context.sync();
      
      // Step 2: Apply formatting
      // Set Times New Roman font for all cells
      worksheet.getRange("A1:B34").format.font.name = "Times New Roman";
      worksheet.getRange("A1:B34").format.font.size = 12;
      
      // Title formatting - merge cells and format
      worksheet.getRange("A1:B1").merge();
      worksheet.getRange("A1").format.font.bold = true;
      worksheet.getRange("A1").format.fill.color = "#D9D9D9";
      
      // Section headers - merge cells and format
      worksheet.getRange("A9:B9").merge();
      worksheet.getRange("A9").format.fill.color = "#1F4E79";
      worksheet.getRange("A9").format.font.color = "white";
      worksheet.getRange("A9").format.font.bold = true;
      
      worksheet.getRange("A20:B20").merge();
      worksheet.getRange("A20").format.fill.color = "#1F4E79";
      worksheet.getRange("A20").format.font.color = "white";
      worksheet.getRange("A20").format.font.bold = true;
      
      worksheet.getRange("A30:B30").merge();
      worksheet.getRange("A30").format.fill.color = "#1F4E79";
      worksheet.getRange("A30").format.font.color = "white";
      worksheet.getRange("A30").format.font.bold = true;
      
      await context.sync();
      
      // Step 3: Set column widths and row heights
      worksheet.getRange("A:A").format.columnWidth = 200;
      worksheet.getRange("B:B").format.columnWidth = 150;
      
      // Set row heights for empty rows before section headers
      worksheet.getRange("8:8").format.rowHeight = 5; // Before "Acquisition Assumptions"
      worksheet.getRange("19:19").format.rowHeight = 5; // Before "Cost Items"
      worksheet.getRange("29:29").format.rowHeight = 5; // Before "Exit Assumptions"
      
      await context.sync();
      
    } catch (error) {
      console.error("Template generation error:", error);
      throw error;
    }
  }

  async fillAssumptionsData(context, data) {
    try {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Map the data to the corresponding Excel cells
      const cellValues = [
        ["B2", data.dealType || "Business Acquisition"],
        ["B3", data.sector || "Technology"],
        ["B4", data.geography || "United States"],
        ["B5", data.businessModel || "SaaS"],
        ["B6", data.ownership || "Private Equity"],
        ["B7", data.purchasePrice || 100],
        ["B10", data.acquisitionDate || "31/03/2025"],
        ["B11", data.holdingPeriod || 60],
        ["B12", data.currency || "USD"],
        ["B13", data.transactionFees || "1.5%"],
        ["B14", data.acquisitionLTV || "75%"],
        ["B17", data.debtIssuanceFees || "1.0%"],
        ["B18", data.interestRateMargin || "3.5%"],
        ["B21", data.staffExpenses || 5000000],
        ["B22", data.salaryGrowth || "3.0%"],
        ["B23", data.costItem1 || 2000000],
        ["B24", data.costItem2 || 800000],
        ["B25", data.costItem3 || 1200000],
        ["B26", data.costItem4 || 400000],
        ["B27", data.costItem5 || 600000],
        ["B28", data.costItem6 || 300000],
        ["B31", data.disposalCosts || "0.5%"],
        ["B32", data.terminalEquityMultiple || 12.5],
        ["B33", data.terminalEBITDA || 15000000],
        ["B34", data.salePrice || 187500000]
      ];
      
      // Fill all data at once using bulk operations
      for (const [cell, value] of cellValues) {
        worksheet.getRange(cell).values = [[value]];
      }
      
      await context.sync();
      
    } catch (error) {
      console.error("Data filling error:", error);
      throw error;
    }
  }

  addChatMessage(role, content) {
    console.log(`Adding ${role} message:`, content);
    this.chatMessages.push({ role, content });
    
    const messagesDiv = document.getElementById('chatMessages');
    if (!messagesDiv) {
      console.error('Chat messages div not found');
      return;
    }
    
    console.log('Creating message bubble');
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
    
    // Scroll to bottom
    messagesDiv.scrollTop = messagesDiv.scrollHeight;
    console.log('Message added to chat');
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

  initializeCollapsibleSections() {
    console.log('Initializing collapsible sections...');
    
    // Add delay to ensure DOM is fully loaded
    setTimeout(() => {
      // Deal Assumptions section collapse/expand functionality
      const minimizeBtn = document.getElementById('minimizeAssumptions');
      const dealAssumptionsSection = document.getElementById('dealAssumptionsSection');
      
      console.log('DOM ready state:', document.readyState);
      console.log('Looking for elements:', {
        minimizeBtnExists: !!minimizeBtn,
        dealAssumptionsSectionExists: !!dealAssumptionsSection,
        minimizeBtnId: minimizeBtn ? minimizeBtn.id : 'not found',
        sectionId: dealAssumptionsSection ? dealAssumptionsSection.id : 'not found'
      });
      
      // Debug: List all elements with these IDs
      console.log('All elements with minimizeAssumptions ID:', document.querySelectorAll('#minimizeAssumptions'));
      console.log('All elements with dealAssumptionsSection ID:', document.querySelectorAll('#dealAssumptionsSection'));
      
      if (minimizeBtn && dealAssumptionsSection) {
        minimizeBtn.addEventListener('click', (e) => {
          e.preventDefault();
          console.log('Minimize button clicked');
          
          // Toggle collapsed class
          dealAssumptionsSection.classList.toggle('collapsed');
          
          // Update aria-label for accessibility
          const isCollapsed = dealAssumptionsSection.classList.contains('collapsed');
          minimizeBtn.setAttribute('aria-label', 
            isCollapsed ? 'Expand Deal Assumptions' : 'Minimize Deal Assumptions');
          
          console.log('Deal Assumptions section', isCollapsed ? 'collapsed' : 'expanded');
        });
        
        console.log('✅ Collapsible sections initialized successfully');
      } else {
        console.error('❌ Could not find collapsible section elements');
        console.log('Available elements in DOM:', {
          totalElements: document.querySelectorAll('*').length,
          sections: document.querySelectorAll('.section').length,
          buttons: document.querySelectorAll('button').length,
          bodyHTML: document.body ? document.body.innerHTML.substring(0, 500) + '...' : 'No body'
        });
      }
    }, 500); // 500ms delay to ensure DOM is ready
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
}

// Global error handler for better debugging
window.addEventListener('error', (e) => {
  console.error('Global error caught:', e.error, e.filename, e.lineno);
});

// Global unhandled promise rejection handler
window.addEventListener('unhandledrejection', (e) => {
  console.error('Unhandled promise rejection:', e.reason);
});

// Initialize the add-in
console.log('Initializing MAModelingAddin...');
console.log('Office availability:', typeof Office !== 'undefined');
console.log('Excel availability:', typeof Excel !== 'undefined');

// Wait for everything to load properly
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    console.log('DOM loaded, creating MAModelingAddin');
    new MAModelingAddin();
  });
} else {
  console.log('DOM already loaded, creating MAModelingAddin');
  new MAModelingAddin();
}