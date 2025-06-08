class FileUploader {
  constructor() {
    this.uploadedFiles = [];
    this.mainUploadedFiles = [];
    this.isInitialized = false;
  }

  initialize() {
    if (this.isInitialized) return;
    
    console.log('Initializing file upload system...');
    
    // Get upload system elements
    const mainUploadZone = document.getElementById('mainUploadZone');
    const mainFileInput = document.getElementById('mainFileInput');
    const browseFilesBtn = document.getElementById('browseFilesBtn');
    const autoFillBtn = document.getElementById('autoFillBtn');
    const uploadedFilesDisplay = document.getElementById('uploadedFilesDisplay');
    const filesGrid = document.getElementById('filesGrid');

    console.log('Upload elements found:', {
      mainUploadZone: !!mainUploadZone,
      mainFileInput: !!mainFileInput,
      browseFilesBtn: !!browseFilesBtn,
      autoFillBtn: !!autoFillBtn,
      uploadedFilesDisplay: !!uploadedFilesDisplay,
      filesGrid: !!filesGrid
    });

    // Check if we're in Excel Online
    const isExcelOnline = window.location.hostname.includes('officeapps.live.com') || 
                         window.location.hostname.includes('excel.officeapps.live.com') ||
                         window.location.hostname.includes('excel.cloud.microsoft');
    
    if (isExcelOnline) {
      console.log('⚠️ Running in Excel Online - using alternative file handling');
    }

    // Main upload zone click handler
    if (mainUploadZone) {
      mainUploadZone.addEventListener('click', (e) => {
        console.log('Main upload zone clicked');
        if (e.target !== mainFileInput) {
          e.preventDefault();
        }
        if (mainFileInput) {
          console.log('Triggering main file input click');
          try {
            if (mainFileInput.click) {
              mainFileInput.click();
            } else {
              const evt = new MouseEvent('click', {
                bubbles: true,
                cancelable: true,
                view: window
              });
              mainFileInput.dispatchEvent(evt);
            }
          } catch (error) {
            console.error('Error triggering file input:', error);
            this.showUploadMessage('Please use the Browse Files button or drag and drop files.', 'info');
          }
        }
      });
    }

    // Browse files button click handler
    if (browseFilesBtn) {
      browseFilesBtn.addEventListener('click', (e) => {
        console.log('Browse files button clicked');
        e.preventDefault();
        e.stopPropagation();
        if (mainFileInput) {
          console.log('Triggering file input from browse button');
          try {
            if (mainFileInput.click) {
              mainFileInput.click();
            } else {
              const evt = new MouseEvent('click', {
                bubbles: true,
                cancelable: true,
                view: window
              });
              mainFileInput.dispatchEvent(evt);
            }
          } catch (error) {
            console.error('Error triggering file input from button:', error);
            this.showUploadMessage('File upload may be restricted in Excel Online. Try drag and drop.', 'info');
          }
        }
      });
    }

    // Main file input change handler
    if (mainFileInput) {
      mainFileInput.addEventListener('change', (e) => {
        console.log('Main file input changed');
        const files = e.target.files;
        console.log('Files selected:', files ? files.length : 0);
        if (files && files.length > 0) {
          console.log('Processing main files:', Array.from(files).map(f => f.name));
          this.handleFileSelection(Array.from(files));
        }
        e.target.value = '';
      });
    }

    // Drag and drop handlers for main upload zone
    if (mainUploadZone) {
      mainUploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        mainUploadZone.classList.add('dragover');
      });

      mainUploadZone.addEventListener('dragleave', (e) => {
        if (!mainUploadZone.contains(e.relatedTarget)) {
          mainUploadZone.classList.remove('dragover');
        }
      });

      mainUploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        mainUploadZone.classList.remove('dragover');
        const files = Array.from(e.dataTransfer.files || []);
        console.log('Files dropped:', files.length);
        this.handleFileSelection(files);
      });
    }

    // Auto Fill button handler
    if (autoFillBtn) {
      autoFillBtn.addEventListener('click', () => {
        this.processAutoFill();
      });
      autoFillBtn.disabled = true;
    }

    this.isInitialized = true;
    console.log('✅ File upload system initialized');
  }

  handleFileSelection(files) {
    console.log('Handling file selection:', files.length, 'files');
    
    // Filter valid files (PDF, CSV, PNG, JPG)
    const validFiles = files.filter(file => {
      const validTypes = [
        'application/pdf',
        'text/csv',
        'image/png',
        'image/jpeg',
        'image/jpg'
      ];
      const isValidType = validTypes.includes(file.type) || 
                         file.name.endsWith('.csv') || 
                         file.name.endsWith('.pdf') ||
                         file.name.endsWith('.png') ||
                         file.name.endsWith('.jpg') ||
                         file.name.endsWith('.jpeg');
      const isValidSize = file.size <= 10 * 1024 * 1024; // 10MB limit
      console.log(`File ${file.name}: type=${file.type}, size=${file.size}, valid=${isValidType && isValidSize}`);
      return isValidType && isValidSize;
    });

    console.log('Valid files:', validFiles.length);

    // Check total file limit
    if (this.mainUploadedFiles.length + validFiles.length > 4) {
      console.log('Too many files uploaded');
      this.showUploadMessage('Maximum 4 files allowed. Please remove some files first.', 'error');
      return;
    }

    // Add files to uploaded list
    this.mainUploadedFiles.push(...validFiles);
    console.log('Total uploaded files:', this.mainUploadedFiles.length);
    this.updateFileDisplay();

    if (validFiles.length > 0) {
      console.log('Files uploaded successfully');
      this.showUploadMessage(`Successfully uploaded ${validFiles.length} file(s). Ready for auto-fill!`, 'success');
      
      // Enable auto-fill button
      const autoFillBtn = document.getElementById('autoFillBtn');
      if (autoFillBtn) {
        autoFillBtn.disabled = false;
      }
    } else {
      console.log('No valid files to upload');
      this.showUploadMessage('Please upload PDF, CSV, PNG, or JPG files only (max 10MB each).', 'error');
    }
  }

  updateFileDisplay() {
    const uploadedFilesDisplay = document.getElementById('uploadedFilesDisplay');
    const filesGrid = document.getElementById('filesGrid');

    if (this.mainUploadedFiles.length === 0) {
      if (uploadedFilesDisplay) uploadedFilesDisplay.style.display = 'none';
      return;
    }

    if (uploadedFilesDisplay) uploadedFilesDisplay.style.display = 'block';
    if (filesGrid) filesGrid.innerHTML = '';

    this.mainUploadedFiles.forEach((file, index) => {
      const fileCard = document.createElement('div');
      fileCard.className = 'file-card';
      
      // Determine file icon based on type
      let iconSVG = '';
      if (file.type.includes('pdf')) {
        iconSVG = `<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14,2 14,8 20,8"></polyline>`;
      } else if (file.type.includes('csv') || file.name.endsWith('.csv')) {
        iconSVG = `<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14,2 14,8 20,8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line>`;
      } else if (file.type.includes('image')) {
        iconSVG = `<rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect><circle cx="8.5" cy="8.5" r="1.5"></circle><polyline points="21,15 16,10 5,21"></polyline>`;
      } else {
        iconSVG = `<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14,2 14,8 20,8"></polyline>`;
      }
      
      fileCard.innerHTML = `
        <svg class="file-card-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          ${iconSVG}
        </svg>
        <div class="file-card-info">
          <div class="file-card-name">${file.name}</div>
          <div class="file-card-size">${this.formatFileSize(file.size)}</div>
        </div>
        <button class="file-card-remove" onclick="window.fileUploader.removeFile(${index})" title="Remove file">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <line x1="18" y1="6" x2="6" y2="18"></line>
            <line x1="6" y1="6" x2="18" y2="18"></line>
          </svg>
        </button>
      `;

      if (filesGrid) {
        filesGrid.appendChild(fileCard);
      }
    });
  }

  removeFile(index) {
    console.log('Removing file at index:', index);
    this.mainUploadedFiles.splice(index, 1);
    this.updateFileDisplay();

    // Disable auto-fill button if no files
    if (this.mainUploadedFiles.length === 0) {
      const autoFillBtn = document.getElementById('autoFillBtn');
      if (autoFillBtn) {
        autoFillBtn.disabled = true;
      }
    }

    this.showUploadMessage('File removed successfully.', 'info');
  }

  formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  showUploadMessage(message, type = 'info') {
    // Create or update a message display element
    let messageElement = document.getElementById('uploadMessage');
    if (!messageElement) {
      messageElement = document.createElement('div');
      messageElement.id = 'uploadMessage';
      messageElement.style.cssText = `
        padding: 12px;
        margin: 10px 0;
        border-radius: 6px;
        font-size: 14px;
        font-weight: 500;
        text-align: center;
        transition: all 0.3s ease;
      `;
      
      const uploadZone = document.getElementById('mainUploadZone');
      if (uploadZone && uploadZone.parentNode) {
        uploadZone.parentNode.insertBefore(messageElement, uploadZone.nextSibling);
      }
    }

    // Set message and styling based on type
    messageElement.textContent = message;
    
    switch (type) {
      case 'success':
        messageElement.style.backgroundColor = '#10B981';
        messageElement.style.color = 'white';
        break;
      case 'error':
        messageElement.style.backgroundColor = '#EF4444';
        messageElement.style.color = 'white';
        break;
      case 'info':
      default:
        messageElement.style.backgroundColor = '#F3F4F6';
        messageElement.style.color = '#374151';
        break;
    }

    // Auto-hide after 5 seconds
    setTimeout(() => {
      if (messageElement) {
        messageElement.style.opacity = '0';
        setTimeout(() => {
          if (messageElement && messageElement.parentNode) {
            messageElement.parentNode.removeChild(messageElement);
          }
        }, 300);
      }
    }, 5000);
  }

  async processAutoFill() {
    console.log('🤖 Processing auto-fill...');
    
    if (!this.mainUploadedFiles || this.mainUploadedFiles.length === 0) {
      this.showUploadMessage('No files uploaded for processing.', 'error');
      return;
    }

    try {
      // Show loading state
      const autoFillBtn = document.getElementById('autoFillBtn');
      if (autoFillBtn) {
        autoFillBtn.disabled = true;
        autoFillBtn.innerHTML = `
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M21 12a9 9 0 11-6.219-8.56"></path>
          </svg>
          Processing...
        `;
      }

      this.showUploadMessage('Processing files with AI...', 'info');

      // Process files and extract data
      const extractedData = await this.processUploadedFiles();
      
      if (extractedData) {
        // Apply extracted data to form
        await this.applyExtractedData(extractedData);
        this.showUploadMessage('Auto-fill completed successfully!', 'success');
      } else {
        this.showUploadMessage('Could not extract data from files. Please try manually filling the form.', 'error');
      }

    } catch (error) {
      console.error('Error during auto-fill:', error);
      this.showUploadMessage('Error during auto-fill. Please try again.', 'error');
    } finally {
      // Reset button state
      const autoFillBtn = document.getElementById('autoFillBtn');
      if (autoFillBtn) {
        autoFillBtn.disabled = false;
        autoFillBtn.innerHTML = `
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M9 12l2 2 4-4"></path>
            <circle cx="12" cy="12" r="9"></circle>
          </svg>
          Auto Fill with AI
        `;
      }
    }
  }

  async processUploadedFiles() {
    console.log('Processing uploaded files for data extraction...');
    
    const fileContents = [];
    
    // Read all files
    for (const file of this.mainUploadedFiles) {
      try {
        console.log(`Reading file: ${file.name}`);
        let content = '';
        
        if (file.type === 'text/csv' || file.name.endsWith('.csv')) {
          content = await this.readTextFile(file);
        } else if (file.type === 'application/pdf') {
          content = `PDF file: ${file.name} (${this.formatFileSize(file.size)})`;
        } else if (file.type.startsWith('image/')) {
          content = `Image file: ${file.name} (${this.formatFileSize(file.size)})`;
        }
        
        fileContents.push({
          name: file.name,
          type: file.type,
          content: content,
          size: file.size
        });
        
      } catch (error) {
        console.error(`Error reading file ${file.name}:`, error);
      }
    }

    // Create mock extracted data for demonstration
    // In a real implementation, this would call an AI service
    return this.createMockExtractedData(fileContents);
  }

  async readTextFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        resolve(e.target.result);
      };
      reader.onerror = (e) => {
        reject(new Error(`Failed to read file: ${file.name}`));
      };
      reader.readAsText(file);
    });
  }

  createMockExtractedData(fileContents) {
    // This is a mock implementation
    // In reality, you would send the file contents to an AI service
    console.log('Creating mock extracted data from files:', fileContents.map(f => f.name));
    
    return {
      // High-Level Parameters
      currency: 'USD',
      projectStartDate: '2024-01-01',
      projectEndDate: '2027-01-01',
      modelPeriods: 'monthly',
      
      // Deal Assumptions
      dealName: 'Sample Tech Acquisition',
      dealValue: 50000000,
      transactionFee: 2.5,
      dealLTV: 70,
      
      // Revenue Items
      revenueItems: [
        {
          name: 'Software Licenses',
          value: 2000000,
          growthType: 'annual',
          annualGrowthRate: 15
        },
        {
          name: 'Consulting Services',
          value: 800000,
          growthType: 'linear',
          linearGrowthRate: 10
        }
      ],
      
      // Operating Expenses
      operatingExpenses: [
        {
          name: 'Staff Costs',
          value: 1200000,
          growthType: 'annual',
          annualGrowthRate: 5
        },
        {
          name: 'Marketing',
          value: 300000,
          growthType: 'linear',
          linearGrowthRate: 8
        }
      ],
      
      // Capital Expenses
      capitalExpenses: [
        {
          name: 'IT Infrastructure',
          value: 500000,
          growthType: 'linear',
          linearGrowthRate: 3
        }
      ],
      
      // Exit Assumptions
      disposalCost: 2.5,
      terminalCapRate: 8.5
    };
  }

  async applyExtractedData(data) {
    console.log('Applying extracted data to form...');
    
    // Apply high-level parameters
    if (data.currency) this.setInputValue('currency', data.currency);
    if (data.projectStartDate) this.setInputValue('projectStartDate', data.projectStartDate);
    if (data.projectEndDate) this.setInputValue('projectEndDate', data.projectEndDate);
    if (data.modelPeriods) this.setInputValue('modelPeriods', data.modelPeriods);
    
    // Apply deal assumptions
    if (data.dealName) this.setInputValue('dealName', data.dealName);
    if (data.dealValue) this.setInputValue('dealValue', data.dealValue);
    if (data.transactionFee) this.setInputValue('transactionFee', data.transactionFee);
    if (data.dealLTV) this.setInputValue('dealLTV', data.dealLTV);
    
    // Apply exit assumptions
    if (data.disposalCost) this.setInputValue('disposalCost', data.disposalCost);
    if (data.terminalCapRate) this.setInputValue('terminalCapRate', data.terminalCapRate);
    
    // Apply revenue items
    if (data.revenueItems && data.revenueItems.length > 0) {
      await this.applyRevenueItems(data.revenueItems);
    }
    
    // Apply operating expenses
    if (data.operatingExpenses && data.operatingExpenses.length > 0) {
      await this.applyOperatingExpenses(data.operatingExpenses);
    }
    
    // Apply capital expenses
    if (data.capitalExpenses && data.capitalExpenses.length > 0) {
      await this.applyCapitalExpenses(data.capitalExpenses);
    }
    
    // Trigger calculations
    if (window.formHandler) {
      window.formHandler.triggerCalculations();
    }
    
    console.log('Data application completed');
  }

  async applyRevenueItems(items) {
    console.log('Applying revenue items:', items.length);
    
    // Clear existing revenue items
    const container = document.getElementById('revenueItemsContainer');
    if (container) {
      container.innerHTML = '';
    }
    
    // Add new revenue items
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      
      // Add revenue item using FormHandler
      if (window.formHandler) {
        window.formHandler.addRevenueItem();
        
        // Wait a bit for DOM update
        await new Promise(resolve => setTimeout(resolve, 100));
        
        // Fill in the data
        this.setInputValue(`revenueName_${i + 1}`, item.name);
        this.setInputValue(`revenueValue_${i + 1}`, item.value);
        
        if (item.growthType) {
          this.setInputValue(`growthType_${i + 1}`, item.growthType);
          
          // Update growth inputs and fill data
          if (window.formHandler) {
            window.formHandler.updateGrowthInputs(`revenue_${i + 1}`, item.growthType);
            
            await new Promise(resolve => setTimeout(resolve, 100));
            
            if (item.growthType === 'annual' && item.annualGrowthRate) {
              this.setInputValue(`annualGrowth_${i + 1}`, item.annualGrowthRate);
            } else if (item.growthType === 'linear' && item.linearGrowthRate) {
              this.setInputValue(`linearGrowth_${i + 1}`, item.linearGrowthRate);
            }
          }
        }
      }
    }
  }

  async applyOperatingExpenses(items) {
    console.log('Applying operating expenses:', items.length);
    
    // Clear existing items
    const container = document.getElementById('operatingExpensesContainer');
    if (container) {
      container.innerHTML = '';
    }
    
    // Add new items
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      
      if (window.formHandler) {
        window.formHandler.addOperatingExpense();
        
        await new Promise(resolve => setTimeout(resolve, 100));
        
        this.setInputValue(`opExName_${i + 1}`, item.name);
        this.setInputValue(`opExValue_${i + 1}`, item.value);
        
        if (item.growthType) {
          this.setInputValue(`opExGrowthType_${i + 1}`, item.growthType);
          
          if (window.formHandler) {
            window.formHandler.updateCostGrowthInputs(`opEx_${i + 1}`, item.growthType);
            
            await new Promise(resolve => setTimeout(resolve, 100));
            
            if (item.growthType === 'annual' && item.annualGrowthRate) {
              this.setInputValue(`annualGrowth_opEx_${i + 1}`, item.annualGrowthRate);
            } else if (item.growthType === 'linear' && item.linearGrowthRate) {
              this.setInputValue(`linearGrowth_opEx_${i + 1}`, item.linearGrowthRate);
            }
          }
        }
      }
    }
  }

  async applyCapitalExpenses(items) {
    console.log('Applying capital expenses:', items.length);
    
    // Clear existing items
    const container = document.getElementById('capitalExpensesContainer');
    if (container) {
      container.innerHTML = '';
    }
    
    // Add new items
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      
      if (window.formHandler) {
        window.formHandler.addCapitalExpense();
        
        await new Promise(resolve => setTimeout(resolve, 100));
        
        this.setInputValue(`capExName_${i + 1}`, item.name);
        this.setInputValue(`capExValue_${i + 1}`, item.value);
        
        if (item.growthType) {
          this.setInputValue(`capExGrowthType_${i + 1}`, item.growthType);
          
          if (window.formHandler) {
            window.formHandler.updateCostGrowthInputs(`capEx_${i + 1}`, item.growthType);
            
            await new Promise(resolve => setTimeout(resolve, 100));
            
            if (item.growthType === 'annual' && item.annualGrowthRate) {
              this.setInputValue(`annualGrowth_capEx_${i + 1}`, item.annualGrowthRate);
            } else if (item.growthType === 'linear' && item.linearGrowthRate) {
              this.setInputValue(`linearGrowth_capEx_${i + 1}`, item.linearGrowthRate);
            }
          }
        }
      }
    }
  }

  setInputValue(elementId, value) {
    const element = document.getElementById(elementId);
    if (element && value !== null && value !== undefined) {
      element.value = value;
      element.dispatchEvent(new Event('change', { bubbles: true }));
      element.dispatchEvent(new Event('input', { bubbles: true }));
    }
  }

  getUploadedFiles() {
    return this.mainUploadedFiles;
  }

  clearAllFiles() {
    this.mainUploadedFiles = [];
    this.updateFileDisplay();
    
    // Disable auto-fill button
    const autoFillBtn = document.getElementById('autoFillBtn');
    if (autoFillBtn) {
      autoFillBtn.disabled = true;
    }
  }
}

// Export for use in main application
window.FileUploader = FileUploader;