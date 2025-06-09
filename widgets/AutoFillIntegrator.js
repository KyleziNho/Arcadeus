/**
 * AutoFillIntegrator.js - Integrates the new AI extraction system with the existing UI
 * This is the main orchestrator that connects all the new extraction widgets with the form
 */

class AutoFillIntegrator {
  constructor() {
    this.isInitialized = false;
    this.isProcessing = false;
    this.uploadedFiles = [];
    
    // Initialize all extraction components
    this.aiExtractionService = null;
    this.dataStandardizer = null;
    this.fieldMappingEngine = null;
    this.confidenceIndicator = null;
    this.reviewModal = null;
    this.extractionHistory = null;
    
    // Extraction widgets
    this.highLevelExtractor = null;
    this.dealAssumptionsExtractor = null;
    this.revenueItemsExtractor = null;
    this.costItemsExtractor = null;
    this.debtModelExtractor = null;
    this.exitAssumptionsExtractor = null;
  }

  async initialize() {
    if (this.isInitialized) return;
    
    console.log('🚀 Initializing AutoFillIntegrator...');
    
    try {
      // Initialize core services
      await this.initializeCoreServices();
      
      // Initialize extraction widgets
      await this.initializeExtractionWidgets();
      
      // Initialize UI components
      await this.initializeUIComponents();
      
      // Setup event handlers
      this.setupEventHandlers();
      
      this.isInitialized = true;
      console.log('✅ AutoFillIntegrator initialized successfully');
      
      // Show integration status to user
      this.showIntegrationStatus();
      
    } catch (error) {
      console.error('❌ Failed to initialize AutoFillIntegrator:', error);
      this.showError('Failed to initialize AI extraction system. Some features may not work properly.');
    }
  }

  async initializeCoreServices() {
    console.log('🔧 Initializing core services...');
    
    // Initialize AI extraction service
    if (window.AIExtractionService) {
      this.aiExtractionService = new window.AIExtractionService();
      console.log('✅ AIExtractionService initialized');
    }
    
    // Initialize data standardizer
    if (window.DataStandardizer) {
      this.dataStandardizer = new window.DataStandardizer();
      console.log('✅ DataStandardizer initialized');
    }
    
    // Initialize field mapping engine
    if (window.FieldMappingEngine) {
      this.fieldMappingEngine = new window.FieldMappingEngine();
      console.log('✅ FieldMappingEngine initialized');
    }
    
    // Initialize extraction history
    if (window.ExtractionHistory) {
      this.extractionHistory = new window.ExtractionHistory();
      console.log('✅ ExtractionHistory initialized');
    }
  }

  async initializeExtractionWidgets() {
    console.log('🧩 Initializing extraction widgets...');
    
    const services = {
      extractionService: this.aiExtractionService,
      standardizer: this.dataStandardizer,
      mappingEngine: this.fieldMappingEngine
    };
    
    console.log('🧩 Services available:', {
      extractionService: !!this.aiExtractionService,
      standardizer: !!this.dataStandardizer,
      mappingEngine: !!this.fieldMappingEngine
    });
    
    // Initialize each extraction widget with error handling
    try {
      if (window.HighLevelParametersExtractor) {
        this.highLevelExtractor = new window.HighLevelParametersExtractor();
        this.highLevelExtractor.initialize(services);
        console.log('✅ HighLevelParametersExtractor ready, extract method:', typeof this.highLevelExtractor.extract);
      } else {
        console.warn('❌ HighLevelParametersExtractor not found on window');
      }
    } catch (error) {
      console.error('❌ Error initializing HighLevelParametersExtractor:', error);
    }
    
    try {
      if (window.DealAssumptionsExtractor) {
        this.dealAssumptionsExtractor = new window.DealAssumptionsExtractor();
        this.dealAssumptionsExtractor.initialize(services);
        console.log('✅ DealAssumptionsExtractor ready, extract method:', typeof this.dealAssumptionsExtractor.extract);
      } else {
        console.warn('❌ DealAssumptionsExtractor not found on window');
      }
    } catch (error) {
      console.error('❌ Error initializing DealAssumptionsExtractor:', error);
    }
    
    try {
      if (window.RevenueItemsExtractor) {
        this.revenueItemsExtractor = new window.RevenueItemsExtractor();
        this.revenueItemsExtractor.initialize(services);
        console.log('✅ RevenueItemsExtractor ready, extract method:', typeof this.revenueItemsExtractor.extract);
      } else {
        console.warn('❌ RevenueItemsExtractor not found on window');
      }
    } catch (error) {
      console.error('❌ Error initializing RevenueItemsExtractor:', error);
    }
    
    try {
      if (window.CostItemsExtractor) {
        this.costItemsExtractor = new window.CostItemsExtractor();
        this.costItemsExtractor.initialize(services);
        console.log('✅ CostItemsExtractor ready, extract method:', typeof this.costItemsExtractor.extract);
      } else {
        console.warn('❌ CostItemsExtractor not found on window');
      }
    } catch (error) {
      console.error('❌ Error initializing CostItemsExtractor:', error);
    }
    
    try {
      if (window.DebtModelExtractor) {
        this.debtModelExtractor = new window.DebtModelExtractor();
        this.debtModelExtractor.initialize(services);
        console.log('✅ DebtModelExtractor ready, extract method:', typeof this.debtModelExtractor.extract);
      } else {
        console.warn('❌ DebtModelExtractor not found on window');
      }
    } catch (error) {
      console.error('❌ Error initializing DebtModelExtractor:', error);
    }
    
    try {
      if (window.ExitAssumptionsExtractor) {
        this.exitAssumptionsExtractor = new window.ExitAssumptionsExtractor();
        this.exitAssumptionsExtractor.initialize(services);
        console.log('✅ ExitAssumptionsExtractor ready, extract method:', typeof this.exitAssumptionsExtractor.extract);
      } else {
        console.warn('❌ ExitAssumptionsExtractor not found on window');
      }
    } catch (error) {
      console.error('❌ Error initializing ExitAssumptionsExtractor:', error);
    }
    
    console.log('✅ All extraction widgets initialization completed');
  }

  async initializeUIComponents() {
    console.log('🎨 Initializing UI components...');
    
    // Initialize confidence indicator
    if (window.ExtractionConfidenceIndicator) {
      this.confidenceIndicator = new window.ExtractionConfidenceIndicator();
      console.log('✅ ExtractionConfidenceIndicator initialized');
    }
    
    // Initialize review modal
    if (window.ExtractionReviewModal) {
      this.reviewModal = new window.ExtractionReviewModal();
      console.log('✅ ExtractionReviewModal initialized');
    }
  }

  setupEventHandlers() {
    console.log('🔗 Setting up event handlers...');
    
    // Hook into the existing autofill button
    const autoFillBtn = document.getElementById('autoFillBtn');
    if (autoFillBtn) {
      // Remove existing handlers and add our new one
      autoFillBtn.replaceWith(autoFillBtn.cloneNode(true));
      const newAutoFillBtn = document.getElementById('autoFillBtn');
      newAutoFillBtn.addEventListener('click', () => this.handleAutoFill());
      console.log('✅ AutoFill button handler attached');
    }
    
    // Hook into file upload handlers
    if (window.fileUploader) {
      // Get current uploaded files
      this.uploadedFiles = window.fileUploader.getUploadedFiles() || [];
      console.log('✅ File upload integration ready, current files:', this.uploadedFiles.length);
      
      // Initialize autofill button state
      this.onFilesChanged(this.uploadedFiles);
    }
    
    // Add extraction history listener
    if (this.extractionHistory) {
      this.extractionHistory.onHistoryChange((event) => {
        this.updateHistoryUI(event.detail);
      });
    }
  }

  async handleAutoFill() {
    if (this.isProcessing) {
      this.showError('Extraction already in progress. Please wait...');
      return;
    }
    
    if (!this.uploadedFiles || this.uploadedFiles.length === 0) {
      this.showError('Please upload files first before using AI autofill.');
      return;
    }
    
    console.log('🤖 Starting comprehensive AI extraction...');
    this.isProcessing = true;
    
    try {
      // Show loading state
      this.showLoadingState(true);
      
      // Step 1: Read file contents
      console.log('📖 Step 1: Reading file contents...');
      const filesWithContent = await this.readFileContents();
      
      // Step 2: Extract data from all widgets
      console.log('📊 Step 2: Extracting data from all sections...');
      const extractionResults = await this.extractAllData(filesWithContent);
      
      // Step 3: Show review modal
      console.log('📊 Step 3: Showing review modal...');
      await this.showReviewModal(extractionResults);
      
    } catch (error) {
      console.error('❌ Extraction failed:', error);
      this.showError('AI extraction failed. Please check the console for details and try again.');
    } finally {
      this.isProcessing = false;
      this.showLoadingState(false);
    }
  }

  async readFileContents() {
    console.log('📖 Reading file contents for AI extraction...');
    
    if (!window.fileUploader || !window.fileUploader.readAllFiles) {
      throw new Error('FileUploader not available or readAllFiles method missing');
    }
    
    const filesWithContent = await window.fileUploader.readAllFiles();
    console.log('📖 File contents read:', filesWithContent.map(f => `${f.name} (${f.content ? f.content.length : 0} chars)`));
    
    return filesWithContent;
  }

  async extractAllData(filesWithContent) {
    const startTime = Date.now();
    const extractionResults = {};
    const allExtractedData = {};
    
    // Extract from each section concurrently for better performance
    const extractionPromises = [];
    
    if (this.highLevelExtractor) {
      console.log('🎯 Starting high-level parameters extraction...');
      extractionPromises.push(
        this.highLevelExtractor.extract(filesWithContent)
          .then(data => {
            console.log('🎯 High-level parameters extraction completed:', data);
            extractionResults.highLevelParameters = data;
            Object.assign(allExtractedData, data);
            this.showProgress('High-level parameters extracted');
          })
          .catch(error => {
            console.error('🎯 High-level parameters extraction failed:', error);
            this.showProgress('High-level parameters extraction failed');
          })
      );
    } else {
      console.warn('🎯 High-level parameters extractor not available');
    }
    
    if (this.dealAssumptionsExtractor) {
      extractionPromises.push(
        this.dealAssumptionsExtractor.extract(filesWithContent)
          .then(data => {
            extractionResults.dealAssumptions = data;
            Object.assign(allExtractedData, data);
            this.showProgress('Deal assumptions extracted');
          })
      );
    }
    
    if (this.revenueItemsExtractor) {
      extractionPromises.push(
        this.revenueItemsExtractor.extract(filesWithContent)
          .then(data => {
            extractionResults.revenueItems = data;
            Object.assign(allExtractedData, data);
            this.showProgress('Revenue items extracted');
          })
      );
    }
    
    if (this.costItemsExtractor) {
      extractionPromises.push(
        this.costItemsExtractor.extract(filesWithContent)
          .then(data => {
            extractionResults.costItems = data;
            Object.assign(allExtractedData, data);
            this.showProgress('Cost items extracted');
          })
      );
    }
    
    if (this.debtModelExtractor) {
      extractionPromises.push(
        this.debtModelExtractor.extract(filesWithContent)
          .then(data => {
            extractionResults.debtModel = data;
            Object.assign(allExtractedData, data);
            this.showProgress('Debt model extracted');
          })
      );
    }
    
    if (this.exitAssumptionsExtractor) {
      extractionPromises.push(
        this.exitAssumptionsExtractor.extract(filesWithContent)
          .then(data => {
            extractionResults.exitAssumptions = data;
            Object.assign(allExtractedData, data);
            this.showProgress('Exit assumptions extracted');
          })
      );
    }
    
    // Wait for all extractions to complete
    await Promise.all(extractionPromises);
    
    const duration = Date.now() - startTime;
    console.log(`✅ All extractions completed in ${duration}ms`);
    
    // Save to extraction history
    if (this.extractionHistory) {
      this.extractionHistory.saveSession({
        files: this.uploadedFiles,
        extractedData: allExtractedData,
        extractionResults: extractionResults,
        duration: duration,
        method: 'comprehensive_ai'
      });
    }
    
    return allExtractedData;
  }

  async showReviewModal(extractedData) {
    if (!this.reviewModal) {
      // If no review modal, apply directly
      await this.applyExtractedData(extractedData);
      return;
    }
    
    return new Promise((resolve) => {
      this.reviewModal.show(extractedData, {
        title: 'Review AI Extracted Data',
        subtitle: 'Review and edit the extracted data before applying to your model',
        onApprove: async (finalData) => {
          console.log('✅ User approved extraction data');
          await this.applyExtractedData(finalData);
          resolve();
        },
        onReject: () => {
          console.log('❌ User rejected extraction data');
          this.showInfo('Extraction cancelled by user');
          resolve();
        },
        confidenceIndicator: this.confidenceIndicator
      });
    });
  }

  async applyExtractedData(extractedData) {
    console.log('📝 Applying extracted data to form...');
    
    try {
      // Apply to each section
      await this.applyHighLevelParameters(extractedData);
      await this.applyDealAssumptions(extractedData);
      await this.applyRevenueItems(extractedData);
      await this.applyCostItems(extractedData);
      await this.applyDebtModel(extractedData);
      await this.applyExitAssumptions(extractedData);
      
      // Trigger form calculations
      if (window.formHandler) {
        window.formHandler.triggerCalculations();
      }
      
      // Show success message
      this.showSuccess('AI extraction completed successfully! Data has been applied to all sections.');
      
      // Update history with applied data
      if (this.extractionHistory) {
        const currentSession = this.extractionHistory.getCurrentSession();
        if (currentSession) {
          currentSession.appliedData = extractedData;
          currentSession.appliedAt = new Date().toISOString();
        }
      }
      
    } catch (error) {
      console.error('❌ Error applying extracted data:', error);
      this.showError('Failed to apply some extracted data. Please check the console and apply manually if needed.');
    }
  }

  async applyHighLevelParameters(data) {
    if (!data) return;
    
    // Apply high-level parameters
    if (data.currency?.value) {
      this.setFieldValue('currency', data.currency.value, data.currency);
    }
    if (data.projectStartDate?.value) {
      this.setFieldValue('projectStartDate', data.projectStartDate.value, data.projectStartDate);
    }
    if (data.projectEndDate?.value) {
      this.setFieldValue('projectEndDate', data.projectEndDate.value, data.projectEndDate);
    }
    if (data.modelPeriods?.value) {
      this.setFieldValue('modelPeriods', data.modelPeriods.value, data.modelPeriods);
    }
  }

  async applyDealAssumptions(data) {
    if (!data) return;
    
    // Apply deal assumptions
    if (data.dealName?.value) {
      this.setFieldValue('dealName', data.dealName.value, data.dealName);
    }
    if (data.dealValue?.value) {
      this.setFieldValue('dealValue', data.dealValue.value, data.dealValue);
    }
    if (data.transactionFee?.value) {
      this.setFieldValue('transactionFee', data.transactionFee.value, data.transactionFee);
    }
    if (data.dealLTV?.value) {
      this.setFieldValue('dealLTV', data.dealLTV.value, data.dealLTV);
    }
  }

  async applyRevenueItems(data) {
    if (!data.revenueItems?.value || !Array.isArray(data.revenueItems.value)) return;
    
    console.log('💰 Applying revenue items:', data.revenueItems.value);
    
    // Clear existing revenue items
    const container = document.getElementById('revenueItemsContainer');
    if (container) {
      container.innerHTML = '';
    }
    
    // Add each revenue item
    for (let i = 0; i < data.revenueItems.value.length; i++) {
      const item = data.revenueItems.value[i];
      
      // Add new revenue item
      if (window.formHandler) {
        window.formHandler.addRevenueItem();
        
        // Wait for DOM update
        await this.sleep(100);
        
        // Set values
        const itemIndex = i + 1;
        this.setFieldValue(`revenueName_${itemIndex}`, item.name);
        this.setFieldValue(`revenueValue_${itemIndex}`, item.value);
        
        if (item.growthType) {
          this.setFieldValue(`growthType_${itemIndex}`, item.growthType);
          
          // Set growth rate based on type
          if (item.growthType === 'linear' && item.growthRate) {
            this.setFieldValue(`linearGrowth_${itemIndex}`, item.growthRate);
          } else if (item.growthType === 'annual' && item.growthRate) {
            this.setFieldValue(`annualGrowth_${itemIndex}`, item.growthRate);
          }
        }
      }
    }
    
    this.showProgress(`Applied ${data.revenueItems.value.length} revenue items`);
  }

  async applyCostItems(data) {
    // Apply operating expenses
    if (data.operatingExpenses?.value && Array.isArray(data.operatingExpenses.value)) {
      console.log('💸 Applying operating expenses:', data.operatingExpenses.value);
      
      const container = document.getElementById('operatingExpensesContainer');
      if (container) {
        container.innerHTML = '';
      }
      
      for (let i = 0; i < data.operatingExpenses.value.length; i++) {
        const item = data.operatingExpenses.value[i];
        
        if (window.formHandler) {
          window.formHandler.addOperatingExpense();
          await this.sleep(100);
          
          const itemIndex = i + 1;
          this.setFieldValue(`opExName_${itemIndex}`, item.name);
          this.setFieldValue(`opExValue_${itemIndex}`, item.value);
          
          if (item.growthType) {
            this.setFieldValue(`opExGrowthType_${itemIndex}`, item.growthType);
          }
        }
      }
      
      this.showProgress(`Applied ${data.operatingExpenses.value.length} operating expenses`);
    }
    
    // Apply capital expenses
    if (data.capitalExpenses?.value && Array.isArray(data.capitalExpenses.value)) {
      console.log('🏗️ Applying capital expenses:', data.capitalExpenses.value);
      
      const container = document.getElementById('capitalExpensesContainer');
      if (container) {
        container.innerHTML = '';
      }
      
      for (let i = 0; i < data.capitalExpenses.value.length; i++) {
        const item = data.capitalExpenses.value[i];
        
        if (window.formHandler) {
          window.formHandler.addCapitalExpense();
          await this.sleep(100);
          
          const itemIndex = i + 1;
          this.setFieldValue(`capExName_${itemIndex}`, item.name);
          this.setFieldValue(`capExValue_${itemIndex}`, item.value);
          
          if (item.growthType) {
            this.setFieldValue(`capExGrowthType_${itemIndex}`, item.growthType);
          }
        }
      }
      
      this.showProgress(`Applied ${data.capitalExpenses.value.length} capital expenses`);
    }
  }

  async applyDebtModel(data) {
    if (!data) return;
    
    // Apply debt model parameters
    if (data.loanIssuanceFees?.value) {
      this.setFieldValue('loanIssuanceFees', data.loanIssuanceFees.value, data.loanIssuanceFees);
    }
    if (data.interestRateType?.value) {
      const radioBtn = document.querySelector(`input[name="rateType"][value="${data.interestRateType.value}"]`);
      if (radioBtn) {
        radioBtn.checked = true;
        radioBtn.dispatchEvent(new Event('change', { bubbles: true }));
        this.addConfidenceIndicator(radioBtn, data.interestRateType);
      }
    }
    if (data.interestRate?.value) {
      if (data.interestRateType?.value === 'fixed') {
        this.setFieldValue('fixedRate', data.interestRate.value, data.interestRate);
      }
    }
    if (data.baseRate?.value) {
      this.setFieldValue('baseRate', data.baseRate.value, data.baseRate);
    }
    if (data.creditMargin?.value) {
      this.setFieldValue('creditMargin', data.creditMargin.value, data.creditMargin);
    }
  }

  async applyExitAssumptions(data) {
    if (!data) return;
    
    // Apply exit assumptions
    if (data.disposalCost?.value) {
      this.setFieldValue('disposalCost', data.disposalCost.value, data.disposalCost);
    }
    if (data.terminalCapRate?.value) {
      this.setFieldValue('terminalCapRate', data.terminalCapRate.value, data.terminalCapRate);
    }
  }

  setFieldValue(fieldId, value, confidenceData = null) {
    const element = document.getElementById(fieldId);
    if (!element) {
      console.warn(`Field not found: ${fieldId}`);
      return;
    }
    
    element.value = value;
    element.dispatchEvent(new Event('input', { bubbles: true }));
    element.dispatchEvent(new Event('change', { bubbles: true }));
    
    // Add confidence indicator if available
    if (confidenceData && this.confidenceIndicator) {
      this.addConfidenceIndicator(element, confidenceData);
    }
    
    // Add visual feedback
    element.classList.add('field-updated');
    setTimeout(() => {
      element.classList.remove('field-updated');
    }, 1000);
  }

  addConfidenceIndicator(element, confidenceData) {
    if (!this.confidenceIndicator || !confidenceData) return;
    
    this.confidenceIndicator.addToField(element, confidenceData, {
      showTooltip: true,
      showProgressBar: false,
      position: 'after'
    });
  }

  // Utility methods
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  showLoadingState(show) {
    const autoFillBtn = document.getElementById('autoFillBtn');
    if (autoFillBtn) {
      if (show) {
        autoFillBtn.disabled = true;
        autoFillBtn.innerHTML = `
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="animate-spin">
            <path d="M21 12a9 9 0 11-6.219-8.56"/>
          </svg>
          Analyzing with AI...
        `;
      } else {
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

  showProgress(message) {
    console.log(`📊 ${message}`);
    // You could add a progress indicator here
  }

  showSuccess(message) {
    console.log(`✅ ${message}`);
    this.showNotification(message, 'success');
  }

  showError(message) {
    console.error(`❌ ${message}`);
    this.showNotification(message, 'error');
  }

  showInfo(message) {
    console.log(`ℹ️ ${message}`);
    this.showNotification(message, 'info');
  }

  showNotification(message, type = 'info') {
    // Create a simple notification
    const notification = document.createElement('div');
    notification.className = `extraction-notification ${type}`;
    notification.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      padding: 12px 16px;
      border-radius: 6px;
      color: white;
      font-weight: 500;
      max-width: 300px;
      z-index: 10000;
      box-shadow: 0 4px 12px rgba(0,0,0,0.15);
      ${type === 'success' ? 'background: #22c55e;' : ''}
      ${type === 'error' ? 'background: #ef4444;' : ''}
      ${type === 'info' ? 'background: #3b82f6;' : ''}
    `;
    notification.textContent = message;
    
    document.body.appendChild(notification);
    
    // Auto remove after 5 seconds
    setTimeout(() => {
      if (notification.parentElement) {
        notification.parentElement.removeChild(notification);
      }
    }, 5000);
  }

  showIntegrationStatus() {
    const components = [
      { name: 'AI Extraction Service', instance: this.aiExtractionService },
      { name: 'Data Standardizer', instance: this.dataStandardizer },
      { name: 'Field Mapping Engine', instance: this.fieldMappingEngine },
      { name: 'Confidence Indicator', instance: this.confidenceIndicator },
      { name: 'Review Modal', instance: this.reviewModal },
      { name: 'Extraction History', instance: this.extractionHistory },
      { name: 'High Level Extractor', instance: this.highLevelExtractor },
      { name: 'Deal Assumptions Extractor', instance: this.dealAssumptionsExtractor },
      { name: 'Revenue Items Extractor', instance: this.revenueItemsExtractor },
      { name: 'Cost Items Extractor', instance: this.costItemsExtractor },
      { name: 'Debt Model Extractor', instance: this.debtModelExtractor },
      { name: 'Exit Assumptions Extractor', instance: this.exitAssumptionsExtractor }
    ];
    
    const activeComponents = components.filter(c => c.instance).length;
    const totalComponents = components.length;
    
    console.log(`🎯 AutoFill Integration Status: ${activeComponents}/${totalComponents} components active`);
    components.forEach(c => {
      console.log(`  ${c.instance ? '✅' : '❌'} ${c.name}`);
    });
    
    if (activeComponents === totalComponents) {
      this.showSuccess('🚀 Advanced AI extraction system fully loaded! Upload files and click "Auto Fill with AI" to see the new features.');
    } else {
      this.showInfo(`⚠️ ${activeComponents}/${totalComponents} extraction components loaded. Some advanced features may not be available.`);
    }
  }

  updateHistoryUI(historyData) {
    // Update any history-related UI elements
    console.log('📚 History updated:', historyData.statistics);
  }

  onFilesChanged(files) {
    console.log('📁 Files changed, new count:', files.length);
    this.uploadedFiles = files;
    
    // Enable/disable autofill button based on file availability
    const autoFillBtn = document.getElementById('autoFillBtn');
    if (autoFillBtn) {
      if (files.length > 0) {
        autoFillBtn.disabled = false;
        console.log('✅ AutoFill button enabled - files ready');
      } else {
        autoFillBtn.disabled = true;
        console.log('⚠️ AutoFill button disabled - no files');
      }
    }
  }

  // Public methods for manual control
  async extractSpecificSection(sectionName) {
    if (!this.uploadedFiles.length) {
      this.showError('No files uploaded');
      return null;
    }
    
    const extractor = this[`${sectionName}Extractor`];
    if (!extractor) {
      this.showError(`Extractor not found: ${sectionName}`);
      return null;
    }
    
    // Read file contents first
    const filesWithContent = await this.readFileContents();
    return await extractor.extract(filesWithContent);
  }

  clearAllConfidenceIndicators() {
    if (this.confidenceIndicator) {
      this.confidenceIndicator.clearAllIndicators();
    }
  }

  getExtractionHistory() {
    return this.extractionHistory ? this.extractionHistory.getAllSessions() : [];
  }

  undoLastExtraction() {
    if (this.extractionHistory) {
      const previousSession = this.extractionHistory.undo();
      if (previousSession) {
        this.applyExtractedData(previousSession.extractedData);
        this.showInfo('Reverted to previous extraction');
      } else {
        this.showInfo('No previous extraction to undo');
      }
    }
  }
}

// Export for use in main application
window.AutoFillIntegrator = AutoFillIntegrator;