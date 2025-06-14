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
    
    // Hook into the test autofill button
    const testAutoFillBtn = document.getElementById('testAutoFillBtn');
    if (testAutoFillBtn) {
      testAutoFillBtn.addEventListener('click', () => this.handleTestAutoFill());
      console.log('✅ Test AutoFill button handler attached');
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
    console.log('🚀 AutoFill button clicked!');
    console.log('🔍 Current state check:');
    console.log('  - isProcessing:', this.isProcessing);
    console.log('  - uploadedFiles:', this.uploadedFiles);
    console.log('  - uploadedFiles length:', this.uploadedFiles ? this.uploadedFiles.length : 'null');
    
    if (this.isProcessing) {
      this.showError('Extraction already in progress. Please wait...');
      return;
    }
    
    if (!this.uploadedFiles || this.uploadedFiles.length === 0) {
      console.error('❌ No files uploaded! Checking file uploader...');
      
      // Try to get files from file uploader directly
      if (window.fileUploader) {
        const files = window.fileUploader.getUploadedFiles();
        console.log('📁 Files from fileUploader.getUploadedFiles():', files);
        if (files && files.length > 0) {
          this.uploadedFiles = files;
          console.log('✅ Found files in fileUploader, continuing...');
        } else {
          this.showError('Please upload files first before using AI autofill.');
          return;
        }
      } else {
        this.showError('File uploader not available. Please refresh the page.');
        return;
      }
    }
    
    console.log('🤖 Starting comprehensive AI extraction...');
    console.log('📄 Files to process:', this.uploadedFiles.map(f => f.name || 'unnamed'));
    this.isProcessing = true;
    
    try {
      // Show loading state
      this.showLoadingState(true);
      
      // Step 1: Read file contents
      console.log('📖 Step 1: Reading file contents...');
      const filesWithContent = await this.readFileContents();
      console.log('📖 Files with content:', filesWithContent.map(f => ({
        name: f.name,
        contentLength: f.content ? f.content.length : 0,
        hasContent: !!f.content
      })));
      
      if (!filesWithContent || filesWithContent.length === 0) {
        throw new Error('No file contents could be read');
      }
      
      // Step 2: Extract data from all widgets
      console.log('📊 Step 2: Extracting data from all sections...');
      const extractionResults = await this.extractAllData(filesWithContent);
      console.log('📊 Extraction results summary:', Object.keys(extractionResults));
      
      // Step 3: Apply extracted data directly (skip modal for now)
      console.log('📊 Step 3: Auto-applying extracted data...');
      
      // Log all extracted data to see what AI returned
      console.log('🔍 EXTRACTED DATA CHECK - What AI returned:');
      console.log('📊 High Level Parameters:', JSON.stringify(extractionResults.highLevelParameters, null, 2));
      console.log('💰 Deal Assumptions:', JSON.stringify(extractionResults.dealAssumptions, null, 2));
      console.log('📈 Revenue Items:', JSON.stringify(extractionResults.revenueItems, null, 2));
      console.log('💸 Cost Items:', JSON.stringify(extractionResults.costItems, null, 2));
      console.log('🏦 Debt Model:', JSON.stringify(extractionResults.debtModel, null, 2));
      console.log('🚪 Exit Assumptions:', JSON.stringify(extractionResults.exitAssumptions, null, 2));
      
      // Apply all extracted data to form
      if (extractionResults.highLevelParameters) {
        console.log('✏️ Applying high level parameters...');
        await this.applyExtractedData('highLevelParameters', extractionResults.highLevelParameters);
      }
      if (extractionResults.dealAssumptions) {
        console.log('✏️ Applying deal assumptions...');
        await this.applyExtractedData('dealAssumptions', extractionResults.dealAssumptions);
      }
      if (extractionResults.revenueItems) {
        console.log('✏️ Applying revenue items...');
        await this.applyExtractedData('revenueItems', extractionResults.revenueItems);
      }
      if (extractionResults.costItems) {
        console.log('✏️ Applying cost items...');
        await this.applyExtractedData('costItems', extractionResults.costItems);
      }
      if (extractionResults.debtModel) {
        console.log('✏️ Applying debt model...');
        await this.applyExtractedData('debtModel', extractionResults.debtModel);
      }
      if (extractionResults.exitAssumptions) {
        console.log('✏️ Applying exit assumptions...');
        await this.applyExtractedData('exitAssumptions', extractionResults.exitAssumptions);
      }
      
      this.showSuccess('✅ AI autofill completed! All sections have been filled with extracted data.');
      
    } catch (error) {
      console.error('❌ Extraction failed:', error);
      console.error('❌ Full error details:', {
        message: error.message,
        stack: error.stack,
        name: error.name
      });
      
      // Check if it's an API service issue
      if (error.message.includes('AI service is currently unavailable') || 
          error.message.includes('500') || 
          error.message.includes('API error')) {
        this.showError('🚨 AI service is currently down. The extraction system requires AI to function. Please contact support or try again later.');
      } else if (error.message.includes('No file contents could be read')) {
        this.showError('🚨 Could not read file contents. Please check your file format and try again.');
      } else if (error.message.includes('FileUploader not available')) {
        this.showError('🚨 File upload system not ready. Please refresh the page and try again.');
      } else {
        this.showError(`🚨 Extraction failed: ${error.message}. Check console for details.`);
      }
    } finally {
      this.isProcessing = false;
      this.showLoadingState(false);
    }
  }

  async handleTestAutoFill() {
    console.log('🧪 Test AutoFill button clicked!');
    
    if (this.isProcessing) {
      this.showError('Extraction already in progress. Please wait...');
      return;
    }
    
    console.log('🎯 Creating sample data for testing...');
    this.isProcessing = true;
    
    try {
      this.showLoadingState(true);
      
      // Create comprehensive sample data
      const sampleData = {
        extractedData: {
          // High Level Parameters
          currency: { value: "USD", confidence: 0.9, source: "test_data" },
          projectStartDate: { value: "2025-01-01", confidence: 0.9, source: "test_data" },
          projectEndDate: { value: "2027-12-31", confidence: 0.9, source: "test_data" },
          modelPeriods: { value: "monthly", confidence: 0.9, source: "test_data" },
          
          // Deal Assumptions
          dealName: { value: "TechCorp Acquisition", confidence: 0.9, source: "test_data" },
          dealValue: { value: 50000000, confidence: 0.9, source: "test_data" },
          transactionFee: { value: 2.5, confidence: 0.8, source: "test_data" },
          dealLTV: { value: 75, confidence: 0.8, source: "test_data" },
          
          // Revenue Items
          revenueItems: {
            value: [
              {
                name: "Software Licensing",
                value: 15000000,
                initialValue: 15000000,
                growthType: "linear",
                growthRate: 3
              },
              {
                name: "Support Services", 
                value: 8000000,
                initialValue: 8000000,
                growthType: "linear",
                growthRate: 2
              },
              {
                name: "Professional Services",
                value: 5000000,
                initialValue: 5000000,
                growthType: "linear",
                growthRate: 5
              }
            ],
            confidence: 0.8,
            source: "test_data"
          },
          
          // Operating Expenses
          operatingExpenses: {
            value: [
              {
                name: "Staff Costs",
                value: 12000000,
                initialValue: 12000000,
                growthType: "linear",
                growthRate: 4
              },
              {
                name: "Marketing",
                value: 3000000,
                initialValue: 3000000,
                growthType: "linear",
                growthRate: 2
              },
              {
                name: "Office Rent",
                value: 1200000,
                initialValue: 1200000,
                growthType: "linear",
                growthRate: 1
              }
            ],
            confidence: 0.8,
            source: "test_data"
          },
          
          // Capital Expenses
          capitalExpenses: {
            value: [
              {
                name: "IT Equipment",
                value: 2000000,
                initialValue: 2000000,
                growthType: "linear",
                growthRate: 0
              },
              {
                name: "Office Furniture",
                value: 500000,
                initialValue: 500000,
                growthType: "linear",
                growthRate: 0
              }
            ],
            confidence: 0.7,
            source: "test_data"
          },
          
          // Exit Assumptions
          disposalCost: { value: 2.0, confidence: 0.8, source: "test_data" },
          terminalCapRate: { value: 8.5, confidence: 0.8, source: "test_data" },
          
          // Debt Model
          loanIssuanceFees: { value: 1.5, confidence: 0.7, source: "test_data" },
          interestRateType: { value: "fixed", confidence: 0.8, source: "test_data" },
          interestRate: { value: 5.5, confidence: 0.8, source: "test_data" }
        }
      };
      
      console.log('📊 Sample data created:', sampleData);
      
      // Apply the sample data using FieldMappingEngine
      if (this.fieldMappingEngine) {
        console.log('🗺️ Applying sample data via FieldMappingEngine...');
        const result = await this.fieldMappingEngine.applyDataToForm(sampleData.extractedData);
        console.log('✅ Sample data applied successfully:', result);
        this.showSuccess('✅ Test autofill completed! Sample data has been applied to all sections.');
      } else {
        console.error('❌ FieldMappingEngine not available');
        this.showError('🚨 FieldMappingEngine not available. Please refresh and try again.');
      }
      
    } catch (error) {
      console.error('❌ Test AutoFill failed:', error);
      this.showError(`🚨 Test failed: ${error.message}`);
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
    
    // Extract from each section SEQUENTIALLY to avoid API overload
    console.log('📊 Starting sequential extraction to avoid API overload...');
    
    if (this.highLevelExtractor) {
      console.log('🎯 Starting high-level parameters extraction...');
      try {
        const data = await this.highLevelExtractor.extract(filesWithContent);
        console.log('🎯 High-level parameters extraction completed:', data);
        extractionResults.highLevelParameters = data;
        Object.assign(allExtractedData, data);
        this.showProgress('High-level parameters extracted');
      } catch (error) {
        console.error('❌ High-level parameters extraction failed:', error);
      }
    } else {
      console.warn('🎯 High-level parameters extractor not available');
    }
    
    if (this.dealAssumptionsExtractor) {
      console.log('💼 Starting deal assumptions extraction...');
      try {
        const data = await this.dealAssumptionsExtractor.extract(filesWithContent);
        extractionResults.dealAssumptions = data;
        Object.assign(allExtractedData, data);
        this.showProgress('Deal assumptions extracted');
      } catch (error) {
        console.error('❌ Deal assumptions extraction failed:', error);
      }
    }
    
    if (this.revenueItemsExtractor) {
      console.log('💰 Starting revenue items extraction...');
      try {
        const data = await this.revenueItemsExtractor.extract(filesWithContent);
        extractionResults.revenueItems = data;
        Object.assign(allExtractedData, data);
        this.showProgress('Revenue items extracted');
      } catch (error) {
        console.error('❌ Revenue items extraction failed:', error);
      }
    }
    
    if (this.costItemsExtractor) {
      console.log('💸 Starting cost items extraction...');
      try {
        const data = await this.costItemsExtractor.extract(filesWithContent);
        extractionResults.costItems = data;
        Object.assign(allExtractedData, data);
        this.showProgress('Cost items extracted');
      } catch (error) {
        console.error('❌ Cost items extraction failed:', error);
      }
    }
    
    if (this.debtModelExtractor) {
      console.log('🏦 Starting debt model extraction...');
      try {
        const data = await this.debtModelExtractor.extract(filesWithContent);
        extractionResults.debtModel = data;
        Object.assign(allExtractedData, data);
        this.showProgress('Debt model extracted');
      } catch (error) {
        console.error('❌ Debt model extraction failed:', error);
      }
    }
    
    if (this.exitAssumptionsExtractor) {
      console.log('🚪 Starting exit assumptions extraction...');
      try {
        const data = await this.exitAssumptionsExtractor.extract(filesWithContent);
        extractionResults.exitAssumptions = data;
        Object.assign(allExtractedData, data);
        this.showProgress('Exit assumptions extracted');
      } catch (error) {
        console.error('❌ Exit assumptions extraction failed:', error);
      }
    }
    
    // All extractions now complete (sequential)
    
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
    if (!data.revenueItems?.value || !Array.isArray(data.revenueItems.value)) {
      console.log('💰 No revenue items to apply');
      return;
    }
    
    console.log('💰 Applying revenue items:', data.revenueItems.value);
    
    try {
      // Clear existing revenue items
      const container = document.getElementById('revenueItemsContainer');
      if (container) {
        container.innerHTML = '';
      } else {
        console.warn('💰 Revenue items container not found');
        return;
      }
      
      // Check if formHandler is available
      if (!window.formHandler || typeof window.formHandler.addRevenueItem !== 'function') {
        console.warn('💰 FormHandler not available, cannot add revenue items');
        return;
      }
      
      // Add each revenue item
      for (let i = 0; i < data.revenueItems.value.length; i++) {
        const item = data.revenueItems.value[i];
        
        try {
          // Add new revenue item
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
          
          console.log(`💰 Applied revenue item ${i + 1}: ${item.name}`);
        } catch (itemError) {
          console.error(`💰 Error applying revenue item ${i + 1}:`, itemError);
        }
      }
      
      this.showProgress(`Applied ${data.revenueItems.value.length} revenue items`);
    } catch (error) {
      console.error('💰 Error applying revenue items:', error);
      this.showProgress('Revenue items application failed');
    }
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

  /**
   * Apply extracted data using the FieldMappingEngine
   * This is the new method that should be used instead of individual apply methods
   */
  async applyExtractedData(sectionType, extractedData) {
    if (!this.fieldMappingEngine || !extractedData) {
      console.warn(`🗺️ Cannot apply ${sectionType}: FieldMappingEngine or data missing`);
      return;
    }

    try {
      // For array-based sections (revenue, costs), we need to transform the data structure
      if (sectionType === 'revenueItems' && extractedData.revenueItems?.value) {
        console.log('🗺️ Applying revenue items via FieldMappingEngine...');
        const standardizedData = {
          revenueItems: extractedData.revenueItems
        };
        await this.fieldMappingEngine.applyDataToForm(standardizedData);
        
      } else if (sectionType === 'costItems') {
        console.log('🗺️ Applying cost items via FieldMappingEngine...');
        const standardizedData = {};
        
        // Add operating expenses if they exist
        if (extractedData.operatingExpenses?.value) {
          standardizedData.operatingExpenses = extractedData.operatingExpenses;
        }
        
        // Add capital expenses if they exist  
        if (extractedData.capitalExpenses?.value) {
          standardizedData.capitalExpenses = extractedData.capitalExpenses;
        }
        
        await this.fieldMappingEngine.applyDataToForm(standardizedData);
        
      } else {
        // For simple field-based sections, use the FieldMappingEngine directly
        console.log(`🗺️ Applying ${sectionType} via FieldMappingEngine...`);
        await this.fieldMappingEngine.applyDataToForm(extractedData);
      }
      
    } catch (error) {
      console.error(`🗺️ Error applying ${sectionType} with FieldMappingEngine:`, error);
      // Fallback to legacy methods
      console.log(`🗺️ Falling back to legacy apply method for ${sectionType}`);
      await this.applyExtractedDataLegacy(sectionType, extractedData);
    }
  }

  /**
   * Legacy apply method - kept for fallback
   */
  async applyExtractedDataLegacy(sectionType, extractedData) {
    switch (sectionType) {
      case 'highLevelParameters':
        await this.applyHighLevelParameters(extractedData);
        break;
      case 'dealAssumptions':
        await this.applyDealAssumptions(extractedData);
        break;
      case 'revenueItems':
        await this.applyRevenueItems(extractedData);
        break;
      case 'costItems':
        await this.applyCostItems(extractedData);
        break;
      case 'debtModel':
        await this.applyDebtModel(extractedData);
        break;
      case 'exitAssumptions':
        await this.applyExitAssumptions(extractedData);
        break;
      default:
        console.warn(`Unknown section type: ${sectionType}`);
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