/* global Office, Excel */

class MAModelingAddin {
  constructor() {
    this.isInitialized = false;
    
    // Initialize widget instances
    this.excelGenerator = null;
    this.formHandler = null;
    this.fileUploader = null;
    this.chatHandler = null;
    this.dataManager = null;
    this.uiController = null;
    
    // AI Extraction widgets
    this.masterDataAnalyzer = null;
    this.highLevelParametersExtractor = null;
    this.dealAssumptionsExtractor = null;
    
    // New AI Extraction System
    this.autoFillIntegrator = null;

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
    
    // Initialize widget instances
    this.initializeWidgets();
    
    // Set up main event listeners
    this.setupMainEventListeners();
    
    // Restore collapsed states
    this.restoreCollapsedStates();
    
    this.isInitialized = true;
    console.log('MAModelingAddin initialized successfully');
    
    // Add-in loaded successfully
    console.log('✅ Add-in loaded successfully! All widgets ready.');
    
    // Test if Office.js is working
    if (typeof Office !== 'undefined' && Office.context) {
      console.log('📊 Excel integration ready! You can use all features.');
    } else {
      console.log('⚠️ Excel integration limited - some features may not work.');
    }
  }

  initializeWidgets() {
    console.log('Initializing widgets...');
    
    // Initialize ExcelGenerator
    if (typeof ExcelGenerator !== 'undefined') {
      this.excelGenerator = new ExcelGenerator();
      window.excelGenerator = this.excelGenerator;
      console.log('✅ ExcelGenerator initialized');
    }
    
    // Initialize FormHandler
    if (typeof FormHandler !== 'undefined') {
      this.formHandler = new FormHandler();
      this.formHandler.initialize();
      window.formHandler = this.formHandler;
      console.log('✅ FormHandler initialized');
    }
    
    // Initialize FileUploader
    if (typeof FileUploader !== 'undefined') {
      this.fileUploader = new FileUploader();
      this.fileUploader.initialize();
      window.fileUploader = this.fileUploader;
      console.log('✅ FileUploader initialized');
    }
    
    // Initialize ChatHandler
    if (typeof ChatHandler !== 'undefined') {
      this.chatHandler = new ChatHandler();
      this.chatHandler.initialize();
      window.chatHandler = this.chatHandler;
      console.log('✅ ChatHandler initialized');
    }
    
    // Initialize DataManager
    if (typeof DataManager !== 'undefined') {
      this.dataManager = new DataManager();
      this.dataManager.initialize();
      window.dataManager = this.dataManager;
      console.log('✅ DataManager initialized');
    }
    
    // Initialize UIController
    if (typeof UIController !== 'undefined') {
      this.uiController = new UIController();
      this.uiController.initialize();
      window.uiController = this.uiController;
      console.log('✅ UIController initialized');
    }
    
    // Initialize AI Extraction widgets (legacy support)
    if (typeof MasterDataAnalyzer !== 'undefined') {
      this.masterDataAnalyzer = new MasterDataAnalyzer();
      this.masterDataAnalyzer.initialize();
      window.masterDataAnalyzer = this.masterDataAnalyzer;
      console.log('✅ MasterDataAnalyzer initialized');
    }
    
    // Initialize New AI Extraction System
    if (typeof AutoFillIntegrator !== 'undefined') {
      this.autoFillIntegrator = new AutoFillIntegrator();
      // Initialize async
      this.autoFillIntegrator.initialize().then(() => {
        console.log('✅ AutoFillIntegrator async initialization completed');
      }).catch(error => {
        console.error('❌ AutoFillIntegrator initialization failed:', error);
      });
      window.autoFillIntegrator = this.autoFillIntegrator;
      console.log('✅ AutoFillIntegrator initialization started');
    }
    
    // Auto-load saved data
    if (this.dataManager) {
      this.dataManager.autoLoadSavedData();
    }
  }

  setupMainEventListeners() {
    console.log('Setting up main event listeners...');
    
    // Generate Model button
    const generateModelBtn = document.getElementById('generateModelBtn');
    if (generateModelBtn) {
      generateModelBtn.addEventListener('click', () => this.generateModel());
      console.log('Generate model button listener added');
    }
    
    // Validate Model button (if exists)
    const validateModelBtn = document.getElementById('validateModelBtn');
    if (validateModelBtn) {
      validateModelBtn.addEventListener('click', () => this.validateModel());
      console.log('Validate model button listener added');
    }
    
    // Setup collapsible sections
    this.setupCollapsibleSections();
  }

  setupCollapsibleSections() {
    console.log('Setting up collapsible sections...');
    
    // Find all collapsible sections
    const collapsibleSections = document.querySelectorAll('.collapsible-section');
    console.log(`Found ${collapsibleSections.length} collapsible sections`);
    
    collapsibleSections.forEach((section, index) => {
      const header = section.querySelector('h3');
      console.log(`Section ${index + 1}: ID=${section.id}, Header found=${!!header}`);
      
      if (header) {
        // Remove any existing event listeners by cloning the element
        const newHeader = header.cloneNode(true);
        header.parentNode.replaceChild(newHeader, header);
        
        // Add our event listener
        newHeader.addEventListener('click', (e) => {
          console.log(`🎯 Header clicked for section: ${section.id}`);
          e.preventDefault();
          e.stopPropagation();
          this.toggleSection(section);
        });
        
        // Add hover effect
        newHeader.style.cursor = 'pointer';
        newHeader.style.userSelect = 'none';
        console.log(`✅ Collapsible header configured for: ${section.id}`);
      } else {
        console.warn(`❌ No header found for section: ${section.id}`);
      }
    });
    
    console.log(`✅ ${collapsibleSections.length} collapsible sections configured`);
    
    // Test function to manually toggle first section
    window.testToggle = () => {
      const firstSection = document.querySelector('.collapsible-section');
      if (firstSection) {
        console.log('🧪 Manual test toggle triggered');
        this.toggleSection(firstSection);
      }
    };
    
    console.log('🧪 Test function available: window.testToggle()');
  }

  toggleSection(section) {
    console.log(`🔄 Toggling section: ${section.id}`);
    
    const isCollapsed = section.classList.contains('collapsed');
    console.log(`📋 Current state - isCollapsed: ${isCollapsed}`);
    console.log(`📋 Current classes: ${section.className}`);
    
    if (isCollapsed) {
      // Show section
      section.classList.remove('collapsed');
      console.log(`✅ Showed section: ${section.id}`);
      console.log(`📋 New classes after show: ${section.className}`);
    } else {
      // Hide section
      section.classList.add('collapsed');
      console.log(`❌ Hidden section: ${section.id}`);
      console.log(`📋 New classes after hide: ${section.className}`);
    }
    
    // Force a reflow to ensure CSS changes take effect
    section.offsetHeight;
    
    // Store the state in localStorage for persistence
    const sectionId = section.id;
    if (sectionId) {
      const newState = !isCollapsed;
      localStorage.setItem(`section-${sectionId}-collapsed`, newState);
      console.log(`💾 Stored state for ${sectionId}: collapsed=${newState}`);
    }
  }

  // Restore collapsed states from localStorage
  restoreCollapsedStates() {
    const collapsibleSections = document.querySelectorAll('.collapsible-section');
    
    collapsibleSections.forEach(section => {
      const sectionId = section.id;
      if (sectionId) {
        const isCollapsed = localStorage.getItem(`section-${sectionId}-collapsed`) === 'true';
        if (isCollapsed) {
          section.classList.add('collapsed');
        }
      }
    });
    
    console.log('✅ Restored section collapsed states from localStorage');
  }

  async generateModel() {
    console.log('Starting model generation...');
    
    try {
      // Validate form data
      if (this.formHandler) {
        const validation = this.formHandler.validateAllFields();
        if (!validation.isValid) {
          const errorMessage = 'Please complete the following required fields:\n' + validation.errors.join('\n');
          alert(errorMessage);
          console.log('Validation failed:', validation.errors);
          return;
        }
      }
      
      // Collect model data
      let modelData = {};
      if (this.formHandler) {
        modelData = this.formHandler.collectAllModelData();
        console.log('Model data collected:', modelData);
      }
      
      // Generate Excel model
      if (this.excelGenerator) {
        const result = await this.excelGenerator.generateModel(modelData);
        
        if (result.success) {
          console.log('Model generated successfully');
          if (this.uiController) {
            this.uiController.showMessage('Excel model generated successfully!', 'success');
          } else {
            alert('Excel model generated successfully!');
          }
        } else {
          console.error('Model generation failed:', result.error);
          if (this.uiController) {
            this.uiController.showMessage('Error generating model: ' + result.error, 'error');
          } else {
            alert('Error generating model: ' + result.error);
          }
        }
      } else {
        console.error('ExcelGenerator not available');
        alert('Excel generator not available. Please refresh the page.');
      }
      
    } catch (error) {
      console.error('Error in generateModel:', error);
      if (this.uiController) {
        this.uiController.showMessage('Unexpected error: ' + error.message, 'error');
      } else {
        alert('Unexpected error: ' + error.message);
      }
    }
  }

  async validateModel() {
    console.log('Validating model...');
    
    try {
      if (this.formHandler) {
        const validation = this.formHandler.validateAllFields();
        
        if (validation.isValid) {
          if (this.uiController) {
            this.uiController.showMessage('All required fields are completed! Ready to generate model.', 'success');
          } else {
            alert('All required fields are completed! Ready to generate model.');
          }
        } else {
          const errorMessage = 'Missing required fields:\n' + validation.errors.join('\n');
          if (this.uiController) {
            this.uiController.showMessage('Validation failed. Check console for details.', 'warning');
          }
          console.log('Validation errors:', validation.errors);
          alert(errorMessage);
        }
      } else {
        console.error('FormHandler not available');
        alert('Form validation not available. Please refresh the page.');
      }
      
    } catch (error) {
      console.error('Error in validateModel:', error);
      if (this.uiController) {
        this.uiController.showMessage('Error during validation: ' + error.message, 'error');
      } else {
        alert('Error during validation: ' + error.message);
      }
    }
  }

  // Utility methods for backward compatibility
  async getExcelContext() {
    try {
      if (typeof Excel === 'undefined') {
        return JSON.stringify({ error: 'Excel not available' });
      }

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
    } catch (error) {
      console.error('Error getting Excel context:', error);
      return JSON.stringify({ error: 'Could not read Excel context' });
    }
  }

  // Widget access methods
  getExcelGenerator() {
    return this.excelGenerator;
  }

  getFormHandler() {
    return this.formHandler;
  }

  getFileUploader() {
    return this.fileUploader;
  }

  getChatHandler() {
    return this.chatHandler;
  }

  getDataManager() {
    return this.dataManager;
  }

  getUIController() {
    return this.uiController;
  }

  // Status methods
  isReady() {
    return this.isInitialized && 
           this.excelGenerator && 
           this.formHandler && 
           this.fileUploader && 
           this.dataManager && 
           this.uiController;
  }

  getStatus() {
    return {
      initialized: this.isInitialized,
      widgets: {
        excelGenerator: !!this.excelGenerator,
        formHandler: !!this.formHandler,
        fileUploader: !!this.fileUploader,
        chatHandler: !!this.chatHandler,
        dataManager: !!this.dataManager,
        uiController: !!this.uiController
      },
      officeAvailable: typeof Office !== 'undefined',
      excelAvailable: typeof Excel !== 'undefined'
    };
  }
}

// Global error handlers
window.addEventListener('error', (e) => {
  console.error('Global error caught:', e.error, e.filename, e.lineno);
});

window.addEventListener('unhandledrejection', (e) => {
  console.error('Unhandled promise rejection:', e.reason);
});

// Initialize the add-in only once
if (!window.maModelingAddin) {
  console.log('Initializing MAModelingAddin...');
  console.log('Office availability:', typeof Office !== 'undefined');
  console.log('Excel availability:', typeof Excel !== 'undefined');

  // Wait for everything to load properly
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      if (!window.maModelingAddin) {
        window.maModelingAddin = new MAModelingAddin();
      }
    });
  } else {
    window.maModelingAddin = new MAModelingAddin();
  }
}

// Export for debugging
window.MAModelingAddin = MAModelingAddin;