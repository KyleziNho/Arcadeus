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
    
    // Check authentication first
    this.checkAuthAndInitialize();
  }

  checkAuthAndInitialize() {
    // Wait for auth.js to be loaded
    if (!window.arcadeusAuth) {
      setTimeout(() => this.checkAuthAndInitialize(), 100);
      return;
    }
    
    console.log('Checking authentication...');
    
    
    // Check if user is authenticated
    const user = window.arcadeusAuth.checkAuth();
    
    if (!user) {
      console.log('User not authenticated, redirecting to login');
      window.location.href = 'login.html';
      return;
    }
    
    console.log('User authenticated:', user.email);
    
    // Check if user has completed onboarding
    const hasOnboarded = localStorage.getItem('arcadeusOnboarding') === 'completed';
    
    if (!hasOnboarded) {
      console.log('User has not completed onboarding, redirecting...');
      window.location.href = 'onboarding.html';
      return;
    }
    
    console.log('User has completed onboarding, showing app...');
    
    // User is authenticated and onboarded, show the app
    this.hideAuthLoadingScreen();
    
    // Continue with Office initialization
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

  hideAuthLoadingScreen() {
    const authLoading = document.getElementById('authLoading');
    const mainApp = document.getElementById('mainApp');
    
    if (authLoading && mainApp) {
      // Fade out loading screen
      authLoading.style.opacity = '0';
      authLoading.style.transition = 'opacity 0.3s ease-out';
      
      setTimeout(() => {
        authLoading.style.display = 'none';
        mainApp.style.display = 'block';
        mainApp.style.opacity = '0';
        
        // Fade in main app
        setTimeout(() => {
          mainApp.style.transition = 'opacity 0.3s ease-in';
          mainApp.style.opacity = '1';
        }, 50);
      }, 300);
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
    
    // Initialize assumption tabs explicitly
    this.initializeAssumptionTabs();
    
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
    console.log('Window object keys related to Excel:', Object.keys(window).filter(key => key.includes('Excel')));
    console.log('Available classes:', {
      ExcelGenerator: typeof ExcelGenerator,
      FormHandler: typeof FormHandler,
      UIController: typeof UIController
    });
    
    // Initialize ExcelGenerator
    if (typeof window.ExcelGenerator !== 'undefined') {
      try {
        this.excelGenerator = new window.ExcelGenerator();
        window.excelGenerator = this.excelGenerator;
        console.log('✅ ExcelGenerator initialized successfully');
      } catch (error) {
        console.error('❌ Error creating ExcelGenerator:', error);
        this.excelGenerator = null;
      }
    } else {
      console.error('❌ ExcelGenerator class not found. Check if ExcelGenerator.js is loaded.');
      console.log('Available classes:', {
        ExcelGenerator: typeof ExcelGenerator,
        windowExcelGenerator: typeof window.ExcelGenerator
      });
      this.excelGenerator = null;
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
    
    // Initialize New AI Extraction System - DISABLED
    // if (typeof AutoFillIntegrator !== 'undefined') {
    //   this.autoFillIntegrator = new AutoFillIntegrator();
    //   // Initialize async
    //   this.autoFillIntegrator.initialize().then(() => {
    //     console.log('✅ AutoFillIntegrator async initialization completed');
    //   }).catch(error => {
    //     console.error('❌ AutoFillIntegrator initialization failed:', error);
    //   });
    //   window.autoFillIntegrator = this.autoFillIntegrator;
    //   console.log('✅ AutoFillIntegrator initialization started');
    // }
    
    // Auto-load saved data
    if (this.dataManager) {
      this.dataManager.autoLoadSavedData();
    }
  }

  initializeAssumptionTabs() {
    console.log('🔧 Initializing assumption tabs from main class...');
    
    // Call the global function if it exists
    if (typeof window.initializeAssumptionTabs === 'function') {
      window.initializeAssumptionTabs();
      console.log('✅ Assumption tabs initialized via global function');
    } else {
      console.warn('⚠️ Global initializeAssumptionTabs function not found');
    }
  }

  setupMainEventListeners() {
    console.log('Setting up main event listeners...');
    
    // Generate Assumptions button
    const generateAssumptionsBtn = document.getElementById('generateAssumptionsBtn');
    if (generateAssumptionsBtn) {
      generateAssumptionsBtn.addEventListener('click', () => this.generateAssumptions());
      console.log('Generate assumptions button listener added');
    }
    
    // Generate P&L button
    const generatePLBtn = document.getElementById('generatePLBtn');
    if (generatePLBtn) {
      generatePLBtn.addEventListener('click', () => this.generatePLWithAI());
      console.log('Generate P&L button listener added');
    }
    
    // Generate CapEx button
    const generateCapExBtn = document.getElementById('generateCapExBtn');
    if (generateCapExBtn) {
      generateCapExBtn.addEventListener('click', () => this.generateCapExSheet());
      console.log('Generate CapEx button listener added');
    }
    
    // Generate Debt Model button
    const generateDebtModelBtn = document.getElementById('generateDebtModelBtn');
    if (generateDebtModelBtn) {
      generateDebtModelBtn.addEventListener('click', () => this.generateDebtModelSheet());
      console.log('Generate Debt Model button listener added');
    }
    
    // Generate FCF button
    const generateFCFBtn = document.getElementById('generateFCFBtn');
    if (generateFCFBtn) {
      generateFCFBtn.addEventListener('click', () => this.generateFCFWithAI());
      console.log('Generate FCF button listener added');
    }
    
    // Generate Full Model button (unified generation)
    const generateFullModelBtn = document.getElementById('generateFullModelBtn');
    if (generateFullModelBtn) {
      generateFullModelBtn.addEventListener('click', () => this.generateFullModel());
      console.log('Generate Full Model button listener added');
    }
    
    // IRR & MOIC are now calculated automatically in FCF sheet
    // No separate button needed
    
    // Auto-fill Test Data button - DISABLED to prevent conflict with AI autofill
    const autoFillTestDataBtn = document.getElementById('autoFillTestDataBtn');
    if (autoFillTestDataBtn) {
      console.log('⚠️ AutoFillTestDataBtn found but disabled - use AI AutoFill instead');
      // Remove the button to prevent hardcoded values from being used
      autoFillTestDataBtn.style.display = 'none';
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
      const header = section.querySelector('.section-header');
      console.log(`Section ${index + 1}: ID=${section.id}, Header found=${!!header}`);
      
      if (header) {
        // Remove any existing event listeners
        const existingOnClick = header.onclick;
        header.onclick = null;
        
        // Add our event listener directly
        header.addEventListener('click', (e) => {
          console.log(`🎯 Header clicked for section: ${section.id}`);
          e.preventDefault();
          e.stopPropagation();
          this.toggleSection(section);
        });
        
        // Make sure cursor is pointer
        header.style.cursor = 'pointer';
        header.style.userSelect = 'none';
        
        // Add title attribute for user feedback
        header.setAttribute('title', 'Click to expand/collapse');
        
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
    
    // Add global toggle all function
    window.toggleAll = () => {
      const sections = document.querySelectorAll('.collapsible-section');
      sections.forEach(section => this.toggleSection(section));
      console.log('🧪 Toggled all sections');
    };
    
    console.log('🧪 Test functions available: window.testToggle(), window.toggleAll()');
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

  async generateAssumptions() {
    console.log('Starting assumptions generation...');
    
    try {
      // Validate form data
      if (this.formHandler) {
        const validation = this.formHandler.validateAllFields();
        if (!validation.isValid) {
          const errorMessage = 'Please complete the following required fields: ' + validation.errors.join(', ');
          console.log('Validation failed:', validation.errors);
          if (this.uiController) {
            this.uiController.showMessage(errorMessage, 'error');
          }
          return;
        }
      }
      
      // Collect model data
      let modelData = {};
      if (this.formHandler) {
        modelData = this.formHandler.collectAllModelData();
        console.log('Model data collected:', modelData);
      }
      
      // Generate Excel assumptions sheet
      if (this.excelGenerator) {
        const result = await this.excelGenerator.generateModel(modelData);
        
        if (result.success) {
          console.log('Assumptions generated successfully');
          
          // Show the P&L generation button
          const generatePLBtn = document.getElementById('generatePLBtn');
          if (generatePLBtn) {
            generatePLBtn.style.display = 'inline-flex';
          }
          
          if (this.uiController) {
            this.uiController.showMessage('Assumptions sheet created! You can now generate the P&L.', 'success');
          } else {
            console.log('Assumptions sheet created successfully!');
          }
        } else {
          console.error('Assumptions generation failed:', result.error);
          if (this.uiController) {
            this.uiController.showMessage('Error generating assumptions: ' + result.error, 'error');
          } else {
            console.error('Error generating assumptions:', result.error);
            if (this.uiController) {
              this.uiController.showMessage('Error generating assumptions: ' + result.error, 'error');
            }
          }
        }
      } else {
        console.error('Debug info:', {
          excelGenerator: this.excelGenerator,
          ExcelGeneratorClass: typeof ExcelGenerator,
          windowExcelGenerator: typeof window.ExcelGenerator
        });
        console.error('Excel generator not available. Please refresh the page.');
        if (this.uiController) {
          this.uiController.showMessage('Excel generator not available. Please refresh the page.', 'error');
        }
      }
      
    } catch (error) {
      console.error('Error in generateAssumptions:', error);
      if (this.uiController) {
        this.uiController.showMessage('Unexpected error: ' + error.message, 'error');
      } else {
        console.error('Unexpected error:', error.message);
        if (this.uiController) {
          this.uiController.showMessage('Unexpected error: ' + error.message, 'error');
        }
      }
    }
  }
  
  async generatePLWithAI() {
    console.log('Starting AI P&L generation...');
    
    try {
      // Collect model data
      let modelData = {};
      if (this.formHandler) {
        modelData = this.formHandler.collectAllModelData();
        console.log('Model data for AI P&L:', modelData);
      }
      
      // Generate P&L using AI
      if (this.excelGenerator) {
        const result = await this.excelGenerator.generatePLWithAI(modelData);
        
        if (result.success) {
          console.log('P&L generated successfully');
          
          // Show the CapEx generation button
          const generateCapExBtn = document.getElementById('generateCapExBtn');
          if (generateCapExBtn) {
            generateCapExBtn.style.display = 'inline-flex';
          }
          
          if (this.uiController) {
            this.uiController.showMessage('P&L Statement created! You can now generate the CapEx Summary.', 'success');
          } else {
            console.log('P&L Statement created successfully!');
            if (this.uiController) {
              this.uiController.showMessage('P&L Statement created successfully! You can now generate the Free Cash Flow.', 'success');
            }
          }
        } else {
          console.error('AI P&L generation failed:', result.error);
          if (this.uiController) {
            this.uiController.showMessage('Error generating AI P&L: ' + result.error, 'error');
          } else {
            console.error('Error generating AI P&L:', result.error);
            if (this.uiController) {
              this.uiController.showMessage('Error generating AI P&L: ' + result.error, 'error');
            }
          }
        }
      } else {
        console.error('Debug info:', {
          excelGenerator: this.excelGenerator,
          ExcelGeneratorClass: typeof ExcelGenerator,
          windowExcelGenerator: typeof window.ExcelGenerator
        });
        console.error('Excel generator not available. Please refresh the page.');
        if (this.uiController) {
          this.uiController.showMessage('Excel generator not available. Please refresh the page.', 'error');
        }
      }
      
    } catch (error) {
      console.error('Error in generatePLWithAI:', error);
      if (this.uiController) {
        this.uiController.showMessage('Unexpected error: ' + error.message, 'error');
      } else {
        console.error('Unexpected error:', error.message);
        if (this.uiController) {
          this.uiController.showMessage('Unexpected error: ' + error.message, 'error');
        }
      }
    }
  }
  
  async generateCapExSheet() {
    console.log('Starting CapEx sheet generation...');
    
    try {
      // Collect model data
      let modelData = {};
      if (this.formHandler) {
        modelData = this.formHandler.collectAllModelData();
        console.log('Model data for CapEx:', modelData);
      }
      
      // Generate CapEx sheet
      if (this.excelGenerator) {
        const result = await this.excelGenerator.generateCapExSheet(modelData);
        
        if (result && result.success !== false) {
          console.log('CapEx sheet generated successfully');
          
          // Show the Debt Model generation button
          const generateDebtModelBtn = document.getElementById('generateDebtModelBtn');
          if (generateDebtModelBtn) {
            generateDebtModelBtn.style.display = 'inline-flex';
          }
          
          if (this.uiController) {
            this.uiController.showMessage('CapEx Summary created! You can now generate the Debt Model.', 'success');
          }
        } else {
          console.error('CapEx generation failed');
          if (this.uiController) {
            this.uiController.showMessage('Error generating CapEx sheet', 'error');
          }
        }
      }
    } catch (error) {
      console.error('Error in generateCapExSheet:', error);
      if (this.uiController) {
        this.uiController.showMessage('Unexpected error: ' + error.message, 'error');
      }
    }
  }

  async generateDebtModelSheet() {
    console.log('Starting Debt Model sheet generation...');
    
    try {
      // Collect model data
      let modelData = {};
      if (this.formHandler) {
        modelData = this.formHandler.collectAllModelData();
        console.log('Model data for Debt Model:', modelData);
      }
      
      // Generate Debt Model sheet
      if (this.excelGenerator) {
        const result = await this.excelGenerator.generateDebtModelSheet(modelData);
        
        if (result && result.success !== false) {
          console.log('Debt Model sheet generated successfully');
          
          // Show the FCF generation button
          const generateFCFBtn = document.getElementById('generateFCFBtn');
          if (generateFCFBtn) {
            generateFCFBtn.style.display = 'inline-flex';
          }
          
          if (this.uiController) {
            this.uiController.showMessage('Debt Financing created! You can now generate the Free Cash Flow.', 'success');
          }
        } else {
          console.error('Debt Model generation failed');
          if (this.uiController) {
            this.uiController.showMessage('Error generating Debt Financing sheet', 'error');
          }
        }
      } else {
        console.error('ExcelGenerator not available');
        if (this.uiController) {
          this.uiController.showMessage('Excel generator not available. Please refresh the page.', 'error');
        }
      }
    } catch (error) {
      console.error('Error in generateDebtModelSheet:', error);
      if (this.uiController) {
        this.uiController.showMessage('Unexpected error: ' + error.message, 'error');
      }
    }
  }

  async generateFCFWithAI() {
    console.log('Starting FCF generation...');
    
    try {
      // Collect model data
      let modelData = {};
      if (this.formHandler) {
        modelData = this.formHandler.collectAllModelData();
        console.log('Model data for FCF:', modelData);
      }
      
      // Generate FCF using AI
      if (this.excelGenerator) {
        const result = await this.excelGenerator.generateFCFWithAI(modelData);
        
        if (result.success) {
          console.log('FCF generated successfully');
          if (this.uiController) {
            this.uiController.showMessage('Free Cash Flow Statement created! Check the Free Cash Flow sheet.', 'success');
          } else {
            console.log('Free Cash Flow Statement created successfully!');
            if (this.uiController) {
              this.uiController.showMessage('Free Cash Flow Statement created successfully! Check the Free Cash Flow sheet in Excel.', 'success');
            }
          }
          
          // IRR & MOIC are now calculated automatically in the FCF sheet
        } else {
          console.error('FCF generation failed:', result.error);
          if (this.uiController) {
            this.uiController.showMessage('Error generating FCF: ' + result.error, 'error');
          } else {
            console.error('Error generating FCF:', result.error);
            if (this.uiController) {
              this.uiController.showMessage('Error generating FCF: ' + result.error, 'error');
            }
          }
        }
      } else {
        console.error('Debug info:', {
          excelGenerator: this.excelGenerator,
          ExcelGeneratorClass: typeof ExcelGenerator,
          windowExcelGenerator: typeof window.ExcelGenerator
        });
        console.error('Excel generator not available. Please refresh the page.');
        if (this.uiController) {
          this.uiController.showMessage('Excel generator not available. Please refresh the page.', 'error');
        }
      }
      
    } catch (error) {
      console.error('Error in generateFCFWithAI:', error);
      if (this.uiController) {
        this.uiController.showMessage('Unexpected error: ' + error.message, 'error');
      } else {
        console.error('Unexpected error:', error.message);
        if (this.uiController) {
          this.uiController.showMessage('Unexpected error: ' + error.message, 'error');
        }
      }
    }
  }
  
  // IRR & MOIC calculations are now automatically included in the FCF sheet
  // This method is no longer needed
  /*
  async generateMultiplesAndIRR() {
    // This functionality has been moved to the FCF sheet
    // IRR and MOIC are calculated automatically using Excel's built-in functions
    console.log('IRR & MOIC are now calculated automatically in the FCF sheet');
  }
  */

  async addFinalResults() {
    console.log('🎯 Adding final results to Assumptions sheet...');
    
    try {
      // Add IRR and MOIC results to Assumptions sheet
      if (this.excelGenerator) {
        const result = await this.excelGenerator.addIRRMOICToAssumptions();
        
        if (result && result.success !== false) {
          console.log('Final results added successfully to Assumptions sheet');
          return { success: true, message: 'Final results added to Assumptions sheet' };
        } else {
          console.error('Failed to add final results');
          return { success: false, error: 'Failed to add final results' };
        }
      } else {
        console.error('Excel generator not available');
        return { success: false, error: 'Excel generator not available' };
      }
      
    } catch (error) {
      console.error('Error in addFinalResults:', error);
      return { success: false, error: error.message };
    }
  }
  
  async generateFullModel() {
    console.log('🚀 Starting unified full model generation...');
    
    try {
      // Validate form data first
      if (this.formHandler) {
        const validation = this.formHandler.validateAllFields();
        if (!validation.isValid) {
          const errorMessage = 'Please complete the following required fields: ' + validation.errors.join(', ');
          console.log('Validation failed:', validation.errors);
          if (this.uiController) {
            this.uiController.showMessage(errorMessage, 'error');
          }
          return;
        }
      }
      
      // Get the button and progress indicator
      const generateBtn = document.getElementById('generateFullModelBtn');
      const progressIndicator = document.getElementById('generationProgress');
      const progressText = document.getElementById('progressText');
      const progressBar = document.getElementById('progressBar');
      
      // Initialize progress UI
      if (generateBtn) {
        generateBtn.disabled = true;
        generateBtn.innerHTML = `
          <svg class="spinner" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M21 12a9 9 0 11-6.219-8.56"/>
          </svg>
          Generating Model...
        `;
      }
      
      if (progressIndicator) {
        progressIndicator.style.display = 'block';
      }
      
      const steps = [
        { name: 'Assumptions', func: () => this.generateAssumptions(), progress: 16 },
        { name: 'P&L Statement', func: () => this.generatePLWithAI(), progress: 33 },
        { name: 'CapEx Summary', func: () => this.generateCapExSheet(), progress: 50 },
        { name: 'Debt Model', func: () => this.generateDebtModelSheet(), progress: 66 },
        { name: 'Free Cash Flow', func: () => this.generateFCFWithAI(), progress: 83 },
        { name: 'Final Results', func: () => this.addFinalResults(), progress: 100 }
      ];
      
      // Execute each step sequentially
      for (let i = 0; i < steps.length; i++) {
        const step = steps[i];
        
        // Update progress UI
        if (progressText) {
          progressText.textContent = `Generating ${step.name}... (${i + 1}/${steps.length})`;
        }
        if (progressBar) {
          progressBar.style.width = `${step.progress}%`;
        }
        
        console.log(`🔄 Step ${i + 1}/${steps.length}: Generating ${step.name}...`);
        
        // Execute the step
        const result = await step.func();
        
        // Check if step failed (some functions don't return explicit success/failure)
        // We'll continue even if one step has issues, but log it
        if (result && result.success === false) {
          console.warn(`⚠️ Step ${step.name} may have encountered issues, but continuing...`);
        }
        
        console.log(`✅ Step ${i + 1}/${steps.length}: ${step.name} completed`);
        
        // Small delay between steps to ensure Excel operations complete
        if (i < steps.length - 1) {
          await new Promise(resolve => setTimeout(resolve, 1000));
        }
      }
      
      // Model generation completed
      console.log('🎉 Full model generation completed successfully!');
      
      // Update button to "Refresh Full Model"
      if (generateBtn) {
        generateBtn.disabled = false;
        generateBtn.innerHTML = `
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="23,4 23,10 17,10"></polyline>
            <path d="M20.49,15a9,9,0,1,1-2.12-9.36L23,10"></path>
          </svg>
          Refresh Entire Model
        `;
      }
      
      // Hide progress indicator
      if (progressIndicator) {
        setTimeout(() => {
          progressIndicator.style.display = 'none';
        }, 2000);
      }
      
      // Show success message
      if (this.uiController) {
        this.uiController.showMessage('🎉 Complete M&A financial model generated successfully! All sheets created with formulas and calculations.', 'success');
      }
      
    } catch (error) {
      console.error('❌ Error in generateFullModel:', error);
      
      // Reset button state
      const generateBtn = document.getElementById('generateFullModelBtn');
      if (generateBtn) {
        generateBtn.disabled = false;
        generateBtn.innerHTML = `
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
            <line x1="3" y1="9" x2="21" y2="9"></line>
            <line x1="9" y1="21" x2="9" y2="9"></line>
          </svg>
          Generate Full Model
        `;
      }
      
      // Hide progress indicator
      const progressIndicator = document.getElementById('generationProgress');
      if (progressIndicator) {
        progressIndicator.style.display = 'none';
      }
      
      if (this.uiController) {
        this.uiController.showMessage('Error generating full model: ' + error.message, 'error');
      }
    }
  }
  
  async generateFullModelWithProgress() {
    console.log('🚀 Starting full model generation with header progress...');
    
    try {
      // Validate form data first
      if (this.formHandler) {
        const validation = this.formHandler.validateAllFields();
        if (!validation.isValid) {
          const errorMessage = 'Please complete the following required fields: ' + validation.errors.join(', ');
          console.log('Validation failed:', validation.errors);
          if (this.uiController) {
            this.uiController.showMessage(errorMessage, 'error');
          }
          throw new Error(errorMessage);
        }
      }
      
      // Get progress elements from the header
      const progressFill = document.getElementById('generationProgressFill');
      const statusElement = document.getElementById('generationStatus');
      
      const steps = [
        { name: 'Assumptions', func: () => this.generateAssumptions(), progress: 16 },
        { name: 'P&L Statement', func: () => this.generatePLWithAI(), progress: 33 },
        { name: 'CapEx Summary', func: () => this.generateCapExSheet(), progress: 50 },
        { name: 'Debt Model', func: () => this.generateDebtModelSheet(), progress: 66 },
        { name: 'Free Cash Flow', func: () => this.generateFCFWithAI(), progress: 83 },
        { name: 'Final Results', func: () => this.addFinalResults(), progress: 100 }
      ];
      
      // Execute each step sequentially
      for (let i = 0; i < steps.length; i++) {
        const step = steps[i];
        
        // Update progress UI in header
        if (statusElement) {
          statusElement.textContent = `Generating ${step.name}... (${i + 1}/${steps.length})`;
        }
        if (progressFill) {
          progressFill.style.width = `${step.progress}%`;
        }
        
        console.log(`🔄 Step ${i + 1}/${steps.length}: Generating ${step.name}...`);
        
        // Execute the step
        const result = await step.func();
        
        // Check if step failed 
        if (result && result.success === false) {
          console.warn(`⚠️ Step ${step.name} may have encountered issues, but continuing...`);
        }
        
        console.log(`✅ Step ${i + 1}/${steps.length}: ${step.name} completed`);
        
        // Small delay between steps to ensure Excel operations complete
        if (i < steps.length - 1) {
          await new Promise(resolve => setTimeout(resolve, 1000));
        }
      }
      
      // Update final status
      if (statusElement) {
        statusElement.textContent = 'Model generation completed successfully!';
      }
      
      // Model generation completed
      console.log('🎉 Full model generation with progress completed successfully!');
      
      // Show success message
      if (this.uiController) {
        this.uiController.showMessage('🎉 Complete M&A financial model generated successfully! All sheets created with formulas and calculations.', 'success');
      }
      
    } catch (error) {
      console.error('❌ Error in generateFullModelWithProgress:', error);
      
      // Update status to show error
      const statusElement = document.getElementById('generationStatus');
      if (statusElement) {
        statusElement.textContent = 'Model generation failed. Please try again.';
      }
      
      if (this.uiController) {
        this.uiController.showMessage('Error generating full model: ' + error.message, 'error');
      }
      
      throw error; // Re-throw so the HTML can handle the error state
    }
  }
  
  // Legacy function for backward compatibility
  async generateModel() {
    console.log('Legacy generateModel called - redirecting to generateFullModel...');
    return this.generateFullModel();
    
    try {
      // Validate form data
      if (this.formHandler) {
        const validation = this.formHandler.validateAllFields();
        if (!validation.isValid) {
          const errorMessage = 'Please complete the following required fields: ' + validation.errors.join(', ');
          console.log('Validation failed:', validation.errors);
          if (this.uiController) {
            this.uiController.showMessage(errorMessage, 'error');
          }
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
            console.log('Excel model generated successfully!');
            if (this.uiController) {
              this.uiController.showMessage('Excel model generated successfully!', 'success');
            }
          }
        } else {
          console.error('Model generation failed:', result.error);
          if (this.uiController) {
            this.uiController.showMessage('Error generating model: ' + result.error, 'error');
          } else {
            console.error('Error generating model:', result.error);
            if (this.uiController) {
              this.uiController.showMessage('Error generating model: ' + result.error, 'error');
            }
          }
        }
      } else {
        console.error('Debug info:', {
          excelGenerator: this.excelGenerator,
          ExcelGeneratorClass: typeof ExcelGenerator,
          windowExcelGenerator: typeof window.ExcelGenerator
        });
        console.error('Excel generator not available. Please refresh the page.');
        if (this.uiController) {
          this.uiController.showMessage('Excel generator not available. Please refresh the page.', 'error');
        }
      }
      
    } catch (error) {
      console.error('Error in generateModel:', error);
      if (this.uiController) {
        this.uiController.showMessage('Unexpected error: ' + error.message, 'error');
      } else {
        console.error('Unexpected error:', error.message);
        if (this.uiController) {
          this.uiController.showMessage('Unexpected error: ' + error.message, 'error');
        }
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
            console.log('All required fields are completed!');
            if (this.uiController) {
              this.uiController.showMessage('All required fields are completed! Ready to generate model.', 'success');
            }
          }
        } else {
          const errorMessage = 'Missing required fields:\n' + validation.errors.join('\n');
          if (this.uiController) {
            this.uiController.showMessage('Validation failed. Check console for details.', 'warning');
          }
          console.log('Validation errors:', validation.errors);
          console.log('Validation errors:', validation.errors);
          if (this.uiController) {
            this.uiController.showMessage(errorMessage, 'error');
          }
        }
      } else {
        console.error('FormHandler not available');
        console.error('Form validation not available. Please refresh the page.');
        if (this.uiController) {
          this.uiController.showMessage('Form validation not available. Please refresh the page.', 'error');
        }
      }
      
    } catch (error) {
      console.error('Error in validateModel:', error);
      if (this.uiController) {
        this.uiController.showMessage('Error during validation: ' + error.message, 'error');
      } else {
        console.error('Error during validation:', error.message);
        if (this.uiController) {
          this.uiController.showMessage('Error during validation: ' + error.message, 'error');
        }
      }
    }
  }

  autoFillTestData() {
    console.log('⚠️ autoFillTestData() function disabled - hardcoded values not allowed');
    console.log('🤖 Please use the AI AutoFill functionality instead');
    return;
    
    try {
      // Helper function to generate random numbers within a range
      const randomBetween = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
      const randomDecimal = (min, max, decimals = 1) => 
        parseFloat((Math.random() * (max - min) + min).toFixed(decimals));
      
      // Helper function to get future date
      const getFutureDate = (monthsFromNow) => {
        const date = new Date();
        date.setMonth(date.getMonth() + monthsFromNow);
        return date.toISOString().split('T')[0];
      };
      
      // Company names for random selection
      const companyNames = [
        'TechCorp Solutions', 'InnovateCo Ltd', 'DataDrive Systems', 'CloudTech Enterprises',
        'FinTech Innovations', 'MedTech Partners', 'GreenEnergy Co', 'RetailMax Group',
        'LogisticsPro Inc', 'ManufacturingPlus'
      ];
      
      // 1. HIGH-LEVEL PARAMETERS
      document.getElementById('currency').value = ['USD', 'EUR', 'GBP'][randomBetween(0, 2)];
      document.getElementById('projectStartDate').value = getFutureDate(1);
      document.getElementById('projectEndDate').value = getFutureDate(randomBetween(25, 61)); // 2-5 years
      document.getElementById('modelPeriods').value = ['monthly', 'quarterly'][randomBetween(0, 1)];
      
      // 2. DEAL ASSUMPTIONS
      const dealValue = randomBetween(50000000, 500000000); // $50M - $500M
      document.getElementById('dealName').value = companyNames[randomBetween(0, companyNames.length - 1)] + ' Acquisition';
      document.getElementById('dealValue').value = dealValue;
      document.getElementById('transactionFee').value = randomDecimal(1.5, 3.5);
      document.getElementById('dealLTV').value = randomBetween(60, 80);
      
      // 3. EXIT ASSUMPTIONS
      document.getElementById('disposalCost').value = randomDecimal(1.5, 3.5);
      document.getElementById('terminalCapRate').value = randomDecimal(6.5, 12.0);
      document.getElementById('discountRate').value = randomDecimal(8.0, 15.0);
      
      // Clear existing items first
      if (this.formHandler) {
        this.formHandler.clearAndResetAllItems();
      }
      
      // 4. REVENUE ITEMS (2-4 items)
      const revenueCount = randomBetween(2, 4);
      const revenueNames = ['Product Sales', 'Service Revenue', 'Subscription Fees', 'Licensing Income', 'Consulting Revenue'];
      
      for (let i = 0; i < revenueCount; i++) {
        if (this.formHandler) {
          this.formHandler.addRevenueItem();
          
          // Add small delay to ensure DOM elements are created
          setTimeout(() => {
            const nameField = document.getElementById(`revenueName_${i + 1}`);
            const valueField = document.getElementById(`revenueValue_${i + 1}`);
            const growthField = document.getElementById(`revenueGrowthRate_${i + 1}`);
            
            if (nameField) nameField.value = revenueNames[i % revenueNames.length];
            if (valueField) valueField.value = randomBetween(5000000, 50000000);
            if (growthField) growthField.value = randomDecimal(3.0, 15.0);
          }, 100 * (i + 1));
        }
      }
      
      // 5. OPERATING EXPENSES (3-5 items)
      const opexCount = randomBetween(3, 5);
      const opexNames = ['Staff Costs', 'Marketing & Sales', 'Technology & IT', 'Office & Admin', 'Professional Services'];
      
      for (let i = 0; i < opexCount; i++) {
        if (this.formHandler) {
          this.formHandler.addOperatingExpense();
          
          setTimeout(() => {
            const nameField = document.getElementById(`opExName_${i + 1}`);
            const valueField = document.getElementById(`opExValue_${i + 1}`);
            const growthField = document.getElementById(`opExGrowthRate_${i + 1}`);
            
            if (nameField) nameField.value = opexNames[i % opexNames.length];
            if (valueField) valueField.value = randomBetween(1000000, 15000000);
            if (growthField) growthField.value = randomDecimal(2.0, 8.0);
          }, 100 * (i + 1));
        }
      }
      
      // 6. CAPITAL INVESTMENTS (1-3 items)
      const capexCount = randomBetween(1, 3);
      const capexNames = ['Property Improvements', 'HVAC Upgrades', 'Landscaping', 'Building Maintenance'];
      
      for (let i = 0; i < capexCount; i++) {
        if (this.formHandler) {
          this.formHandler.addCapEx();
          
          setTimeout(() => {
            const nameField = document.getElementById(`capExName_${i + 1}`);
            const valueField = document.getElementById(`capExValue_${i + 1}`);
            const growthRateField = document.getElementById(`capExGrowthRate_${i + 1}`);
            
            if (nameField) nameField.value = capexNames[i % capexNames.length];
            if (valueField) valueField.value = randomBetween(50000, 500000);
            if (growthRateField) growthRateField.value = randomDecimal(2.0, 5.0);
          }, 100 * (i + 1));
        }
      }
      
      // Trigger calculations after a delay to ensure all fields are populated
      setTimeout(() => {
        if (this.formHandler) {
          this.formHandler.triggerCalculations();
        }
        
        if (this.uiController) {
          this.uiController.showMessage('✅ Test data auto-filled successfully! All fields populated with realistic M&A numbers.', 'success');
        }
        
        console.log('✅ Test data auto-fill completed');
      }, 1000);
      
    } catch (error) {
      console.error('❌ Error auto-filling test data:', error);
      if (this.uiController) {
        this.uiController.showMessage('Error auto-filling test data: ' + error.message, 'error');
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

// Fallback collapsible initialization for immediate functionality
document.addEventListener('DOMContentLoaded', () => {
  console.log('DOM loaded - setting up immediate collapsible functionality');
  
  // Simple click handler for all collapsible sections
  document.querySelectorAll('.collapsible-section .section-header').forEach(header => {
    header.style.cursor = 'pointer';
    header.addEventListener('click', function(e) {
      e.preventDefault();
      const section = this.closest('.collapsible-section');
      if (section) {
        section.classList.toggle('collapsed');
        console.log(`Toggled section: ${section.id}, collapsed: ${section.classList.contains('collapsed')}`);
      }
    });
  });
});

// Export for debugging
window.MAModelingAddin = MAModelingAddin;