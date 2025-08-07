class DataManager {
  constructor() {
    this.storageKey = 'maModelingData';
    this.isInitialized = false;
    this.autoSaveTimeout = null;
    this.autoSaveDelay = 2000; // 2 seconds delay after last change
  }

  initialize() {
    if (this.isInitialized) return;
    
    console.log('Initializing data manager with auto-caching...');
    
    // Set up save/load button event listeners
    const saveDataBtn = document.getElementById('saveDataBtn');
    const loadDataBtn = document.getElementById('loadDataBtn');
    
    if (saveDataBtn) {
      saveDataBtn.addEventListener('click', () => this.saveData());
      console.log('Save data button listener added');
    }
    
    if (loadDataBtn) {
      loadDataBtn.addEventListener('click', () => this.loadData());
      console.log('Load data button listener added');
    }
    
    // Set up automatic saving on form changes
    this.setupAutoSave();
    
    // Auto-load saved data immediately after setup
    setTimeout(() => {
      this.autoLoadSavedData();
    }, 500);
    
    this.isInitialized = true;
    console.log('✅ Data manager initialized with auto-caching');
  }

  async saveData() {
    console.log('Saving data...');
    
    try {
      // Collect all form data
      const formData = this.collectAllFormData();
      
      // Save to localStorage
      localStorage.setItem(this.storageKey, JSON.stringify(formData));
      
      // Also save to a backup key with timestamp
      const backupKey = `${this.storageKey}_backup_${Date.now()}`;
      localStorage.setItem(backupKey, JSON.stringify(formData));
      
      // Clean up old backups (keep only the last 5)
      this.cleanupOldBackups();
      
      console.log('Data saved successfully:', formData);
      this.showSaveStatus('Data saved successfully!', 'success');
      
      return { success: true, data: formData };
      
    } catch (error) {
      console.error('Error saving data:', error);
      this.showSaveStatus('Error saving data: ' + error.message, 'error');
      return { success: false, error: error.message };
    }
  }

  async loadData() {
    console.log('Loading data...');
    
    try {
      const savedData = localStorage.getItem(this.storageKey);
      
      if (!savedData) {
        this.showSaveStatus('No saved data found', 'error');
        return { success: false, error: 'No saved data found' };
      }
      
      const formData = JSON.parse(savedData);
      console.log('Loaded data:', formData);
      
      // Populate form with loaded data
      this.populateAllFormData(formData);
      
      // Trigger any necessary recalculations
      if (window.formHandler) {
        window.formHandler.triggerCalculations();
      }
      
      this.showSaveStatus('Data loaded successfully!', 'success');
      
      return { success: true, data: formData };
      
    } catch (error) {
      console.error('Error loading data:', error);
      this.showSaveStatus('Error loading data: ' + error.message, 'error');
      return { success: false, error: error.message };
    }
  }

  collectAllFormData() {
    const data = {};
    
    // Collect all input values
    const inputs = document.querySelectorAll('input, select, textarea');
    inputs.forEach(input => {
      if (input.id) {
        if (input.type === 'checkbox' || input.type === 'radio') {
          data[input.id] = input.checked;
        } else {
          data[input.id] = input.value;
        }
      }
    });
    
    // Collect dynamic items using FormHandler if available
    if (window.formHandler) {
      data.revenueItems = window.formHandler.collectRevenueItems();
      data.operatingExpenses = window.formHandler.collectOperatingExpenses();
      data.capitalExpenses = window.formHandler.collectCapitalExpenses();
    } else {
      // Fallback to direct collection
      data.revenueItems = this.collectRevenueItemsDirectly();
      data.operatingExpenses = this.collectOperatingExpensesDirectly();
      data.capitalExpenses = this.collectCapitalExpensesDirectly();
    }
    
    // Add metadata
    data.savedAt = new Date().toISOString();
    data.version = '1.0';
    
    return data;
  }

  populateAllFormData(data) {
    console.log('Populating form with data:', data);
    
    // Populate basic form fields
    Object.keys(data).forEach(key => {
      const element = document.getElementById(key);
      if (element && !['revenueItems', 'operatingExpenses', 'capitalExpenses', 'savedAt', 'version'].includes(key)) {
        if (element.type === 'checkbox' || element.type === 'radio') {
          element.checked = data[key];
        } else {
          element.value = data[key];
        }
        // Trigger change events
        element.dispatchEvent(new Event('change', { bubbles: true }));
        element.dispatchEvent(new Event('input', { bubbles: true }));
      }
    });
    
    // Populate dynamic items
    if (data.revenueItems) {
      this.populateRevenueItems(data.revenueItems);
    }
    
    if (data.operatingExpenses) {
      this.populateOperatingExpenses(data.operatingExpenses);
    }
    
    if (data.capitalExpenses) {
      this.populateCapitalExpenses(data.capitalExpenses);
    }
  }

  populateRevenueItems(items) {
    console.log('Populating revenue items:', items);
    
    // Clear existing items
    const container = document.getElementById('revenueItemsContainer');
    if (container) {
      container.innerHTML = '';
    }
    
    // Add items using FormHandler if available
    if (window.formHandler) {
      items.forEach((item, index) => {
        window.formHandler.addRevenueItem();
        
        // Wait a bit for DOM update
        setTimeout(() => {
          this.setInputValue(`revenueName_${index + 1}`, item.name);
          this.setInputValue(`revenueValue_${index + 1}`, item.value);
          
          if (item.growthType) {
            this.setInputValue(`growthType_${index + 1}`, item.growthType);
            window.formHandler.updateGrowthInputs(`revenue_${index + 1}`, item.growthType);
            
            setTimeout(() => {
              if (item.growthType === 'annual' && item.annualGrowthRate !== undefined) {
                this.setInputValue(`annualGrowth_${index + 1}`, item.annualGrowthRate);
              } else if (item.growthType === 'linear' && item.linearGrowthRate !== undefined) {
                this.setInputValue(`linearGrowth_${index + 1}`, item.linearGrowthRate);
              } else if (item.growth !== undefined) {
                this.setInputValue(`linearGrowth_${index + 1}`, item.growth);
              }
            }, 100);
          }
        }, 100 * index);
      });
    }
  }

  populateOperatingExpenses(items) {
    console.log('Populating operating expenses:', items);
    
    // Clear existing items
    const container = document.getElementById('operatingExpensesContainer');
    if (container) {
      container.innerHTML = '';
    }
    
    // Add items using FormHandler if available
    if (window.formHandler) {
      items.forEach((item, index) => {
        window.formHandler.addOperatingExpense();
        
        setTimeout(() => {
          this.setInputValue(`opExName_${index + 1}`, item.name);
          this.setInputValue(`opExValue_${index + 1}`, item.value);
          
          if (item.growthType) {
            this.setInputValue(`opExGrowthType_${index + 1}`, item.growthType);
            window.formHandler.updateCostGrowthInputs(`opEx_${index + 1}`, item.growthType);
            
            setTimeout(() => {
              if (item.growthType === 'annual' && item.annualGrowthRate !== undefined) {
                this.setInputValue(`annualGrowth_opEx_${index + 1}`, item.annualGrowthRate);
              } else if (item.growthType === 'linear' && item.linearGrowthRate !== undefined) {
                this.setInputValue(`linearGrowth_opEx_${index + 1}`, item.linearGrowthRate);
              } else if (item.growth !== undefined) {
                this.setInputValue(`linearGrowth_opEx_${index + 1}`, item.growth);
              }
            }, 100);
          }
        }, 100 * index);
      });
    }
  }

  populateCapitalExpenses(items) {
    console.log('Populating capital expenses:', items);
    
    // Clear existing items
    const container = document.getElementById('capExContainer');
    if (container) {
      container.innerHTML = '';
    }
    
    // Add items using FormHandler if available
    if (window.formHandler) {
      items.forEach((item, index) => {
        window.formHandler.addCapitalExpense();
        
        setTimeout(() => {
          this.setInputValue(`capExName_${index + 1}`, item.name);
          this.setInputValue(`capExValue_${index + 1}`, item.value);
          
          if (item.growthType) {
            this.setInputValue(`capExGrowthType_${index + 1}`, item.growthType);
            window.formHandler.updateCostGrowthInputs(`capEx_${index + 1}`, item.growthType);
            
            setTimeout(() => {
              if (item.growthType === 'annual' && item.annualGrowthRate !== undefined) {
                this.setInputValue(`annualGrowth_capEx_${index + 1}`, item.annualGrowthRate);
              } else if (item.growthType === 'linear' && item.linearGrowthRate !== undefined) {
                this.setInputValue(`linearGrowth_capEx_${index + 1}`, item.linearGrowthRate);
              } else if (item.growth !== undefined) {
                this.setInputValue(`linearGrowth_capEx_${index + 1}`, item.growth);
              }
            }, 100);
          }
        }, 100 * index);
      });
    }
  }

  // Fallback methods for direct collection if FormHandler not available
  collectRevenueItemsDirectly() {
    const items = [];
    const container = document.getElementById('revenueItemsContainer');
    if (!container) return items;
    
    const revenueItems = container.querySelectorAll('.revenue-item');
    revenueItems.forEach((item, index) => {
      const nameInput = item.querySelector(`input[id*="revenueName"]`);
      const valueInput = item.querySelector(`input[id*="revenueValue"]`);
      
      if (nameInput && valueInput) {
        items.push({
          name: nameInput.value,
          value: parseFloat(valueInput.value) || 0
        });
      }
    });
    
    return items;
  }

  collectOperatingExpensesDirectly() {
    const items = [];
    const container = document.getElementById('operatingExpensesContainer');
    if (!container) return items;
    
    const expenseItems = container.querySelectorAll('.cost-item');
    expenseItems.forEach((item, index) => {
      const nameInput = item.querySelector(`input[id*="opExName"]`);
      const valueInput = item.querySelector(`input[id*="opExValue"]`);
      
      if (nameInput && valueInput) {
        items.push({
          name: nameInput.value,
          value: parseFloat(valueInput.value) || 0
        });
      }
    });
    
    return items;
  }

  collectCapitalExpensesDirectly() {
    const items = [];
    const container = document.getElementById('capExContainer');
    if (!container) return items;
    
    const expenseItems = container.querySelectorAll('.cost-item');
    expenseItems.forEach((item, index) => {
      const nameInput = item.querySelector(`input[id*="capExName"]`);
      const valueInput = item.querySelector(`input[id*="capExValue"]`);
      
      if (nameInput && valueInput) {
        items.push({
          name: nameInput.value,
          value: parseFloat(valueInput.value) || 0
        });
      }
    });
    
    return items;
  }

  setInputValue(elementId, value) {
    const element = document.getElementById(elementId);
    if (element && value !== null && value !== undefined) {
      element.value = value;
      element.dispatchEvent(new Event('change', { bubbles: true }));
      element.dispatchEvent(new Event('input', { bubbles: true }));
    }
  }

  showSaveStatus(message, type) {
    const statusElement = document.getElementById('saveStatus');
    if (statusElement) {
      statusElement.textContent = message;
      statusElement.className = `save-status ${type}`;
      
      // Clear message after 3 seconds
      setTimeout(() => {
        statusElement.textContent = '';
        statusElement.className = 'save-status';
      }, 3000);
    }
    
    console.log(`Save status (${type}):`, message);
  }

  cleanupOldBackups() {
    try {
      const backupKeys = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.startsWith(`${this.storageKey}_backup_`)) {
          backupKeys.push(key);
        }
      }
      
      // Sort by timestamp (newest first)
      backupKeys.sort((a, b) => {
        const timestampA = parseInt(a.split('_').pop());
        const timestampB = parseInt(b.split('_').pop());
        return timestampB - timestampA;
      });
      
      // Remove old backups (keep only the 5 newest)
      if (backupKeys.length > 5) {
        for (let i = 5; i < backupKeys.length; i++) {
          localStorage.removeItem(backupKeys[i]);
        }
        console.log(`Cleaned up ${backupKeys.length - 5} old backups`);
      }
    } catch (error) {
      console.error('Error cleaning up old backups:', error);
    }
  }

  exportData() {
    try {
      const data = this.collectAllFormData();
      const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      
      const a = document.createElement('a');
      a.href = url;
      a.download = `ma-model-data-${new Date().toISOString().split('T')[0]}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      
      URL.revokeObjectURL(url);
      
      this.showSaveStatus('Data exported successfully!', 'success');
      return { success: true };
    } catch (error) {
      console.error('Error exporting data:', error);
      this.showSaveStatus('Error exporting data: ' + error.message, 'error');
      return { success: false, error: error.message };
    }
  }

  async importData(file) {
    try {
      const text = await this.readFileAsText(file);
      const data = JSON.parse(text);
      
      this.populateAllFormData(data);
      
      // Trigger calculations
      if (window.formHandler) {
        window.formHandler.triggerCalculations();
      }
      
      this.showSaveStatus('Data imported successfully!', 'success');
      return { success: true, data: data };
    } catch (error) {
      console.error('Error importing data:', error);
      this.showSaveStatus('Error importing data: ' + error.message, 'error');
      return { success: false, error: error.message };
    }
  }

  readFileAsText(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = (e) => reject(new Error('Failed to read file'));
      reader.readAsText(file);
    });
  }

  setupAutoSave() {
    console.log('Setting up automatic saving...');
    
    // Debounced auto-save function
    const debouncedAutoSave = () => {
      clearTimeout(this.autoSaveTimeout);
      this.autoSaveTimeout = setTimeout(() => {
        this.autoSaveData();
      }, this.autoSaveDelay);
    };
    
    // Listen to all form changes
    document.addEventListener('input', debouncedAutoSave);
    document.addEventListener('change', debouncedAutoSave);
    
    // Also listen for dynamic item additions/removals
    const observer = new MutationObserver((mutations) => {
      let hasRelevantChange = false;
      
      mutations.forEach((mutation) => {
        if (mutation.type === 'childList') {
          // Check if revenue, operating expense, or capital expense items were added/removed
          mutation.addedNodes.forEach((node) => {
            if (node.nodeType === 1 && (
              node.classList?.contains('revenue-item') ||
              node.classList?.contains('cost-item') ||
              node.querySelector?.('.revenue-item') ||
              node.querySelector?.('.cost-item')
            )) {
              hasRelevantChange = true;
            }
          });
          
          mutation.removedNodes.forEach((node) => {
            if (node.nodeType === 1 && (
              node.classList?.contains('revenue-item') ||
              node.classList?.contains('cost-item') ||
              node.querySelector?.('.revenue-item') ||
              node.querySelector?.('.cost-item')
            )) {
              hasRelevantChange = true;
            }
          });
        }
      });
      
      if (hasRelevantChange) {
        debouncedAutoSave();
      }
    });
    
    // Observe the entire document for dynamic changes
    observer.observe(document.body, {
      childList: true,
      subtree: true
    });
    
    console.log('✅ Auto-save listeners established');
  }
  
  async autoSaveData() {
    try {
      const formData = this.collectAllFormData();
      localStorage.setItem(this.storageKey, JSON.stringify(formData));
      console.log('📦 Data auto-saved');
      
      // Show subtle notification
      this.showAutoSaveStatus();
      
      return { success: true, data: formData };
    } catch (error) {
      console.warn('Auto-save failed:', error);
      return { success: false, error: error.message };
    }
  }
  
  showAutoSaveStatus() {
    // Create or update auto-save indicator
    let indicator = document.getElementById('autoSaveIndicator');
    if (!indicator) {
      indicator = document.createElement('div');
      indicator.id = 'autoSaveIndicator';
      indicator.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #28a745;
        color: white;
        padding: 8px 12px;
        border-radius: 4px;
        font-size: 12px;
        z-index: 10000;
        opacity: 0;
        transition: opacity 0.3s ease;
        pointer-events: none;
      `;
      indicator.textContent = '✓ Saved';
      document.body.appendChild(indicator);
    }
    
    // Show and hide the indicator
    indicator.style.opacity = '1';
    setTimeout(() => {
      indicator.style.opacity = '0';
    }, 1500);
  }

  autoLoadSavedData() {
    console.log('Checking for saved data to auto-load...');
    
    try {
      const savedData = localStorage.getItem(this.storageKey);
      
      if (savedData) {
        const formData = JSON.parse(savedData);
        console.log('Found saved data, auto-loading...');
        
        // Load ALL data including dynamic items
        this.populateAllFormData(formData);
        
        // Trigger recalculations
        if (window.formHandler) {
          window.formHandler.triggerCalculations();
        }
        
        console.log('✅ Auto-load completed - all form data restored');
        
        // Show restoration notification
        const notification = document.createElement('div');
        notification.style.cssText = `
          position: fixed;
          top: 20px;
          left: 50%;
          transform: translateX(-50%);
          background: #17a2b8;
          color: white;
          padding: 12px 20px;
          border-radius: 6px;
          font-size: 14px;
          z-index: 10000;
          box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        `;
        notification.textContent = '📋 Previous form data restored';
        document.body.appendChild(notification);
        
        setTimeout(() => {
          notification.remove();
        }, 3000);
        
      } else {
        console.log('No saved data found');
      }
    } catch (error) {
      console.warn('Error auto-loading saved data:', error);
    }
  }

  getSavedDataInfo() {
    try {
      const savedData = localStorage.getItem(this.storageKey);
      if (!savedData) {
        return { exists: false };
      }
      
      const data = JSON.parse(savedData);
      return {
        exists: true,
        savedAt: data.savedAt,
        version: data.version || 'unknown',
        hasRevenueItems: data.revenueItems && data.revenueItems.length > 0,
        hasOperatingExpenses: data.operatingExpenses && data.operatingExpenses.length > 0,
        hasCapitalExpenses: data.capitalExpenses && data.capitalExpenses.length > 0
      };
    } catch (error) {
      console.error('Error getting saved data info:', error);
      return { exists: false, error: error.message };
    }
  }

  clearSavedData() {
    try {
      localStorage.removeItem(this.storageKey);
      
      // Also remove all backups
      const keysToRemove = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.startsWith(`${this.storageKey}_backup_`)) {
          keysToRemove.push(key);
        }
      }
      
      keysToRemove.forEach(key => localStorage.removeItem(key));
      
      this.showSaveStatus('All saved data cleared!', 'success');
      return { success: true };
    } catch (error) {
      console.error('Error clearing saved data:', error);
      this.showSaveStatus('Error clearing data: ' + error.message, 'error');
      return { success: false, error: error.message };
    }
  }
}

// Export for use in main application
window.DataManager = DataManager;