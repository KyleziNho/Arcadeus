class FormHandler {
  constructor() {
    this.isInitialized = false;
  }

  initialize() {
    if (this.isInitialized) return;
    
    this.initializeHighLevelParameters();
    this.initializeDealAssumptions();
    this.initializeRevenueItems();
    this.initializeCostItems();
    this.initializeExitAssumptions();
    this.initializeDebtModel();
    this.initializeNumberFormatting();
    
    // Skip clearing items on initialization to preserve user data
    // this.clearAndResetAllItems();
    
    this.isInitialized = true;
  }

  clearAndResetAllItems() {
    // Clear any existing revenue items that might have old growth rate HTML
    const revenueContainer = document.getElementById('revenueItemsContainer');
    if (revenueContainer) {
      revenueContainer.innerHTML = '';
      console.log('Cleared existing revenue items with old structure');
    }

    // Clear any existing operating expense items
    const opExContainer = document.getElementById('operatingExpensesContainer');
    if (opExContainer) {
      opExContainer.innerHTML = '';
      console.log('Cleared existing operating expense items with old structure');
    }

    // Clear any existing capital expense items
    const capExContainer = document.getElementById('capitalExpensesContainer');
    if (capExContainer) {
      capExContainer.innerHTML = '';
      console.log('Cleared existing capital expense items with old structure');
    }
  }

  validateAllFields() {
    const errors = [];
    const requiredFields = [
      // High-Level Parameters
      { id: 'currency', name: 'Currency' },
      { id: 'projectStartDate', name: 'Project Start Date' },
      { id: 'projectEndDate', name: 'Project End Date' },
      { id: 'modelPeriods', name: 'Model Periods' },
      
      // Deal Assumptions
      { id: 'dealName', name: 'Deal Name' },
      { id: 'dealValue', name: 'Deal Value' },
      { id: 'transactionFee', name: 'Transaction Fee' },
      { id: 'dealLTV', name: 'Deal LTV' },
      
      // Exit Assumptions
      { id: 'disposalCost', name: 'Disposal Cost' },
      { id: 'terminalCapRate', name: 'Terminal Cap Rate' },
      { id: 'discountRate', name: 'Discount Rate (WACC)' }
    ];
    
    // Check required fields
    requiredFields.forEach(field => {
      const element = document.getElementById(field.id);
      if (!element || !element.value || element.value.trim() === '') {
        errors.push(`• ${field.name}`);
      }
    });
    
    // Check at least one revenue item exists
    const revenueItems = document.querySelectorAll('.revenue-item');
    if (revenueItems.length === 0) {
      errors.push('• At least one Revenue Item');
    }
    
    // Check at least one cost item exists
    const costItems = document.querySelectorAll('.cost-item');
    if (costItems.length === 0) {
      errors.push('• At least one Cost Item');
    }
    
    return {
      isValid: errors.length === 0,
      errors: errors
    };
  }

  collectAllModelData() {
    console.log('📊 ====== COLLECTING ALL MODEL DATA ======');
    
    const data = {
      // High-Level Parameters
      currency: document.getElementById('currency')?.value || 'USD',
      projectStartDate: document.getElementById('projectStartDate')?.value || '',
      projectEndDate: document.getElementById('projectEndDate')?.value || '',
      modelPeriods: document.getElementById('modelPeriods')?.value || 'monthly',
      holdingPeriodsCalculated: document.getElementById('holdingPeriodsCalculated')?.value || '',
      
      // Deal Assumptions
      dealName: document.getElementById('dealName')?.value || 'Sample Company Ltd.',
      dealValue: parseFloat(this.removeCommas(document.getElementById('dealValue')?.value || '0')) || 0,
      transactionFee: parseFloat(document.getElementById('transactionFee')?.value) || 2.5,
      dealLTV: parseFloat(document.getElementById('dealLTV')?.value) || 70,
      equityContribution: document.getElementById('equityContribution')?.value || '',
      debtFinancing: document.getElementById('debtFinancing')?.value || '',
      
      // Revenue Items
      revenueItems: this.collectRevenueItems(),
      
      // Operating Expenses
      operatingExpenses: this.collectOperatingExpenses(),
      
      // Capital Expenses
      capEx: this.collectCapEx(),
      
      // Exit Assumptions
      disposalCost: parseFloat(document.getElementById('disposalCost')?.value) || 2.5,
      terminalCapRate: parseFloat(document.getElementById('terminalCapRate')?.value) || 8.5,
      discountRate: parseFloat(document.getElementById('discountRate')?.value) || 10.0,
      
      // Debt Model
      hasDebt: this.checkDebtEligibility(),
      debtSettings: this.collectDebtSettings()
    };
    
    console.log('📊 ====== COMPLETE MODEL DATA COLLECTED ======');
    console.log('📊 Revenue Items:', data.revenueItems);
    console.log('📊 Operating Expenses:', data.operatingExpenses);
    console.log('📊 Capital Expenses:', data.capitalExpenses);
    console.log('📊 Complete Data Object:', data);
    
    return data;
  }

  collectRevenueItems() {
    const items = [];
    const revenueContainer = document.getElementById('revenueItemsContainer');
    if (!revenueContainer) return items;
    
    const revenueContainers = revenueContainer.querySelectorAll('.revenue-item');
    
    revenueContainers.forEach((container, index) => {
      // Extract the item number from the container ID or use index + 1
      const itemId = container.id; // e.g., "revenueItem_1"
      let itemNumber = index + 1;
      if (itemId && itemId.includes('_')) {
        itemNumber = parseInt(itemId.split('_')[1]) || (index + 1);
      }
      
      // Use specific IDs instead of generic selectors
      const nameInput = document.getElementById(`revenueName_${itemNumber}`);
      const valueInput = document.getElementById(`revenueValue_${itemNumber}`);
      
      console.log(`📊 Collecting revenue item ${itemNumber}:`, {
        nameInput: nameInput?.value,
        valueInput: valueInput?.value,
        nameElement: !!nameInput,
        valueElement: !!valueInput
      });
      
      if (nameInput && valueInput && nameInput.value && valueInput.value) {
        const growthRateInput = document.getElementById(`revenueGrowthRate_${itemNumber}`);
        const item = {
          name: nameInput.value || `Revenue Item ${itemNumber}`,
          value: parseFloat(this.removeCommas(valueInput.value)) || 0,
          growthRate: growthRateInput ? (parseFloat(growthRateInput.value) || 0) : 0
        };

        console.log(`📊 Revenue item ${itemNumber}:`, item);
        items.push(item);
      }
    });
    
    return items;
  }

  collectOperatingExpenses() {
    const items = [];
    const opExContainer = document.getElementById('operatingExpensesContainer');
    if (!opExContainer) return items;
    
    const costContainers = opExContainer.querySelectorAll('.cost-item');
    
    costContainers.forEach((container, index) => {
      // Extract the item number from the container ID or use index + 1
      const itemId = container.id; // e.g., "opExItem_1"
      let itemNumber = index + 1;
      if (itemId && itemId.includes('_')) {
        itemNumber = parseInt(itemId.split('_')[1]) || (index + 1);
      }
      
      // Use specific IDs instead of generic selectors
      const nameInput = document.getElementById(`opExName_${itemNumber}`);
      const valueInput = document.getElementById(`opExValue_${itemNumber}`);
      
      if (nameInput && valueInput && nameInput.value && valueInput.value) {
        const growthRateInput = document.getElementById(`opExGrowthRate_${itemNumber}`);
        const item = {
          name: nameInput.value || `Operating Expense ${itemNumber}`,
          value: parseFloat(this.removeCommas(valueInput.value)) || 0,
          growthRate: growthRateInput ? (parseFloat(growthRateInput.value) || 0) : 0
        };

        items.push(item);
      }
    });
    
    return items;
  }

  collectCapEx() {
    const items = [];
    const capExContainer = document.getElementById('capExContainer');
    if (!capExContainer) return items;
    
    const costContainers = capExContainer.querySelectorAll('.cost-item');
    
    costContainers.forEach((container, index) => {
      // Extract the item number from the container ID or use index + 1
      const itemId = container.id; // e.g., "capExItem_1"
      let itemNumber = index + 1;
      if (itemId && itemId.includes('_')) {
        itemNumber = parseInt(itemId.split('_')[1]) || (index + 1);
      }
      
      // Use specific IDs instead of generic selectors
      const nameInput = document.getElementById(`capExName_${itemNumber}`);
      const valueInput = document.getElementById(`capExValue_${itemNumber}`);
      const growthRateInput = document.getElementById(`capExGrowthRate_${itemNumber}`);
      
      if (nameInput && valueInput && nameInput.value && valueInput.value) {
        const item = {
          name: nameInput.value || `CapEx ${itemNumber}`,
          value: parseFloat(this.removeCommas(valueInput.value)) || 0,
          growthRate: parseFloat(growthRateInput?.value) || 0,
          type: 'capex'
        };

        items.push(item);
      }
    });
    
    return items;
  }


  collectDebtSettings() {
    const loanIssuanceFees = document.getElementById('loanIssuanceFees')?.value || '1.5';
    
    const settings = {
      loanIssuanceFees: parseFloat(this.removeCommas(loanIssuanceFees)),
      rateType: 'fixed',
      fixedRate: parseFloat(this.removeCommas(document.getElementById('fixedRate')?.value || '5.5'))
    };

    return settings;
  }

  checkDebtEligibility() {
    const dealLTV = parseFloat(this.removeCommas(document.getElementById('dealLTV')?.value || '0')) || 0;
    return dealLTV > 0;
  }

  formatDateForExcel(dateString) {
    if (!dateString) return '';
    const date = new Date(dateString);
    return date.toLocaleDateString();
  }

  calculateHoldingPeriod() {
    const startDate = document.getElementById('projectStartDate')?.value;
    const endDate = document.getElementById('projectEndDate')?.value;
    
    if (!startDate || !endDate) return 24;
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    const monthsDiff = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth());
    
    return Math.max(1, monthsDiff);
  }

  calculatePeriods(startDate, endDate, periodType) {
    if (!startDate || !endDate) return 12;
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    const diffTime = Math.abs(end - start);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    switch (periodType) {
      case 'daily':
        return Math.min(diffDays, 1000); // Increased cap for daily periods
      case 'monthly':
        return Math.ceil(diffDays / 30); // Removed cap for monthly periods
      case 'quarterly':
        return Math.ceil(diffDays / 90); // Removed cap for quarterly periods
      case 'yearly':
        return Math.ceil(diffDays / 365); // Removed cap for yearly periods
      default:
        return 12;
    }
  }

  initializeHighLevelParameters() {
    // Project date change handlers
    const projectStartDate = document.getElementById('projectStartDate');
    const projectEndDate = document.getElementById('projectEndDate');
    const modelPeriods = document.getElementById('modelPeriods');
    const holdingPeriodsCalculated = document.getElementById('holdingPeriodsCalculated');

    const updateHoldingPeriods = () => {
      if (projectStartDate && projectEndDate && modelPeriods && holdingPeriodsCalculated) {
        const startDate = projectStartDate.value;
        const endDate = projectEndDate.value;
        const periodType = modelPeriods.value;

        if (startDate && endDate) {
          const periods = this.calculatePeriods(startDate, endDate, periodType);
          let periodLabel = '';
          
          switch (periodType) {
            case 'daily': periodLabel = 'days'; break;
            case 'monthly': periodLabel = 'months'; break;
            case 'quarterly': periodLabel = 'quarters'; break;
            case 'yearly': periodLabel = 'years'; break;
            default: periodLabel = 'periods';
          }
          
          holdingPeriodsCalculated.value = `${periods} ${periodLabel}`;
        }
      }
    };

    if (projectStartDate) projectStartDate.addEventListener('change', updateHoldingPeriods);
    if (projectEndDate) projectEndDate.addEventListener('change', updateHoldingPeriods);
    if (modelPeriods) modelPeriods.addEventListener('change', updateHoldingPeriods);
  }

  initializeDealAssumptions() {
    const dealValue = document.getElementById('dealValue');
    const dealLTV = document.getElementById('dealLTV');
    const equityContribution = document.getElementById('equityContribution');
    const debtFinancing = document.getElementById('debtFinancing');

    const updateCalculations = () => {
      if (dealValue && dealLTV && equityContribution && debtFinancing) {
        const value = parseFloat(this.removeCommas(dealValue.value)) || 0;
        const ltv = parseFloat(this.removeCommas(dealLTV.value)) || 0;
        
        const debt = value * (ltv / 100);
        const equity = value - debt;
        
        equityContribution.value = this.formatCurrency(equity);
        debtFinancing.value = this.formatCurrency(debt);

        // Update debt eligibility
        this.updateDebtEligibility(ltv);
      }
    };

    if (dealValue) dealValue.addEventListener('input', updateCalculations);
    if (dealLTV) dealLTV.addEventListener('input', updateCalculations);
  }

  initializeRevenueItems() {
    const addRevenueBtn = document.getElementById('addRevenueItem');
    if (addRevenueBtn) {
      addRevenueBtn.addEventListener('click', () => this.addRevenueItem());
    }
  }

  initializeCostItems() {
    const addOpExBtn = document.getElementById('addOperatingExpense');
    const addCapExBtn = document.getElementById('addCapExItem');
    
    if (addOpExBtn) {
      addOpExBtn.addEventListener('click', () => this.addOperatingExpense());
    }
    
    if (addCapExBtn) {
      addCapExBtn.addEventListener('click', () => this.addCapEx());
    }
  }

  initializeExitAssumptions() {
    // Exit assumptions are simple input fields, no special initialization needed
  }

  initializeDebtModel() {
    // Initialize debt model with fixed rate only
    console.log('Debt model initialized with fixed rate option');
  }

  updateDebtEligibility(ltv) {
    const debtSettings = document.getElementById('debtSettings');
    const debtStatusMessage = document.getElementById('debtStatusMessage');

    if (ltv > 0) {
      if (debtSettings) debtSettings.style.display = 'block';
      if (debtStatusMessage) debtStatusMessage.textContent = 'Debt financing options available';
    } else {
      if (debtSettings) debtSettings.style.display = 'none';
      if (debtStatusMessage) debtStatusMessage.textContent = 'Please input a higher LTV to access debt financing options';
    }
  }

  addRevenueItem() {
    const container = document.getElementById('revenueItemsContainer');
    if (!container) return;

    const itemCount = container.children.length + 1;
    const itemId = `revenue_${itemCount}`;

    const itemHTML = `
      <div class="revenue-item" id="revenueItem_${itemCount}">
        <div class="revenue-item-header">
          <span class="revenue-item-title">Revenue Item ${itemCount}</span>
        </div>
        <button class="remove-revenue-item" onclick="this.parentElement.remove()">Remove</button>
        
        <div class="form-group">
          <label for="revenueName_${itemCount}">Revenue Source Name</label>
          <input type="text" id="revenueName_${itemCount}" placeholder="e.g., Product Sales" />
        </div>
        
        <div class="form-group">
          <label for="revenueValue_${itemCount}">Base Value (Year 1)</label>
          <input type="number" id="revenueValue_${itemCount}" placeholder="100000" step="1000" />
        </div>
        
        <div class="form-group">
          <label for="revenueGrowthRate_${itemCount}">Linear Growth Rate (%)</label>
          <input type="number" id="revenueGrowthRate_${itemCount}" placeholder="e.g., 5" step="0.1" />
          <small class="help-text">Annual growth rate (e.g., 5 for 5% growth)</small>
        </div>
      </div>
    `;

    container.insertAdjacentHTML('beforeend', itemHTML);
    
    // Set up number formatting for the newly added value input
    const newValueInput = document.getElementById(`revenueValue_${itemCount}`);
    if (newValueInput) {
      this.setupNumberFormatting(newValueInput);
    }
    
    // Update progress tracking
    if (typeof updateProgress === 'function') {
      updateProgress();
    }
  }

  addOperatingExpense() {
    const container = document.getElementById('operatingExpensesContainer');
    if (!container) return;

    const itemCount = container.children.length + 1;
    const itemId = `opEx_${itemCount}`;

    const itemHTML = `
      <div class="cost-item" id="opExItem_${itemCount}">
        <div class="cost-item-header">
          <span class="cost-item-title">Operating Expense ${itemCount}</span>
        </div>
        <button class="remove-cost-item" onclick="this.parentElement.remove()">Remove</button>
        
        <div class="form-group">
          <label for="opExName_${itemCount}">Expense Name</label>
          <input type="text" id="opExName_${itemCount}" placeholder="e.g., Staff Costs" />
        </div>
        
        <div class="form-group">
          <label for="opExValue_${itemCount}">Annual Value</label>
          <input type="number" id="opExValue_${itemCount}" placeholder="50000" step="1000" />
        </div>
        
        <div class="form-group">
          <label for="opExGrowthRate_${itemCount}">Linear Growth Rate (%)</label>
          <input type="number" id="opExGrowthRate_${itemCount}" placeholder="e.g., 3" step="0.1" />
          <small class="help-text">Annual growth rate (e.g., 3 for 3% growth)</small>
        </div>
      </div>
    `;

    container.insertAdjacentHTML('beforeend', itemHTML);
    
    // Set up number formatting for the newly added value input
    const newValueInput = document.getElementById(`opExValue_${itemCount}`);
    if (newValueInput) {
      this.setupNumberFormatting(newValueInput);
    }
    
    // Update progress tracking
    if (typeof updateProgress === 'function') {
      updateProgress();
    }
  }

  addCapEx() {
    const container = document.getElementById('capExContainer');
    if (!container) return;

    const itemCount = container.children.length + 1;
    const itemId = `capEx_${itemCount}`;

    const itemHTML = `
      <div class="cost-item" id="capExItem_${itemCount}">
        <div class="cost-item-header">
          <span class="cost-item-title">CapEx ${itemCount}</span>
        </div>
        <button class="remove-cost-item" onclick="this.parentElement.remove()">Remove</button>
        
        <div class="form-group">
          <label for="capExName_${itemCount}">CapEx Name</label>
          <input type="text" id="capExName_${itemCount}" placeholder="e.g., Property Improvements" />
        </div>
        
        <div class="form-group">
          <label for="capExValue_${itemCount}">Annual Value</label>
          <input type="number" id="capExValue_${itemCount}" placeholder="50000" step="1000" />
        </div>
        
        <div class="form-group">
          <label for="capExGrowthRate_${itemCount}">Linear Growth Rate (%)</label>
          <input type="number" id="capExGrowthRate_${itemCount}" placeholder="e.g., 3" step="0.1" />
          <small class="help-text">Annual growth rate (e.g., 3 for 3% growth)</small>
        </div>
      </div>
    `;

    container.insertAdjacentHTML('beforeend', itemHTML);
    
    // Set up number formatting for the newly added value input
    const newValueInput = document.getElementById(`capExValue_${itemCount}`);
    if (newValueInput) {
      this.setupNumberFormatting(newValueInput);
    }
    
    // Update progress tracking
    if (typeof updateProgress === 'function') {
      updateProgress();
    }
  }


  formatCurrency(value) {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    }).format(value);
  }

  setInputValue(elementId, value) {
    const element = document.getElementById(elementId);
    if (element && value !== null && value !== undefined) {
      element.value = value;
      element.dispatchEvent(new Event('change', { bubbles: true }));
      element.dispatchEvent(new Event('input', { bubbles: true }));
    }
  }

  triggerCalculations() {
    // Trigger holding period calculations
    const projectStartDate = document.getElementById('projectStartDate');
    if (projectStartDate) {
      projectStartDate.dispatchEvent(new Event('change', { bubbles: true }));
    }

    // Trigger deal assumptions calculations
    const dealValue = document.getElementById('dealValue');
    if (dealValue) {
      dealValue.dispatchEvent(new Event('input', { bubbles: true }));
    }

    // Trigger debt eligibility check
    const dealLTV = document.getElementById('dealLTV');
    if (dealLTV) {
      dealLTV.dispatchEvent(new Event('input', { bubbles: true }));
    }

    // Refresh number formatting for all value fields
    this.refreshNumberFormatting();
  }

  // Number formatting methods
  formatNumberWithCommas(num) {
    if (num === null || num === undefined || num === '') return '';
    return parseFloat(num).toLocaleString('en-US');
  }

  removeCommas(str) {
    return str.replace(/,/g, '');
  }

  initializeNumberFormatting() {
    // Fields that should have comma formatting for large numbers (only text inputs)
    const numberFields = [
      'equityContribution', 
      'debtFinancing'
    ];

    numberFields.forEach(fieldId => {
      const field = document.getElementById(fieldId);
      if (field && field.type === 'text') {
        this.setupNumberFormatting(field);
      }
    });

    // Also format any existing text input value fields that need formatting
    this.formatAllExistingValueFields();

    // Use MutationObserver for better performance (replaces deprecated DOMNodeInserted)
    const observer = new MutationObserver((mutations) => {
      mutations.forEach((mutation) => {
        mutation.addedNodes.forEach((node) => {
          if (node.nodeType === 1) { // Element node
            // Check if the added node itself is a revenue/opex/capex item
            if (node.classList && (node.classList.contains('revenue-item') || 
                node.classList.contains('cost-item'))) {
              const textInputs = node.querySelectorAll('input[type="text"]');
              textInputs.forEach(input => {
                if (input.id.includes('Value') || input.id.includes('value')) {
                  this.setupNumberFormatting(input);
                }
              });
            }
            // Also check for any text inputs within the added node that need formatting
            const textInputs = node.querySelectorAll ? node.querySelectorAll('input[type="text"]') : [];
            textInputs.forEach(input => {
              if (input.id.includes('Value') || input.id.includes('value')) {
                this.setupNumberFormatting(input);
              }
            });
          }
        });
      });
    });

    // Start observing the document for added nodes
    observer.observe(document.body, {
      childList: true,
      subtree: true
    });
  }

  setupNumberFormatting(field) {
    // Skip number formatting entirely for type="number" inputs
    if (field.type === 'number') {
      console.log('⏭️ Skipping number formatting for number input:', field.id);
      return;
    }
    
    if (field.hasAttribute('data-formatted')) return; // Already set up
    field.setAttribute('data-formatted', 'true');
    console.log('🔧 Setting up number formatting for:', field.id);

    // Format existing value immediately if it exists
    if (field.value && !isNaN(field.value) && field.value.trim() !== '') {
      const existingValue = this.removeCommas(field.value);
      if (existingValue && !isNaN(existingValue)) {
        field.value = this.formatNumberWithCommas(existingValue);
      }
    }

    field.addEventListener('blur', (e) => {
      const value = this.removeCommas(e.target.value);
      if (value && !isNaN(value)) {
        e.target.value = this.formatNumberWithCommas(value);
      }
    });

    field.addEventListener('focus', (e) => {
      const value = this.removeCommas(e.target.value);
      if (value && !isNaN(value)) {
        e.target.value = value;
      }
    });

    field.addEventListener('input', (e) => {
      // Remove any non-numeric characters except decimal point and minus sign
      let value = e.target.value.replace(/[^0-9.-]/g, '');
      // Ensure only one decimal point
      const decimalCount = (value.match(/\./g) || []).length;
      if (decimalCount > 1) {
        const parts = value.split('.');
        value = parts[0] + '.' + parts.slice(1).join('');
      }
      // Only update if the value actually changed to avoid cursor jumping
      if (e.target.value !== value) {
        e.target.value = value;
      }
    });
  }

  formatAllExistingValueFields() {
    // Find all text input fields that need number formatting (skip type="number" inputs)
    const valueFields = document.querySelectorAll('input[type="text"]');
    valueFields.forEach(field => {
      if (field.id.includes('Value') || field.id.includes('value')) {
        this.setupNumberFormatting(field);
      }
    });
  }

  // Public method to refresh number formatting (can be called from outside)
  refreshNumberFormatting() {
    this.formatAllExistingValueFields();
  }

}

// Export for use in main application
window.FormHandler = FormHandler;