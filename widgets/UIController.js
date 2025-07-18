class UIController {
  constructor() {
    this.isInitialized = false;
    this.sectionMonitoringEnabled = false;
    this.debounceTimer = null;
  }

  initialize() {
    if (this.isInitialized) return;
    
    console.log('Initializing UI controller...');
    
    this.initializeCollapsibleSections();
    this.initializeSectionMonitoring();
    
    this.isInitialized = true;
    console.log('✅ UI controller initialized');
  }

  initializeCollapsibleSections() {
    console.log('Initializing collapsible sections...');
    
    // Add delay to ensure DOM is fully loaded
    setTimeout(() => {
      // High-Level Parameters section
      const minimizeHighLevelBtn = document.getElementById('minimizeHighLevel');
      const highLevelParametersSection = document.getElementById('highLevelParametersSection');
      
      // Deal Assumptions section
      const minimizeAssumptionsBtn = document.getElementById('minimizeAssumptions');
      const dealAssumptionsSection = document.getElementById('dealAssumptionsSection');
      
      // Revenue Items section
      const minimizeRevenueBtn = document.getElementById('minimizeRevenue');
      const revenueItemsSection = document.getElementById('revenueItemsSection');
      
      // Operating Expenses section
      const minimizeOpExBtn = document.getElementById('minimizeOpEx');
      const operatingExpensesSection = document.getElementById('operatingExpensesSection');
      
      // Capital Expenses section
      const minimizeCapExBtn = document.getElementById('minimizeCapEx');
      const capExSection = document.getElementById('capExSection');
      
      // Exit Assumptions section
      const minimizeExitBtn = document.getElementById('minimizeExit');
      const exitAssumptionsSection = document.getElementById('exitAssumptionsSection');
      
      // Debt Model section
      const minimizeDebtBtn = document.getElementById('minimizeDebtModel');
      const debtModelSection = document.getElementById('debtModelSection');
      
      console.log('Found collapsible elements:', {
        highLevel: !!minimizeHighLevelBtn && !!highLevelParametersSection,
        assumptions: !!minimizeAssumptionsBtn && !!dealAssumptionsSection,
        revenue: !!minimizeRevenueBtn && !!revenueItemsSection,
        opEx: !!minimizeOpExBtn && !!operatingExpensesSection,
        capEx: !!minimizeCapExBtn && !!capExSection,
        exit: !!minimizeExitBtn && !!exitAssumptionsSection,
        debt: !!minimizeDebtBtn && !!debtModelSection
      });

      // Initialize each section
      this.initializeSection(minimizeHighLevelBtn, highLevelParametersSection, 'High-Level Parameters');
      this.initializeSection(minimizeAssumptionsBtn, dealAssumptionsSection, 'Deal Assumptions');
      this.initializeSection(minimizeRevenueBtn, revenueItemsSection, 'Revenue Items');
      this.initializeSection(minimizeOpExBtn, operatingExpensesSection, 'Operating Expenses');
      this.initializeSection(minimizeCapExBtn, capExSection, 'Capital Expenses');
      this.initializeSection(minimizeExitBtn, exitAssumptionsSection, 'Exit Assumptions');
      this.initializeSection(minimizeDebtBtn, debtModelSection, 'Debt Model');
      
    }, 100);
  }

  initializeSection(minimizeBtn, section, sectionName) {
    if (minimizeBtn && section) {
      minimizeBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        console.log(`${sectionName} minimize button clicked`);
        this.toggleSection(section, minimizeBtn);
      });
      
      console.log(`✅ ${sectionName} collapsible section initialized successfully`);
      
      // Add click-to-expand functionality for collapsed section
      this.addClickToExpandListener(section, minimizeBtn);
    } else {
      console.error(`❌ Could not find ${sectionName} collapsible section elements`);
    }
  }

  addClickToExpandListener(section, minimizeBtn) {
    if (!section || !minimizeBtn) return;
    
    const header = section.querySelector('h3');
    if (header) {
      header.addEventListener('click', (e) => {
        // Only expand if section is collapsed
        if (section.classList.contains('collapsed')) {
          e.preventDefault();
          e.stopPropagation();
          console.log('Expanding collapsed section via header click');
          this.toggleSection(section, minimizeBtn);
        }
      });
      
      // Add visual feedback for clickable header when collapsed
      header.style.cursor = 'pointer';
    }
  }

  toggleSection(section, minimizeBtn) {
    if (!section || !minimizeBtn) {
      console.error('Missing section or minimize button for toggle');
      return;
    }
    
    const isCurrentlyCollapsed = section.classList.contains('collapsed');
    const sectionContent = section.querySelector('.section-content');
    const minimizeIcon = minimizeBtn.querySelector('.minimize-icon');
    
    console.log('Toggling section:', {
      sectionId: section.id,
      currentlyCollapsed: isCurrentlyCollapsed,
      hasContent: !!sectionContent,
      hasIcon: !!minimizeIcon
    });
    
    if (isCurrentlyCollapsed) {
      // Expand section
      section.classList.remove('collapsed');
      if (minimizeIcon) minimizeIcon.textContent = '−';
      
      if (sectionContent) {
        sectionContent.style.maxHeight = '600px';
        sectionContent.style.opacity = '1';
        sectionContent.style.paddingTop = 'var(--space-6)';
        sectionContent.style.paddingBottom = 'var(--space-6)';
        sectionContent.style.pointerEvents = 'auto';
      }
      
      console.log('Section expanded');
    } else {
      // Collapse section
      section.classList.add('collapsed');
      if (minimizeIcon) minimizeIcon.textContent = '+';
      
      if (sectionContent) {
        sectionContent.style.maxHeight = '0';
        sectionContent.style.opacity = '0';
        sectionContent.style.paddingTop = '0';
        sectionContent.style.paddingBottom = '0';
        sectionContent.style.pointerEvents = 'none';
      }
      
      console.log('Section collapsed');
    }
    
    // Save section state to localStorage
    this.saveSectionState(section.id, !isCurrentlyCollapsed);
  }

  saveSectionState(sectionId, isCollapsed) {
    try {
      const sectionStates = JSON.parse(localStorage.getItem('sectionStates') || '{}');
      sectionStates[sectionId] = isCollapsed;
      localStorage.setItem('sectionStates', JSON.stringify(sectionStates));
    } catch (error) {
      console.error('Error saving section state:', error);
    }
  }

  restoreSectionStates() {
    try {
      const sectionStates = JSON.parse(localStorage.getItem('sectionStates') || '{}');
      
      Object.keys(sectionStates).forEach(sectionId => {
        const section = document.getElementById(sectionId);
        const isCollapsed = sectionStates[sectionId];
        
        if (section && typeof isCollapsed === 'boolean') {
          const minimizeBtn = section.querySelector('.minimize-btn');
          
          if (isCollapsed && !section.classList.contains('collapsed')) {
            this.toggleSection(section, minimizeBtn);
          } else if (!isCollapsed && section.classList.contains('collapsed')) {
            this.toggleSection(section, minimizeBtn);
          }
        }
      });
      
      console.log('Section states restored');
    } catch (error) {
      console.error('Error restoring section states:', error);
    }
  }

  initializeSectionMonitoring() {
    console.log('Initializing section completion monitoring...');
    
    // Monitor form changes for completion indicators
    document.addEventListener('input', (e) => {
      this.debounceUpdateSectionIndicators();
    });
    
    document.addEventListener('change', (e) => {
      this.debounceUpdateSectionIndicators();
    });
    
    // Initial update
    setTimeout(() => {
      this.updateSectionIndicators();
    }, 1000);
    
    this.sectionMonitoringEnabled = true;
    console.log('✅ Section monitoring initialized');
  }

  debounceUpdateSectionIndicators() {
    if (this.debounceTimer) {
      clearTimeout(this.debounceTimer);
    }
    
    this.debounceTimer = setTimeout(() => {
      this.updateSectionIndicators();
    }, 500);
  }

  updateSectionIndicators() {
    if (!this.sectionMonitoringEnabled) return;
    
    try {
      // Update each section's completion status
      this.updateSectionIndicator('highLevelIndicator', 'highLevelParametersSection', this.checkHighLevelParametersCompletion());
      this.updateSectionIndicator('assumptionsIndicator', 'dealAssumptionsSection', this.checkDealAssumptionsCompletion());
      this.updateSectionIndicator('revenueIndicator', 'revenueItemsSection', this.checkRevenueItemsCompletion());
      this.updateSectionIndicator('opExIndicator', 'operatingExpensesSection', this.checkOperatingExpensesCompletion());
      this.updateSectionIndicator('capExIndicator', 'capExSection', this.checkCapitalExpensesCompletion());
      this.updateSectionIndicator('exitIndicator', 'exitAssumptionsSection', this.checkExitAssumptionsCompletion());
      this.updateSectionIndicator('debtIndicator', 'debtModelSection', this.checkDebtModelCompletion());
      
    } catch (error) {
      console.error('Error updating section indicators:', error);
    }
  }

  updateSectionIndicator(indicatorId, sectionId, completionStatus) {
    const indicator = document.getElementById(indicatorId);
    const section = document.getElementById(sectionId);
    
    if (!indicator || !section) return;
    
    // Remove existing completion classes
    section.classList.remove('incomplete', 'partial', 'complete');
    
    // Add new completion class
    section.classList.add(completionStatus);
    
    // Update indicator visual
    switch (completionStatus) {
      case 'complete':
        indicator.style.backgroundColor = '#10B981';
        indicator.textContent = '✓';
        break;
      case 'partial':
        indicator.style.backgroundColor = '#F59E0B';
        indicator.textContent = '◐';
        break;
      case 'incomplete':
      default:
        indicator.style.backgroundColor = '#EF4444';
        indicator.textContent = '○';
        break;
    }
  }

  checkHighLevelParametersCompletion() {
    const currency = document.getElementById('currency')?.value;
    const projectStartDate = document.getElementById('projectStartDate')?.value;
    const projectEndDate = document.getElementById('projectEndDate')?.value;
    const modelPeriods = document.getElementById('modelPeriods')?.value;
    
    const requiredFields = [currency, projectStartDate, projectEndDate, modelPeriods];
    const completedFields = requiredFields.filter(field => field && field.trim() !== '').length;
    
    if (completedFields === requiredFields.length) return 'complete';
    if (completedFields > 0) return 'partial';
    return 'incomplete';
  }

  checkDealAssumptionsCompletion() {
    const dealName = document.getElementById('dealName')?.value;
    const dealValue = document.getElementById('dealValue')?.value;
    const transactionFee = document.getElementById('transactionFee')?.value;
    const dealLTV = document.getElementById('dealLTV')?.value;
    
    const requiredFields = [dealName, dealValue, transactionFee, dealLTV];
    const completedFields = requiredFields.filter(field => field && field.trim() !== '').length;
    
    if (completedFields === requiredFields.length) return 'complete';
    if (completedFields > 0) return 'partial';
    return 'incomplete';
  }

  checkRevenueItemsCompletion() {
    const revenueContainer = document.getElementById('revenueItemsContainer');
    if (!revenueContainer) return 'incomplete';
    
    const revenueItems = revenueContainer.querySelectorAll('.revenue-item');
    if (revenueItems.length === 0) return 'incomplete';
    
    let completeItems = 0;
    revenueItems.forEach(item => {
      const nameInput = item.querySelector('input[id*="revenueName"]');
      const valueInput = item.querySelector('input[id*="revenueValue"]');
      
      if (nameInput?.value?.trim() && valueInput?.value && parseFloat(valueInput.value) > 0) {
        completeItems++;
      }
    });
    
    if (completeItems === revenueItems.length) return 'complete';
    if (completeItems > 0) return 'partial';
    return 'incomplete';
  }

  checkOperatingExpensesCompletion() {
    const opExContainer = document.getElementById('operatingExpensesContainer');
    if (!opExContainer) return 'incomplete';
    
    const opExItems = opExContainer.querySelectorAll('.cost-item');
    if (opExItems.length === 0) return 'incomplete';
    
    let completeItems = 0;
    opExItems.forEach(item => {
      const nameInput = item.querySelector('input[id*="opExName"]');
      const valueInput = item.querySelector('input[id*="opExValue"]');
      
      if (nameInput?.value?.trim() && valueInput?.value && parseFloat(valueInput.value) > 0) {
        completeItems++;
      }
    });
    
    if (completeItems === opExItems.length) return 'complete';
    if (completeItems > 0) return 'partial';
    return 'incomplete';
  }

  checkCapitalExpensesCompletion() {
    const capExContainer = document.getElementById('capExContainer');
    if (!capExContainer) return 'incomplete';
    
    const capExItems = capExContainer.querySelectorAll('.cost-item');
    if (capExItems.length === 0) return 'incomplete';
    
    let completeItems = 0;
    capExItems.forEach(item => {
      const nameInput = item.querySelector('input[id*="capExName"]');
      const valueInput = item.querySelector('input[id*="capExValue"]');
      
      if (nameInput?.value?.trim() && valueInput?.value && parseFloat(valueInput.value) > 0) {
        completeItems++;
      }
    });
    
    if (completeItems === capExItems.length) return 'complete';
    if (completeItems > 0) return 'partial';
    return 'incomplete';
  }

  checkExitAssumptionsCompletion() {
    const disposalCost = document.getElementById('disposalCost')?.value;
    const terminalCapRate = document.getElementById('terminalCapRate')?.value;
    
    const requiredFields = [disposalCost, terminalCapRate];
    const completedFields = requiredFields.filter(field => field && field.trim() !== '').length;
    
    if (completedFields === requiredFields.length) return 'complete';
    if (completedFields > 0) return 'partial';
    return 'incomplete';
  }

  checkDebtModelCompletion() {
    const dealLTV = parseFloat(document.getElementById('dealLTV')?.value) || 0;
    
    // If no LTV, debt model is complete (no debt)
    if (dealLTV === 0) return 'complete';
    
    // If LTV > 0, check debt settings
    const loanIssuanceFees = document.getElementById('loanIssuanceFees')?.value;
    const rateTypeChecked = document.querySelector('input[name="rateType"]:checked');
    
    if (loanIssuanceFees && rateTypeChecked) {
      const rateType = rateTypeChecked.value;
      
      if (rateType === 'fixed') {
        const fixedRate = document.getElementById('fixedRate')?.value;
        return fixedRate ? 'complete' : 'partial';
      } else {
        const baseRate = document.getElementById('baseRate')?.value;
        const creditMargin = document.getElementById('creditMargin')?.value;
        return (baseRate && creditMargin) ? 'complete' : 'partial';
      }
    }
    
    return 'partial';
  }

  showMessage(message, type = 'info', duration = 3000) {
    // Create or update a global message display
    let messageElement = document.getElementById('globalMessage');
    if (!messageElement) {
      messageElement = document.createElement('div');
      messageElement.id = 'globalMessage';
      messageElement.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 12px 20px;
        border-radius: 6px;
        font-size: 14px;
        font-weight: 500;
        z-index: 1000;
        transition: all 0.3s ease;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
      `;
      document.body.appendChild(messageElement);
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
      case 'warning':
        messageElement.style.backgroundColor = '#F59E0B';
        messageElement.style.color = 'white';
        break;
      case 'info':
      default:
        messageElement.style.backgroundColor = '#3B82F6';
        messageElement.style.color = 'white';
        break;
    }

    // Show message
    messageElement.style.opacity = '1';
    messageElement.style.transform = 'translateY(0)';

    // Auto-hide after duration
    setTimeout(() => {
      if (messageElement) {
        messageElement.style.opacity = '0';
        messageElement.style.transform = 'translateY(-20px)';
        setTimeout(() => {
          if (messageElement && messageElement.parentNode) {
            messageElement.parentNode.removeChild(messageElement);
          }
        }, 300);
      }
    }, duration);
  }

  toggleAllSections(collapse = true) {
    const sections = document.querySelectorAll('.collapsible-section');
    
    sections.forEach(section => {
      const minimizeBtn = section.querySelector('.minimize-btn');
      const isCurrentlyCollapsed = section.classList.contains('collapsed');
      
      if (collapse && !isCurrentlyCollapsed) {
        this.toggleSection(section, minimizeBtn);
      } else if (!collapse && isCurrentlyCollapsed) {
        this.toggleSection(section, minimizeBtn);
      }
    });
    
    console.log(collapse ? 'All sections collapsed' : 'All sections expanded');
  }

  expandAllSections() {
    this.toggleAllSections(false);
  }

  collapseAllSections() {
    this.toggleAllSections(true);
  }

  getCompletionSummary() {
    const summary = {
      highLevel: this.checkHighLevelParametersCompletion(),
      assumptions: this.checkDealAssumptionsCompletion(),
      revenue: this.checkRevenueItemsCompletion(),
      opEx: this.checkOperatingExpensesCompletion(),
      capEx: this.checkCapitalExpensesCompletion(),
      exit: this.checkExitAssumptionsCompletion(),
      debt: this.checkDebtModelCompletion()
    };
    
    const counts = {
      complete: 0,
      partial: 0,
      incomplete: 0
    };
    
    Object.values(summary).forEach(status => {
      counts[status]++;
    });
    
    return {
      sections: summary,
      counts: counts,
      overall: counts.complete === Object.keys(summary).length ? 'complete' : 
               counts.complete > 0 || counts.partial > 0 ? 'partial' : 'incomplete'
    };
  }
}

// Export for use in main application
window.UIController = UIController;