/**
 * FieldMappingEngine.js - Maps standardized data to form fields
 * Handles field relationships, calculated fields, and data dependencies
 */

class FieldMappingEngine {
  constructor() {
    // Define field mappings
    this.fieldMappings = {
      // High-Level Parameters
      currency: {
        elementId: 'currency',
        type: 'select',
        section: 'highLevelParameters'
      },
      projectStartDate: {
        elementId: 'projectStartDate',
        type: 'date',
        section: 'highLevelParameters'
      },
      projectEndDate: {
        elementId: 'projectEndDate',
        type: 'date',
        section: 'highLevelParameters'
      },
      modelPeriods: {
        elementId: 'modelPeriods',
        type: 'select',
        section: 'highLevelParameters'
      },
      
      // Deal Assumptions
      dealName: {
        elementId: 'dealName',
        type: 'text',
        section: 'dealAssumptions'
      },
      dealValue: {
        elementId: 'dealValue',
        type: 'number',
        section: 'dealAssumptions',
        triggers: ['equityContribution', 'debtFinancing']
      },
      transactionFee: {
        elementId: 'transactionFee',
        type: 'number',
        section: 'dealAssumptions'
      },
      dealLTV: {
        elementId: 'dealLTV',
        type: 'number',
        section: 'dealAssumptions',
        triggers: ['equityContribution', 'debtFinancing']
      },
      
      // Exit Assumptions
      disposalCost: {
        elementId: 'disposalCost',
        type: 'number',
        section: 'exitAssumptions'
      },
      terminalCapRate: {
        elementId: 'terminalCapRate',
        type: 'number',
        section: 'exitAssumptions'
      },
      
      // Debt Model
      loanIssuanceFees: {
        elementId: 'loanIssuanceFees',
        type: 'number',
        section: 'debtModel'
      },
      interestRateType: {
        elementId: 'rateType',
        type: 'radio',
        section: 'debtModel',
        valueMap: {
          'fixed': 'fixed',
          'floating': 'floating'
        }
      },
      interestRate: {
        elementId: 'fixedRate',
        type: 'number',
        section: 'debtModel',
        condition: { field: 'interestRateType', value: 'fixed' }
      },
      baseRate: {
        elementId: 'baseRate',
        type: 'number',
        section: 'debtModel',
        condition: { field: 'interestRateType', value: 'floating' }
      },
      creditMargin: {
        elementId: 'creditMargin',
        type: 'number',
        section: 'debtModel',
        condition: { field: 'interestRateType', value: 'floating' }
      }
    };

    // Calculated field definitions
    this.calculatedFields = {
      equityContribution: {
        elementId: 'equityContribution',
        type: 'calculated',
        formula: (data) => {
          const dealValue = data.dealValue?.value || 0;
          const ltv = data.dealLTV?.value || 0;
          return dealValue * (1 - ltv / 100);
        },
        dependencies: ['dealValue', 'dealLTV']
      },
      debtFinancing: {
        elementId: 'debtFinancing',
        type: 'calculated',
        formula: (data) => {
          const dealValue = data.dealValue?.value || 0;
          const ltv = data.dealLTV?.value || 0;
          return dealValue * (ltv / 100);
        },
        dependencies: ['dealValue', 'dealLTV']
      },
      holdingPeriodsCalculated: {
        elementId: 'holdingPeriodsCalculated',
        type: 'calculated',
        formula: (data) => {
          const start = data.projectStartDate?.value;
          const end = data.projectEndDate?.value;
          const periods = data.modelPeriods?.value || 'monthly';
          
          if (!start || !end) return null;
          
          const startDate = new Date(start);
          const endDate = new Date(end);
          const diffMs = endDate - startDate;
          
          const calculations = {
            daily: Math.ceil(diffMs / (1000 * 60 * 60 * 24)),
            monthly: Math.ceil(diffMs / (1000 * 60 * 60 * 24 * 30.44)),
            quarterly: Math.ceil(diffMs / (1000 * 60 * 60 * 24 * 30.44 * 3)),
            yearly: Math.ceil(diffMs / (1000 * 60 * 60 * 24 * 365.25))
          };
          
          return calculations[periods] || 0;
        },
        dependencies: ['projectStartDate', 'projectEndDate', 'modelPeriods']
      }
    };

    this.appliedMappings = new Map();
    this.changeHistory = [];
  }

  /**
   * Apply standardized data to form fields
   */
  async applyDataToForm(standardizedData, options = {}) {
    console.log('🗺️ FieldMappingEngine: Applying data to form...');
    
    const {
      skipValidation = false,
      animateChanges = true,
      showConfidence = true,
      reviewMode = false
    } = options;

    // Track application results
    const results = {
      successful: [],
      failed: [],
      skipped: [],
      calculated: []
    };

    try {
      // First pass: Apply direct field mappings
      for (const [field, mapping] of Object.entries(this.fieldMappings)) {
        const data = standardizedData[field];
        
        if (!data || data.value === null || data.value === undefined) {
          results.skipped.push({ field, reason: 'no_data' });
          continue;
        }

        // Check conditions
        if (mapping.condition && !this.checkCondition(mapping.condition, standardizedData)) {
          results.skipped.push({ field, reason: 'condition_not_met' });
          continue;
        }

        // Apply mapping
        const success = await this.applyFieldMapping(field, data, mapping, {
          animate: animateChanges,
          showConfidence
        });

        if (success) {
          results.successful.push({ field, value: data.value });
          this.recordChange(field, null, data.value, data.source);
        } else {
          results.failed.push({ field, reason: 'application_failed' });
        }
      }

      // Second pass: Calculate and apply calculated fields
      for (const [field, calc] of Object.entries(this.calculatedFields)) {
        if (this.canCalculate(calc.dependencies, standardizedData)) {
          const value = calc.formula(standardizedData);
          
          if (value !== null && value !== undefined) {
            const element = document.getElementById(calc.elementId);
            if (element) {
              element.value = this.formatCalculatedValue(field, value);
              results.calculated.push({ field, value });
              
              // Trigger change event
              element.dispatchEvent(new Event('change', { bubbles: true }));
            }
          }
        }
      }

      // Third pass: Handle array fields (revenue items, expenses)
      await this.applyArrayFields(standardizedData, results);

      // Validate if requested
      if (!skipValidation) {
        const validation = this.validateApplication(results, standardizedData);
        results.validation = validation;
      }

      // Show review modal if requested
      if (reviewMode) {
        await this.showReviewModal(results, standardizedData);
      }

      console.log('🗺️ Application complete:', results);
      return results;

    } catch (error) {
      console.error('🗺️ Error applying data to form:', error);
      throw error;
    }
  }

  /**
   * Apply a single field mapping
   */
  async applyFieldMapping(field, data, mapping, options) {
    const element = document.getElementById(mapping.elementId);
    if (!element) {
      console.warn(`🗺️ Element not found: ${mapping.elementId}`);
      return false;
    }

    try {
      // Store current value for undo
      const currentValue = this.getElementValue(element, mapping.type);
      
      // Apply value based on type
      switch (mapping.type) {
        case 'select':
          return this.applySelectValue(element, data.value, mapping.valueMap);
          
        case 'radio':
          return this.applyRadioValue(mapping.elementId, data.value, mapping.valueMap);
          
        case 'number':
          return this.applyNumberValue(element, data.value, field, options);
          
        case 'date':
          return this.applyDateValue(element, data.value);
          
        case 'text':
        default:
          return this.applyTextValue(element, data.value, options);
      }
      
    } catch (error) {
      console.error(`🗺️ Error applying field ${field}:`, error);
      return false;
    }
  }

  /**
   * Apply value to select element
   */
  applySelectValue(element, value, valueMap) {
    const mappedValue = valueMap ? valueMap[value] || value : value;
    
    // Check if option exists
    const optionExists = Array.from(element.options).some(opt => opt.value === mappedValue);
    
    if (optionExists) {
      element.value = mappedValue;
      element.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    }
    
    console.warn(`🗺️ Option not found: ${mappedValue}`);
    return false;
  }

  /**
   * Apply value to radio buttons
   */
  applyRadioValue(name, value, valueMap) {
    const mappedValue = valueMap ? valueMap[value] || value : value;
    const radios = document.querySelectorAll(`input[name="${name}"][value="${mappedValue}"]`);
    
    if (radios.length > 0) {
      radios[0].checked = true;
      radios[0].dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    }
    
    return false;
  }

  /**
   * Apply numeric value with formatting
   */
  applyNumberValue(element, value, field, options) {
    let formattedValue = value;
    
    // Don't format percentages - they're already in the right format
    const percentageFields = ['transactionFee', 'dealLTV', 'disposalCost', 'terminalCapRate', 'interestRate'];
    if (!percentageFields.includes(field)) {
      // For currency values, you might want to format differently
      formattedValue = Math.round(value);
    }
    
    element.value = formattedValue;
    
    // Add confidence indicator if requested
    if (options.showConfidence) {
      this.addConfidenceIndicator(element, options.confidence || 0.5);
    }
    
    // Animate if requested
    if (options.animate) {
      element.classList.add('field-updated');
      setTimeout(() => element.classList.remove('field-updated'), 1000);
    }
    
    element.dispatchEvent(new Event('input', { bubbles: true }));
    element.dispatchEvent(new Event('change', { bubbles: true }));
    
    return true;
  }

  /**
   * Apply date value
   */
  applyDateValue(element, value) {
    // Ensure format is YYYY-MM-DD for date inputs
    const dateValue = value.split('T')[0];
    element.value = dateValue;
    element.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  }

  /**
   * Apply text value
   */
  applyTextValue(element, value, options) {
    element.value = value;
    
    if (options.animate) {
      element.classList.add('field-updated');
      setTimeout(() => element.classList.remove('field-updated'), 1000);
    }
    
    element.dispatchEvent(new Event('input', { bubbles: true }));
    element.dispatchEvent(new Event('change', { bubbles: true }));
    
    return true;
  }

  /**
   * Apply array fields (revenue items, expenses)
   */
  async applyArrayFields(standardizedData, results) {
    // Revenue Items
    if (standardizedData.revenueItems?.value) {
      await this.applyItemsArray(
        standardizedData.revenueItems.value,
        'revenue',
        'addRevenueItem',
        results
      );
    }
    
    // Operating Expenses
    if (standardizedData.operatingExpenses?.value) {
      await this.applyItemsArray(
        standardizedData.operatingExpenses.value,
        'opEx',
        'addOperatingExpense',
        results
      );
    }
    
    // Capital Expenses
    if (standardizedData.capitalExpenses?.value) {
      await this.applyItemsArray(
        standardizedData.capitalExpenses.value,
        'capEx',
        'addCapitalExpense',
        results
      );
    }
  }

  /**
   * Apply array of items (generic for revenue/expenses)
   */
  async applyItemsArray(items, prefix, addMethod, results) {
    console.log(`🗺️ Applying ${items.length} ${prefix} items...`);
    
    // Clear existing items first
    const container = document.getElementById(`${prefix}ItemsContainer`) || 
                     document.getElementById(`${prefix === 'opEx' ? 'operatingExpenses' : prefix}Container`);
    
    if (container) {
      container.innerHTML = '';
    }
    
    // Add each item
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      
      // Add new item row
      if (window.formHandler && window.formHandler[addMethod]) {
        window.formHandler[addMethod]();
        
        // Wait for DOM update
        await new Promise(resolve => setTimeout(resolve, 100));
        
        // Apply values
        const fields = {
          [`${prefix}Name_${i + 1}`]: item.name,
          [`${prefix}Value_${i + 1}`]: item.value,
          [`${prefix === 'revenue' ? 'growthType' : prefix + 'GrowthType'}_${i + 1}`]: item.growthType || 'linear',
          [`${prefix === 'revenue' ? 'linearGrowth' : 'linearGrowth_' + prefix}_${i + 1}`]: item.growthType === 'linear' ? item.growthRate : null,
          [`${prefix === 'revenue' ? 'annualGrowth' : 'annualGrowth_' + prefix}_${i + 1}`]: item.growthType === 'compound' ? item.growthRate : null
        };
        
        for (const [fieldId, value] of Object.entries(fields)) {
          if (value !== null && value !== undefined) {
            const element = document.getElementById(fieldId);
            if (element) {
              element.value = value;
              element.dispatchEvent(new Event('change', { bubbles: true }));
            }
          }
        }
        
        results.successful.push({
          field: `${prefix}Item_${i + 1}`,
          value: item.name
        });
      }
    }
  }

  /**
   * Check if a condition is met
   */
  checkCondition(condition, data) {
    const fieldData = data[condition.field];
    return fieldData && fieldData.value === condition.value;
  }

  /**
   * Check if all dependencies have values
   */
  canCalculate(dependencies, data) {
    return dependencies.every(dep => 
      data[dep] && data[dep].value !== null && data[dep].value !== undefined
    );
  }

  /**
   * Format calculated values
   */
  formatCalculatedValue(field, value) {
    if (field === 'holdingPeriodsCalculated') {
      const periods = document.getElementById('modelPeriods')?.value || 'monthly';
      const labels = {
        daily: 'days',
        monthly: 'months',
        quarterly: 'quarters',
        yearly: 'years'
      };
      return `${value} ${labels[periods] || periods}`;
    }
    
    // For currency values
    if (field === 'equityContribution' || field === 'debtFinancing') {
      return Math.round(value).toLocaleString();
    }
    
    return value;
  }

  /**
   * Add confidence indicator to field
   */
  addConfidenceIndicator(element, confidence) {
    // Remove existing indicator
    const existing = element.parentElement.querySelector('.confidence-indicator');
    if (existing) existing.remove();
    
    // Create new indicator
    const indicator = document.createElement('div');
    indicator.className = 'confidence-indicator';
    
    const level = confidence >= 0.8 ? 'high' : confidence >= 0.5 ? 'medium' : 'low';
    indicator.classList.add(`confidence-${level}`);
    
    indicator.innerHTML = `
      <span class="confidence-icon" title="Confidence: ${(confidence * 100).toFixed(0)}%">
        ${this.getConfidenceIcon(level)}
      </span>
    `;
    
    element.parentElement.appendChild(indicator);
  }

  getConfidenceIcon(level) {
    const icons = {
      high: '✓',
      medium: '?',
      low: '!'
    };
    return icons[level] || '?';
  }

  /**
   * Get current element value
   */
  getElementValue(element, type) {
    switch (type) {
      case 'radio':
        const checked = document.querySelector(`input[name="${element.name}"]:checked`);
        return checked ? checked.value : null;
      default:
        return element.value;
    }
  }

  /**
   * Record change for history
   */
  recordChange(field, oldValue, newValue, source) {
    this.changeHistory.push({
      field,
      oldValue,
      newValue,
      source,
      timestamp: new Date().toISOString()
    });
  }

  /**
   * Validate application results
   */
  validateApplication(results, standardizedData) {
    const validation = {
      isValid: true,
      warnings: [],
      errors: []
    };
    
    // Check critical fields
    const criticalFields = ['dealValue', 'dealName', 'projectStartDate', 'projectEndDate'];
    const missingCritical = criticalFields.filter(field => 
      results.skipped.some(s => s.field === field)
    );
    
    if (missingCritical.length > 0) {
      validation.errors.push(`Missing critical fields: ${missingCritical.join(', ')}`);
      validation.isValid = false;
    }
    
    // Check for failed applications
    if (results.failed.length > 0) {
      validation.warnings.push(`Failed to apply ${results.failed.length} fields`);
    }
    
    // Validate data relationships
    if (standardizedData.projectStartDate?.value && standardizedData.projectEndDate?.value) {
      const start = new Date(standardizedData.projectStartDate.value);
      const end = new Date(standardizedData.projectEndDate.value);
      if (start >= end) {
        validation.errors.push('Project end date must be after start date');
        validation.isValid = false;
      }
    }
    
    return validation;
  }

  /**
   * Show review modal (placeholder - implement UI component)
   */
  async showReviewModal(results, standardizedData) {
    console.log('🗺️ Review modal would show here with results:', results);
    // This would integrate with ExtractionReviewModal component
  }

  /**
   * Undo last changes
   */
  undoLastApplication() {
    // Implement undo functionality using changeHistory
    console.log('🗺️ Undo functionality to be implemented');
  }

  /**
   * Get mapping statistics
   */
  getMappingStats() {
    return {
      totalMappings: Object.keys(this.fieldMappings).length,
      appliedMappings: this.appliedMappings.size,
      changeHistory: this.changeHistory.length,
      lastChange: this.changeHistory[this.changeHistory.length - 1] || null
    };
  }
}

// Export for use
window.FieldMappingEngine = FieldMappingEngine;