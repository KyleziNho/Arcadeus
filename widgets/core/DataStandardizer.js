/**
 * DataStandardizer.js - Converts extracted data into standardized formats
 * Handles currency conversion, date formatting, number normalization, etc.
 */

class DataStandardizer {
  constructor() {
    // Currency conversion rates (would be fetched from API in production)
    this.currencyRates = {
      USD: 1.0,
      EUR: 0.85,
      GBP: 0.73,
      JPY: 110.0,
      CAD: 1.25,
      AUD: 1.35,
      CHF: 0.92,
      CNY: 6.45
    };

    // Common date formats to parse
    this.dateFormats = [
      /(\d{4})-(\d{1,2})-(\d{1,2})/, // YYYY-MM-DD
      /(\d{1,2})\/(\d{1,2})\/(\d{4})/, // MM/DD/YYYY or DD/MM/YYYY
      /(\d{1,2})-(\d{1,2})-(\d{4})/, // MM-DD-YYYY or DD-MM-YYYY
      /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+(\d{1,2}),?\s+(\d{4})/i, // Month DD, YYYY
      /(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+(\d{4})/i // DD Month YYYY
    ];

    // Number format patterns
    this.numberFormats = {
      million: { pattern: /(\d+(?:\.\d+)?)\s*(?:million|mil|m)/i, multiplier: 1000000 },
      billion: { pattern: /(\d+(?:\.\d+)?)\s*(?:billion|bil|b)/i, multiplier: 1000000000 },
      thousand: { pattern: /(\d+(?:\.\d+)?)\s*(?:thousand|k)/i, multiplier: 1000 },
      percentage: { pattern: /(\d+(?:\.\d+)?)\s*%/, divider: 100 }
    };
  }

  /**
   * Main standardization method
   */
  standardize(extractedData, targetCurrency = 'USD') {
    console.log('📊 Standardizing extracted data...');
    
    const standardized = {};
    
    for (const [field, data] of Object.entries(extractedData)) {
      if (!data || data.value === null || data.value === undefined) {
        standardized[field] = data;
        continue;
      }

      const value = data.value;
      const fieldType = this.getFieldType(field);
      
      try {
        switch (fieldType) {
          case 'currency':
            standardized[field] = {
              ...data,
              value: this.standardizeCurrency(value, data.currency || 'USD', targetCurrency),
              standardizedCurrency: targetCurrency
            };
            break;
            
          case 'date':
            standardized[field] = {
              ...data,
              value: this.standardizeDate(value),
              originalFormat: value
            };
            break;
            
          case 'percentage':
            standardized[field] = {
              ...data,
              value: this.standardizePercentage(value)
            };
            break;
            
          case 'number':
            standardized[field] = {
              ...data,
              value: this.standardizeNumber(value)
            };
            break;
            
          case 'array':
            standardized[field] = {
              ...data,
              value: this.standardizeArray(value, field)
            };
            break;
            
          default:
            standardized[field] = data;
        }
      } catch (error) {
        console.error(`Error standardizing field ${field}:`, error);
        standardized[field] = {
          ...data,
          standardizationError: error.message
        };
      }
    }
    
    // Add metadata
    standardized._metadata = {
      standardizedAt: new Date().toISOString(),
      targetCurrency: targetCurrency,
      version: '1.0'
    };
    
    return standardized;
  }

  /**
   * Determine field type for standardization
   */
  getFieldType(field) {
    const fieldTypes = {
      // Currency fields
      dealValue: 'currency',
      equityContribution: 'currency',
      debtFinancing: 'currency',
      
      // Date fields
      projectStartDate: 'date',
      projectEndDate: 'date',
      closingDate: 'date',
      expectedExitDate: 'date',
      
      // Percentage fields
      transactionFee: 'percentage',
      dealLTV: 'percentage',
      disposalCost: 'percentage',
      terminalCapRate: 'percentage',
      interestRate: 'percentage',
      baseRate: 'percentage',
      creditMargin: 'percentage',
      loanIssuanceFees: 'percentage',
      
      // Number fields
      holdingPeriods: 'number',
      
      // Array fields
      revenueItems: 'array',
      operatingExpenses: 'array',
      capitalExpenses: 'array'
    };
    
    // Check for array
    if (field.includes('Items') || field.includes('Expenses')) {
      return 'array';
    }
    
    return fieldTypes[field] || 'string';
  }

  /**
   * Standardize currency values
   */
  standardizeCurrency(value, fromCurrency, toCurrency) {
    // Parse the number
    const numericValue = this.parseNumber(value);
    
    if (isNaN(numericValue)) {
      throw new Error(`Invalid currency value: ${value}`);
    }
    
    // Convert currency if needed
    if (fromCurrency !== toCurrency) {
      const fromRate = this.currencyRates[fromCurrency] || 1;
      const toRate = this.currencyRates[toCurrency] || 1;
      return numericValue * (toRate / fromRate);
    }
    
    return numericValue;
  }

  /**
   * Standardize date formats to YYYY-MM-DD
   */
  standardizeDate(value) {
    if (!value) return null;
    
    // Already in correct format?
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
      return value;
    }
    
    // Try parsing with various formats
    for (const format of this.dateFormats) {
      const match = value.match(format);
      if (match) {
        let year, month, day;
        
        if (format.source.includes('Jan|Feb')) {
          // Month name format
          const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 
                            'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
          if (match[1] && isNaN(match[1])) {
            // Month DD, YYYY
            month = monthNames.indexOf(match[1].toLowerCase().substring(0, 3)) + 1;
            day = parseInt(match[2]);
            year = parseInt(match[3]);
          } else {
            // DD Month YYYY
            day = parseInt(match[1]);
            month = monthNames.indexOf(match[2].toLowerCase().substring(0, 3)) + 1;
            year = parseInt(match[3]);
          }
        } else if (format.source.includes('YYYY')) {
          // YYYY-MM-DD format
          year = parseInt(match[1]);
          month = parseInt(match[2]);
          day = parseInt(match[3]);
        } else {
          // MM/DD/YYYY or DD/MM/YYYY - assume US format
          month = parseInt(match[1]);
          day = parseInt(match[2]);
          year = parseInt(match[3]);
          
          // Swap if day > 12 (must be DD/MM format)
          if (day > 12 && month <= 12) {
            [day, month] = [month, day];
          }
        }
        
        // Validate date
        const date = new Date(year, month - 1, day);
        if (!isNaN(date.getTime())) {
          return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        }
      }
    }
    
    // Try native Date parsing as last resort
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      return parsed.toISOString().split('T')[0];
    }
    
    throw new Error(`Cannot parse date: ${value}`);
  }

  /**
   * Standardize percentage values
   */
  standardizePercentage(value) {
    if (typeof value === 'number') {
      // Already a number - check if it needs conversion
      return value > 1 ? value : value * 100;
    }
    
    const str = String(value);
    
    // Check for percentage symbol
    if (str.includes('%')) {
      return parseFloat(str.replace('%', '').trim());
    }
    
    // Check for decimal representation (0.025 = 2.5%)
    const num = parseFloat(str);
    if (!isNaN(num) && num < 1) {
      return num * 100;
    }
    
    return num;
  }

  /**
   * Standardize number formats (handle M, B, K suffixes)
   */
  standardizeNumber(value) {
    if (typeof value === 'number') {
      return value;
    }
    
    const str = String(value).trim();
    
    // Check for suffixes
    for (const [name, format] of Object.entries(this.numberFormats)) {
      const match = str.match(format.pattern);
      if (match) {
        const num = parseFloat(match[1]);
        return format.multiplier ? num * format.multiplier : num / format.divider;
      }
    }
    
    // Parse as regular number
    return this.parseNumber(str);
  }

  /**
   * Parse number from string (handle commas, spaces, etc.)
   */
  parseNumber(value) {
    if (typeof value === 'number') return value;
    
    const str = String(value)
      .replace(/[$€£¥,\s]/g, '') // Remove currency symbols, commas, spaces
      .trim();
    
    return parseFloat(str);
  }

  /**
   * Standardize array fields (revenue items, expenses, etc.)
   */
  standardizeArray(items, fieldName) {
    if (!Array.isArray(items)) return [];
    
    return items.map((item, index) => {
      const standardizedItem = {
        id: `${fieldName}_${index + 1}`,
        name: item.name || `${fieldName} ${index + 1}`,
        value: this.standardizeNumber(item.value || item.initialValue || 0)
      };
      
      // Handle growth rates
      if (item.growthRate !== undefined) {
        standardizedItem.growthRate = this.standardizePercentage(item.growthRate);
      }
      
      if (item.growthType) {
        standardizedItem.growthType = this.standardizeGrowthType(item.growthType);
      }
      
      return standardizedItem;
    });
  }

  /**
   * Standardize growth type values
   */
  standardizeGrowthType(type) {
    const typeMap = {
      'linear': 'linear',
      'compound': 'compound',
      'annual': 'compound',
      'custom': 'custom',
      'flat': 'linear',
      'exponential': 'compound'
    };
    
    const normalized = String(type).toLowerCase().trim();
    return typeMap[normalized] || 'linear';
  }

  /**
   * Validate standardized data
   */
  validateStandardizedData(data) {
    const errors = [];
    
    // Check required fields
    const requiredFields = ['dealValue', 'currency', 'projectStartDate', 'projectEndDate'];
    for (const field of requiredFields) {
      if (!data[field] || data[field].value === null) {
        errors.push(`Missing required field: ${field}`);
      }
    }
    
    // Validate date logic
    if (data.projectStartDate && data.projectEndDate) {
      const start = new Date(data.projectStartDate.value);
      const end = new Date(data.projectEndDate.value);
      if (start >= end) {
        errors.push('Project end date must be after start date');
      }
    }
    
    // Validate percentages
    const percentageFields = ['transactionFee', 'dealLTV', 'disposalCost', 'terminalCapRate'];
    for (const field of percentageFields) {
      if (data[field] && data[field].value !== null) {
        const value = data[field].value;
        if (value < 0 || value > 100) {
          errors.push(`${field} must be between 0 and 100`);
        }
      }
    }
    
    return {
      isValid: errors.length === 0,
      errors: errors
    };
  }

  /**
   * Format standardized values for display
   */
  formatForDisplay(field, value, currency = 'USD') {
    const fieldType = this.getFieldType(field);
    
    switch (fieldType) {
      case 'currency':
        return this.formatCurrency(value, currency);
      case 'percentage':
        return `${value.toFixed(2)}%`;
      case 'date':
        return this.formatDate(value);
      case 'number':
        return this.formatNumber(value);
      default:
        return String(value);
    }
  }

  formatCurrency(value, currency = 'USD') {
    const symbols = {
      USD: '$',
      EUR: '€',
      GBP: '£',
      JPY: '¥',
      CAD: 'C$',
      AUD: 'A$',
      CHF: 'CHF',
      CNY: '¥'
    };
    
    const symbol = symbols[currency] || currency;
    const formatted = value.toLocaleString('en-US', {
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    });
    
    return `${symbol}${formatted}`;
  }

  formatDate(value) {
    const date = new Date(value);
    return date.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  }

  formatNumber(value) {
    if (value >= 1000000000) {
      return `${(value / 1000000000).toFixed(1)}B`;
    } else if (value >= 1000000) {
      return `${(value / 1000000).toFixed(1)}M`;
    } else if (value >= 1000) {
      return `${(value / 1000).toFixed(1)}K`;
    }
    return value.toLocaleString();
  }
}

// Export for use
window.DataStandardizer = DataStandardizer;