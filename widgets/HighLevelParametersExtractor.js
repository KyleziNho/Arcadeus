class HighLevelParametersExtractor {
  constructor() {
    this.isInitialized = false;
  }

  initialize() {
    if (this.isInitialized) return;
    this.isInitialized = true;
    console.log('✅ HighLevelParametersExtractor initialized');
  }

  // Create AI prompt for extracting high-level parameters
  createExtractionPrompt(fileContents) {
    return `
You are an AI assistant specialized in extracting high-level financial parameters from M&A documents and financial reports.

TASK: Extract the following high-level parameters from the provided document(s):

REQUIRED PARAMETERS:
1. CURRENCY: Look for currency symbols, currency codes (USD, EUR, GBP, etc.), or currency mentions
2. START DATE: Look for project start dates, acquisition dates, model start dates, or transaction closing dates
3. END DATE: Look for project end dates, exit dates, holding period end dates, or investment horizon dates  
4. PERIODS: Look for reporting frequency - daily, monthly, quarterly, or yearly reporting/projections

SEARCH PATTERNS TO LOOK FOR:
- Currency: "$", "USD", "€", "EUR", "£", "GBP", "¥", "JPY", currency symbols in numbers
- Start Date: "acquisition date", "closing date", "start date", "project commencement", "transaction date"
- End Date: "exit date", "end date", "maturity", "investment horizon", "holding period", "project completion"
- Periods: "monthly", "quarterly", "annual", "yearly", "daily", "reporting frequency", "projection periods"

DOCUMENT CONTENT:
${fileContents.map(file => `
File: ${file.name} (${file.type})
Content: ${file.content}
`).join('\n---\n')}

INSTRUCTIONS:
1. Analyze all provided documents thoroughly
2. Look for explicit mentions of dates, currencies, and reporting periods
3. Infer reasonable values if exact matches aren't found
4. If dates are found, calculate the holding period in months
5. Default to USD if no currency is specified
6. Default to monthly periods if no frequency is specified
7. Use reasonable business dates if none are found (e.g., current year to current year + 5)

RESPONSE FORMAT (JSON only):
{
  "currency": "USD|EUR|GBP|JPY|etc",
  "projectStartDate": "YYYY-MM-DD",
  "projectEndDate": "YYYY-MM-DD", 
  "modelPeriods": "daily|monthly|quarterly|yearly",
  "confidence": {
    "currency": 0.0-1.0,
    "projectStartDate": 0.0-1.0,
    "projectEndDate": 0.0-1.0,
    "modelPeriods": 0.0-1.0
  },
  "sources": {
    "currency": "source text where found",
    "projectStartDate": "source text where found",
    "projectEndDate": "source text where found", 
    "modelPeriods": "source text where found"
  }
}

Return ONLY the JSON response, no other text.`;
  }

  // Extract high-level parameters (now reads from standardized data)
  async extractParameters(standardizedData = null) {
    console.log('🔍 Extracting high-level parameters...');
    
    try {
      let extractedData;
      
      // Check if we have standardized data from master analysis
      if (standardizedData) {
        console.log('🔍 Using standardized data from master analysis...');
        extractedData = this.extractFromStandardizedData(standardizedData);
      } else {
        console.log('🔍 No standardized data available, using fallback...');
        extractedData = this.getFallbackParameters();
      }
      
      // Validate and clean the extracted data
      const validatedData = this.validateParameters(extractedData);
      
      console.log('🔍 Extracted high-level parameters:', validatedData);
      return validatedData;
      
    } catch (error) {
      console.error('Error extracting high-level parameters:', error);
      // Return fallback parameters
      return this.getFallbackParameters();
    }
  }

  // Extract parameters from standardized data table
  extractFromStandardizedData(standardizedData) {
    console.log('🔍 Extracting from standardized data:', standardizedData);
    
    try {
      const transaction = standardizedData.transactionDetails || {};
      const projections = standardizedData.projectionAssumptions || {};
      
      return {
        currency: transaction.currency || 'USD',
        projectStartDate: transaction.closingDate || new Date().toISOString().split('T')[0],
        projectEndDate: transaction.expectedExitDate || new Date(Date.now() + 5*365*24*60*60*1000).toISOString().split('T')[0],
        modelPeriods: projections.reportingFrequency || 'monthly',
        confidence: {
          currency: transaction.currency ? 0.9 : 0.3,
          projectStartDate: transaction.closingDate ? 0.9 : 0.3,
          projectEndDate: transaction.expectedExitDate ? 0.9 : 0.3,
          modelPeriods: projections.reportingFrequency ? 0.9 : 0.3
        },
        sources: {
          currency: 'extracted from master analysis',
          projectStartDate: 'extracted from master analysis',
          projectEndDate: 'extracted from master analysis',
          modelPeriods: 'extracted from master analysis'
        }
      };
      
    } catch (error) {
      console.error('Error extracting from standardized data:', error);
      return this.getFallbackParameters();
    }
  }

  // Intelligent parsing fallback when AI service is unavailable
  intelligentParameterParsing(fileContents) {
    console.log('🔍 Using intelligent parsing for parameter extraction...');
    
    const allText = fileContents.map(f => f.content).join(' ').toLowerCase();
    
    // Extract currency
    const currency = this.extractCurrency(allText);
    
    // Extract dates
    const dates = this.extractDates(allText);
    
    // Extract periods
    const periods = this.extractPeriods(allText);
    
    return {
      currency: currency,
      projectStartDate: dates.startDate,
      projectEndDate: dates.endDate,
      modelPeriods: periods,
      confidence: {
        currency: currency === 'USD' ? 0.5 : 0.8, // Lower confidence for default USD
        projectStartDate: dates.startConfidence,
        projectEndDate: dates.endConfidence,
        modelPeriods: periods === 'monthly' ? 0.5 : 0.8 // Lower confidence for default monthly
      },
      sources: {
        currency: currency === 'USD' ? 'default assumption' : 'extracted from document',
        projectStartDate: dates.startSource,
        projectEndDate: dates.endSource,
        modelPeriods: periods === 'monthly' ? 'default assumption' : 'extracted from document'
      }
    };
  }

  // Extract currency from text
  extractCurrency(text) {
    // Look for explicit currency codes
    const currencyPatterns = [
      /\b(USD|US\$|US Dollar|Dollar)\b/i,
      /\b(EUR|Euro|€)\b/i,
      /\b(GBP|British Pound|Pound Sterling|£)\b/i,
      /\b(JPY|Japanese Yen|¥)\b/i,
      /\b(CAD|Canadian Dollar)\b/i,
      /\b(AUD|Australian Dollar)\b/i,
      /\b(CHF|Swiss Franc)\b/i,
      /\b(CNY|Chinese Yuan|RMB)\b/i
    ];
    
    for (const pattern of currencyPatterns) {
      const match = text.match(pattern);
      if (match) {
        const currency = match[1].toUpperCase();
        if (currency.includes('USD') || currency.includes('DOLLAR')) return 'USD';
        if (currency.includes('EUR') || currency.includes('EURO')) return 'EUR';
        if (currency.includes('GBP') || currency.includes('POUND')) return 'GBP';
        if (currency.includes('JPY') || currency.includes('YEN')) return 'JPY';
        if (currency.includes('CAD')) return 'CAD';
        if (currency.includes('AUD')) return 'AUD';
        if (currency.includes('CHF')) return 'CHF';
        if (currency.includes('CNY') || currency.includes('RMB')) return 'CNY';
        return currency.substr(0, 3); // Return first 3 characters
      }
    }
    
    // Look for currency symbols in financial figures
    if (text.includes('$') && !text.includes('€') && !text.includes('£')) {
      return 'USD';
    }
    if (text.includes('€')) return 'EUR';
    if (text.includes('£')) return 'GBP';
    if (text.includes('¥')) return 'JPY';
    
    return 'USD'; // Default
  }

  // Extract dates from text
  extractDates(text) {
    const datePatterns = [
      // ISO format: YYYY-MM-DD
      /(\d{4})-(\d{1,2})-(\d{1,2})/g,
      // US format: MM/DD/YYYY
      /(\d{1,2})\/(\d{1,2})\/(\d{4})/g,
      // European format: DD/MM/YYYY
      /(\d{1,2})\/(\d{1,2})\/(\d{4})/g,
      // Long format: Month DD, YYYY
      /(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2}),?\s+(\d{4})/gi
    ];
    
    const foundDates = [];
    
    for (const pattern of datePatterns) {
      let match;
      while ((match = pattern.exec(text)) !== null) {
        let dateObj;
        
        if (pattern.source.includes('january|february')) {
          // Long format
          const month = new Date(Date.parse(match[1] + " 1, 2000")).getMonth();
          dateObj = new Date(parseInt(match[3]), month, parseInt(match[2]));
        } else if (pattern.source.includes('\\d{4}-')) {
          // ISO format
          dateObj = new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
        } else {
          // MM/DD/YYYY or DD/MM/YYYY - assume US format for now
          dateObj = new Date(parseInt(match[3]), parseInt(match[1]) - 1, parseInt(match[2]));
        }
        
        if (dateObj.getFullYear() >= 2020 && dateObj.getFullYear() <= 2040) {
          foundDates.push(dateObj);
        }
      }
    }
    
    // Sort dates and pick reasonable start/end dates
    foundDates.sort((a, b) => a - b);
    
    let startDate, endDate;
    let startConfidence = 0.3, endConfidence = 0.3;
    let startSource = 'estimated', endSource = 'estimated';
    
    if (foundDates.length >= 2) {
      startDate = foundDates[0];
      endDate = foundDates[foundDates.length - 1];
      startConfidence = 0.8;
      endConfidence = 0.8;
      startSource = 'extracted from document';
      endSource = 'extracted from document';
    } else if (foundDates.length === 1) {
      startDate = foundDates[0];
      endDate = new Date(startDate.getFullYear() + 5, startDate.getMonth(), startDate.getDate());
      startConfidence = 0.8;
      endConfidence = 0.3;
      startSource = 'extracted from document';
      endSource = 'estimated 5 years from start';
    } else {
      // Default to current year to 5 years from now
      const now = new Date();
      startDate = new Date(now.getFullYear(), 0, 1); // January 1st of current year
      endDate = new Date(now.getFullYear() + 5, 11, 31); // December 31st, 5 years from now
      startSource = 'default assumption';
      endSource = 'default assumption';
    }
    
    return {
      startDate: startDate.toISOString().split('T')[0],
      endDate: endDate.toISOString().split('T')[0],
      startConfidence: startConfidence,
      endConfidence: endConfidence,
      startSource: startSource,
      endSource: endSource
    };
  }

  // Extract reporting periods from text
  extractPeriods(text) {
    const periodPatterns = [
      { pattern: /\b(daily|day-to-day|daily reporting)\b/i, value: 'daily' },
      { pattern: /\b(monthly|month-to-month|monthly reporting|per month)\b/i, value: 'monthly' },
      { pattern: /\b(quarterly|quarter|q1|q2|q3|q4|quarterly reporting)\b/i, value: 'quarterly' },
      { pattern: /\b(annually|annual|yearly|year-to-year|per year|per annum)\b/i, value: 'yearly' }
    ];
    
    for (const { pattern, value } of periodPatterns) {
      if (pattern.test(text)) {
        return value;
      }
    }
    
    return 'monthly'; // Default
  }

  // Validate and clean extracted parameters
  validateParameters(data) {
    const validated = {
      currency: this.validateCurrency(data.currency),
      projectStartDate: this.validateDate(data.projectStartDate),
      projectEndDate: this.validateDate(data.projectEndDate),
      modelPeriods: this.validatePeriods(data.modelPeriods),
      confidence: data.confidence || {},
      sources: data.sources || {}
    };
    
    // Ensure end date is after start date
    if (new Date(validated.projectStartDate) >= new Date(validated.projectEndDate)) {
      const startDate = new Date(validated.projectStartDate);
      validated.projectEndDate = new Date(startDate.getFullYear() + 5, startDate.getMonth(), startDate.getDate()).toISOString().split('T')[0];
    }
    
    return validated;
  }

  // Validate currency code
  validateCurrency(currency) {
    const validCurrencies = ['USD', 'EUR', 'GBP', 'JPY', 'CAD', 'AUD', 'CHF', 'CNY', 'SEK', 'NOK'];
    if (validCurrencies.includes(currency)) {
      return currency;
    }
    return 'USD'; // Default
  }

  // Validate date format
  validateDate(dateStr) {
    if (!dateStr) return new Date().toISOString().split('T')[0];
    
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) {
      return new Date().toISOString().split('T')[0];
    }
    
    // Ensure date is reasonable (between 2020 and 2040)
    if (date.getFullYear() < 2020 || date.getFullYear() > 2040) {
      return new Date().toISOString().split('T')[0];
    }
    
    return date.toISOString().split('T')[0];
  }

  // Validate periods
  validatePeriods(periods) {
    const validPeriods = ['daily', 'monthly', 'quarterly', 'yearly'];
    if (validPeriods.includes(periods)) {
      return periods;
    }
    return 'monthly'; // Default
  }

  // Get fallback parameters when extraction fails
  getFallbackParameters() {
    const currentDate = new Date();
    const startDate = new Date(currentDate.getFullYear(), 0, 1);
    const endDate = new Date(currentDate.getFullYear() + 5, 11, 31);
    
    return {
      currency: 'USD',
      projectStartDate: startDate.toISOString().split('T')[0],
      projectEndDate: endDate.toISOString().split('T')[0],
      modelPeriods: 'monthly',
      confidence: {
        currency: 0.3,
        projectStartDate: 0.3,
        projectEndDate: 0.3,
        modelPeriods: 0.3
      },
      sources: {
        currency: 'fallback default',
        projectStartDate: 'fallback default',
        projectEndDate: 'fallback default',
        modelPeriods: 'fallback default'
      }
    };
  }

  // Apply extracted parameters to the form
  async applyParameters(parameters) {
    console.log('🎯 Applying high-level parameters to form...');
    console.log('🎯 Parameters:', parameters);
    
    try {
      // Wait a moment to ensure DOM is ready
      await new Promise(resolve => setTimeout(resolve, 500));
      
      // Check if all elements exist before setting values
      const elements = {
        currency: document.getElementById('currency'),
        projectStartDate: document.getElementById('projectStartDate'),
        projectEndDate: document.getElementById('projectEndDate'),
        modelPeriods: document.getElementById('modelPeriods')
      };
      
      console.log('🎯 Available elements:', {
        currency: !!elements.currency,
        projectStartDate: !!elements.projectStartDate,
        projectEndDate: !!elements.projectEndDate,
        modelPeriods: !!elements.modelPeriods
      });
      
      // Apply each parameter to form with delays
      const results = [];
      
      if (parameters.currency) {
        results.push(this.setInputValue('currency', parameters.currency));
        await new Promise(resolve => setTimeout(resolve, 100));
      }
      
      if (parameters.projectStartDate) {
        results.push(this.setInputValue('projectStartDate', parameters.projectStartDate));
        await new Promise(resolve => setTimeout(resolve, 100));
      }
      
      if (parameters.projectEndDate) {
        results.push(this.setInputValue('projectEndDate', parameters.projectEndDate));
        await new Promise(resolve => setTimeout(resolve, 100));
      }
      
      if (parameters.modelPeriods) {
        results.push(this.setInputValue('modelPeriods', parameters.modelPeriods));
        await new Promise(resolve => setTimeout(resolve, 100));
      }
      
      // Trigger calculation of holding periods
      if (window.formHandler) {
        console.log('🎯 Triggering form calculations...');
        window.formHandler.triggerCalculations();
      }
      
      const successCount = results.filter(Boolean).length;
      console.log(`🎯 Applied ${successCount}/${results.length} high-level parameters successfully`);
      return successCount > 0;
      
    } catch (error) {
      console.error('Error applying high-level parameters:', error);
      return false;
    }
  }

  // Helper method to set input values (Excel Online compatible)
  setInputValue(elementId, value) {
    console.log(`🔧 Setting ${elementId} = ${value}`);
    
    try {
      const element = document.getElementById(elementId);
      console.log(`🔧 Element found:`, !!element, element ? element.tagName : 'null');
      
      if (!element) {
        console.error(`🔧 Element not found: ${elementId}`);
        return false;
      }
      
      if (value === null || value === undefined || value === '') {
        console.log(`🔧 Skipping ${elementId} - no value extracted`);
        return true; // Consider it successful to skip
      }
      
      // Different handling for different input types
      if (element.tagName === 'SELECT') {
        // For select elements, set the selected option
        console.log(`🔧 Setting select element ${elementId}`);
        element.value = value;
        
        // Verify the option exists
        const option = Array.from(element.options).find(opt => opt.value === value);
        if (option) {
          option.selected = true;
          console.log(`🔧 Selected option:`, option.text);
        } else {
          console.warn(`🔧 Option not found for value:`, value);
        }
      } else {
        // For input elements
        console.log(`🔧 Setting input element ${elementId}`);
        element.value = value;
      }
      
      // Force visual update - Excel Online sometimes needs this
      element.focus();
      element.blur();
      
      // Trigger multiple events to ensure change detection
      const events = ['input', 'change', 'blur', 'keyup'];
      events.forEach(eventType => {
        const event = new Event(eventType, { 
          bubbles: true, 
          cancelable: true 
        });
        element.dispatchEvent(event);
      });
      
      // Verify the value was actually set
      const finalValue = element.value;
      console.log(`🔧 Final value for ${elementId}:`, finalValue);
      
      if (finalValue === value.toString()) {
        console.log(`🔧 Successfully set ${elementId}`);
        return true;
      } else {
        console.error(`🔧 Value not set correctly for ${elementId}. Expected: ${value}, Actual: ${finalValue}`);
        return false;
      }
      
    } catch (error) {
      console.error(`🔧 Error setting ${elementId}:`, error);
      return false;
    }
  }

  // Get extraction summary for user feedback
  getExtractionSummary(parameters) {
    const confidenceText = (conf) => {
      if (conf >= 0.8) return 'High';
      if (conf >= 0.5) return 'Medium';
      return 'Low';
    };
    
    return {
      title: 'High-Level Parameters Extracted',
      items: [
        {
          label: 'Currency',
          value: parameters.currency,
          confidence: confidenceText(parameters.confidence?.currency || 0),
          source: parameters.sources?.currency || 'unknown'
        },
        {
          label: 'Start Date',
          value: parameters.projectStartDate,
          confidence: confidenceText(parameters.confidence?.projectStartDate || 0),
          source: parameters.sources?.projectStartDate || 'unknown'
        },
        {
          label: 'End Date',
          value: parameters.projectEndDate,
          confidence: confidenceText(parameters.confidence?.projectEndDate || 0),
          source: parameters.sources?.projectEndDate || 'unknown'
        },
        {
          label: 'Reporting Periods',
          value: parameters.modelPeriods,
          confidence: confidenceText(parameters.confidence?.modelPeriods || 0),
          source: parameters.sources?.modelPeriods || 'unknown'
        }
      ]
    };
  }
}

// Export for use in main application
window.HighLevelParametersExtractor = HighLevelParametersExtractor;