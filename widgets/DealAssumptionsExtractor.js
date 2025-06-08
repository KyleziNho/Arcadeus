class DealAssumptionsExtractor {
  constructor() {
    this.isInitialized = false;
  }

  initialize() {
    if (this.isInitialized) return;
    this.isInitialized = true;
    console.log('✅ DealAssumptionsExtractor initialized');
  }

  // Create AI prompt for extracting deal assumptions
  createExtractionPrompt(fileContents) {
    return `
You are an AI assistant specialized in extracting deal assumptions and transaction details from M&A documents and financial reports.

TASK: Extract the following deal assumptions from the provided document(s):

REQUIRED PARAMETERS:
1. DEAL NAME: Company name, transaction name, or deal identifier
2. DEAL VALUE: Total transaction value/purchase price (extract number and currency)
3. TRANSACTION FEE: Investment banking fees, advisory fees, transaction costs (as percentage)
4. DEAL LTV: Loan-to-Value ratio, debt-to-value ratio, leverage ratio (as percentage)

SEARCH PATTERNS TO LOOK FOR:
- Deal Name: "company name", "target company", "acquisition of", "transaction name", "deal name", business names
- Deal Value: "purchase price", "transaction value", "deal value", "acquisition price", "enterprise value", "total consideration"
- Transaction Fee: "transaction fee", "advisory fee", "investment banking fee", "closing costs", "transaction costs", "fees"
- Deal LTV: "LTV", "loan-to-value", "debt ratio", "leverage ratio", "debt financing", "debt-to-value", "financing ratio"

DOCUMENT CONTENT:
${fileContents.map(file => `
File: ${file.name} (${file.type})
Content: ${file.content}
`).join('\n---\n')}

INSTRUCTIONS:
1. Analyze all provided documents thoroughly for deal-specific information
2. Look for explicit transaction details, company names, and financial figures
3. Extract percentages for fees and LTV ratios (convert to decimal if needed)
4. For deal value, extract the numerical amount and identify the currency
5. Deal names can be company names, transaction names, or target company names
6. Default to reasonable business assumptions if exact values aren't found
7. Transaction fees typically range from 1-5% for M&A deals
8. LTV ratios typically range from 50-80% for leveraged transactions

RESPONSE FORMAT (JSON only):
{
  "dealName": "string - company or transaction name",
  "dealValue": number - numerical value only,
  "dealCurrency": "USD|EUR|GBP|etc - currency of deal value",
  "transactionFee": number - percentage as decimal (e.g., 2.5 for 2.5%),
  "dealLTV": number - percentage as decimal (e.g., 70 for 70%),
  "confidence": {
    "dealName": 0.0-1.0,
    "dealValue": 0.0-1.0,
    "transactionFee": 0.0-1.0,
    "dealLTV": 0.0-1.0
  },
  "sources": {
    "dealName": "source text where found",
    "dealValue": "source text where found",
    "transactionFee": "source text where found",
    "dealLTV": "source text where found"
  }
}

Return ONLY the JSON response, no other text.`;
  }

  // Extract deal assumptions (now reads from standardized data)
  async extractDealAssumptions(standardizedData = null) {
    console.log('🔍 Extracting deal assumptions...');
    
    try {
      let extractedData;
      
      // Check if we have standardized data from master analysis
      if (standardizedData) {
        console.log('🔍 Using standardized data from master analysis...');
        extractedData = this.extractFromStandardizedData(standardizedData);
      } else {
        console.log('🔍 No standardized data available, using fallback...');
        extractedData = this.getFallbackDealAssumptions();
      }
      
      // Validate and clean the extracted data
      const validatedData = this.validateDealAssumptions(extractedData);
      
      // Calculate derived values
      const completeData = this.calculateDerivedValues(validatedData);
      
      console.log('🔍 Extracted deal assumptions:', completeData);
      return completeData;
      
    } catch (error) {
      console.error('Error extracting deal assumptions:', error);
      // Return fallback assumptions
      return this.getFallbackDealAssumptions();
    }
  }

  // Extract deal assumptions from standardized data table
  extractFromStandardizedData(standardizedData) {
    console.log('🔍 Extracting deal assumptions from standardized data:', standardizedData);
    
    try {
      const company = standardizedData.companyOverview || {};
      const transaction = standardizedData.transactionDetails || {};
      const financing = standardizedData.financingStructure || {};
      
      return {
        dealName: transaction.dealName || company.companyName || 'Unknown Deal',
        dealValue: transaction.dealValue || financing.totalDealValue || 50000000,
        dealCurrency: transaction.currency || 'USD',
        transactionFee: transaction.transactionFees || 2.5,
        dealLTV: financing.debtLTV || 70,
        confidence: {
          dealName: (transaction.dealName || company.companyName) ? 0.9 : 0.3,
          dealValue: transaction.dealValue ? 0.9 : 0.3,
          transactionFee: transaction.transactionFees ? 0.9 : 0.3,
          dealLTV: financing.debtLTV ? 0.9 : 0.3
        },
        sources: {
          dealName: 'extracted from master analysis',
          dealValue: 'extracted from master analysis',
          transactionFee: 'extracted from master analysis',
          dealLTV: 'extracted from master analysis'
        }
      };
      
    } catch (error) {
      console.error('Error extracting from standardized data:', error);
      return this.getFallbackDealAssumptions();
    }
  }

  // Intelligent parsing fallback when AI service is unavailable
  intelligentDealParsing(fileContents) {
    console.log('🔍 Using intelligent parsing for deal assumptions extraction...');
    
    const allText = fileContents.map(f => f.content).join(' ');
    
    // Extract deal name
    const dealName = this.extractDealName(allText, fileContents);
    
    // Extract deal value
    const dealValue = this.extractDealValue(allText);
    
    // Extract transaction fee
    const transactionFee = this.extractTransactionFee(allText);
    
    // Extract deal LTV
    const dealLTV = this.extractDealLTV(allText);
    
    return {
      dealName: dealName.value,
      dealValue: dealValue.value,
      dealCurrency: dealValue.currency,
      transactionFee: transactionFee.value,
      dealLTV: dealLTV.value,
      confidence: {
        dealName: dealName.confidence,
        dealValue: dealValue.confidence,
        transactionFee: transactionFee.confidence,
        dealLTV: dealLTV.confidence
      },
      sources: {
        dealName: dealName.source,
        dealValue: dealValue.source,
        transactionFee: transactionFee.source,
        dealLTV: dealLTV.source
      }
    };
  }

  // Extract deal name from text
  extractDealName(text, fileContents) {
    // Try to extract from filename first
    for (const file of fileContents) {
      const filename = file.name.replace(/\.(csv|pdf|png|jpg|jpeg)$/i, '');
      if (filename.length > 3 && !filename.toLowerCase().includes('data') && !filename.toLowerCase().includes('test')) {
        return {
          value: filename.replace(/[-_]/g, ' '),
          confidence: 0.7,
          source: `extracted from filename: ${file.name}`
        };
      }
    }
    
    // Look for company names in text
    const companyPatterns = [
      /(?:acquisition of|target company|company name|deal name)[\s:]+([a-zA-Z\s&.,]+?)(?:\n|$|,)/gi,
      /([A-Z][a-zA-Z\s&.,]+(?:Inc|Corp|LLC|Ltd|Company|Co\.|Corporation|Limited))/g,
      /deal name[\s:]+([a-zA-Z\s&.,]+)/gi,
      /transaction[\s:]+([a-zA-Z\s&.,]+)/gi
    ];
    
    for (const pattern of companyPatterns) {
      const match = text.match(pattern);
      if (match) {
        let name = match[1] || match[0];
        name = name.trim().replace(/[,\n\r]/g, '').substring(0, 50);
        if (name.length > 3) {
          return {
            value: name,
            confidence: 0.8,
            source: 'extracted from document text'
          };
        }
      }
    }
    
    return {
      value: 'Sample M&A Transaction',
      confidence: 0.3,
      source: 'default assumption'
    };
  }

  // Extract deal value from text
  extractDealValue(text) {
    // Look for monetary amounts with currency
    const valuePatterns = [
      // Deal value, purchase price, transaction value patterns
      /(?:deal value|purchase price|transaction value|acquisition price|enterprise value|total consideration)[\s:$€£¥]*([0-9,]+(?:\.[0-9]+)?)\s*(?:million|billion|thousand|m|b|k)?\s*(USD|EUR|GBP|JPY|\$|€|£|¥)?/gi,
      // Large monetary amounts
      /[\$€£¥]\s*([0-9,]+(?:\.[0-9]+)?)\s*(?:million|billion|thousand|m|b|k)?/gi,
      // Numbers followed by currency
      /([0-9,]+(?:\.[0-9]+)?)\s*(?:million|billion|thousand|m|b|k)?\s*(USD|EUR|GBP|JPY|\$|€|£|¥)/gi
    ];
    
    let bestMatch = { value: 0, currency: 'USD', confidence: 0.3, source: 'default assumption' };
    
    for (const pattern of valuePatterns) {
      let match;
      while ((match = pattern.exec(text)) !== null) {
        let value = parseFloat(match[1].replace(/,/g, ''));
        const currency = this.identifyCurrency(match[2] || match[3] || '$');
        
        // Handle multipliers
        const fullMatch = match[0].toLowerCase();
        if (fullMatch.includes('billion') || fullMatch.includes(' b')) {
          value *= 1000000000;
        } else if (fullMatch.includes('million') || fullMatch.includes(' m')) {
          value *= 1000000;
        } else if (fullMatch.includes('thousand') || fullMatch.includes(' k')) {
          value *= 1000;
        }
        
        // Prefer larger, more reasonable deal values (1M - 100B range)
        if (value >= 1000000 && value <= 100000000000 && value > bestMatch.value) {
          bestMatch = {
            value: value,
            currency: currency,
            confidence: 0.8,
            source: `extracted from: ${match[0]}`
          };
        }
      }
    }
    
    // If no good match found, use default
    if (bestMatch.value === 0) {
      bestMatch.value = 50000000; // $50M default
    }
    
    return bestMatch;
  }

  // Extract transaction fee from text
  extractTransactionFee(text) {
    const feePatterns = [
      /(?:transaction fee|advisory fee|investment banking fee|closing costs|fees)[\s:]*([0-9]+(?:\.[0-9]+)?)\s*%/gi,
      /(?:fees?)[\s:]*([0-9]+(?:\.[0-9]+)?)\s*(?:percent|%)/gi,
      /([0-9]+(?:\.[0-9]+)?)\s*%\s*(?:fee|transaction|advisory)/gi
    ];
    
    for (const pattern of feePatterns) {
      const match = text.match(pattern);
      if (match) {
        const fee = parseFloat(match[1]);
        if (fee >= 0.5 && fee <= 10) { // Reasonable range for transaction fees
          return {
            value: fee,
            confidence: 0.8,
            source: `extracted from: ${match[0]}`
          };
        }
      }
    }
    
    return {
      value: 2.5, // Default 2.5%
      confidence: 0.3,
      source: 'default assumption (typical 2.5% for M&A transactions)'
    };
  }

  // Extract deal LTV from text
  extractDealLTV(text) {
    const ltvPatterns = [
      /(?:ltv|loan.to.value|debt.to.value|leverage ratio|debt ratio)[\s:]*([0-9]+(?:\.[0-9]+)?)\s*%/gi,
      /(?:debt financing|leverage)[\s:]*([0-9]+(?:\.[0-9]+)?)\s*(?:percent|%)/gi,
      /([0-9]+(?:\.[0-9]+)?)\s*%\s*(?:ltv|debt|leverage)/gi
    ];
    
    for (const pattern of ltvPatterns) {
      const match = text.match(pattern);
      if (match) {
        const ltv = parseFloat(match[1]);
        if (ltv >= 30 && ltv <= 90) { // Reasonable range for LTV
          return {
            value: ltv,
            confidence: 0.8,
            source: `extracted from: ${match[0]}`
          };
        }
      }
    }
    
    return {
      value: 70, // Default 70%
      confidence: 0.3,
      source: 'default assumption (typical 70% LTV for leveraged transactions)'
    };
  }

  // Identify currency from text/symbol
  identifyCurrency(currencyText) {
    if (!currencyText) return 'USD';
    
    const curr = currencyText.toUpperCase();
    if (curr.includes('USD') || curr.includes('$')) return 'USD';
    if (curr.includes('EUR') || curr.includes('€')) return 'EUR';
    if (curr.includes('GBP') || curr.includes('£')) return 'GBP';
    if (curr.includes('JPY') || curr.includes('¥')) return 'JPY';
    if (curr.includes('CAD')) return 'CAD';
    if (curr.includes('AUD')) return 'AUD';
    if (curr.includes('CHF')) return 'CHF';
    if (curr.includes('CNY') || curr.includes('RMB')) return 'CNY';
    
    return 'USD'; // Default
  }

  // Validate and clean extracted deal assumptions
  validateDealAssumptions(data) {
    return {
      dealName: this.validateDealName(data.dealName),
      dealValue: this.validateDealValue(data.dealValue),
      dealCurrency: this.validateCurrency(data.dealCurrency),
      transactionFee: this.validateTransactionFee(data.transactionFee),
      dealLTV: this.validateDealLTV(data.dealLTV),
      confidence: data.confidence || {},
      sources: data.sources || {}
    };
  }

  // Calculate derived values (equity and debt contributions)
  calculateDerivedValues(data) {
    const dealValue = data.dealValue || 0;
    const ltvPercent = data.dealLTV || 70;
    
    const debtFinancing = dealValue * (ltvPercent / 100);
    const equityContribution = dealValue - debtFinancing;
    
    return {
      ...data,
      equityContribution: equityContribution,
      debtFinancing: debtFinancing
    };
  }

  // Validation methods
  validateDealName(name) {
    if (!name || typeof name !== 'string' || name.trim().length < 2) {
      return 'Sample M&A Transaction';
    }
    return name.trim().substring(0, 100); // Limit length
  }

  validateDealValue(value) {
    const numValue = parseFloat(value);
    if (isNaN(numValue) || numValue <= 0) {
      return 50000000; // Default $50M
    }
    if (numValue < 1000000) return numValue * 1000000; // Convert to millions if needed
    if (numValue > 1000000000000) return 1000000000000; // Cap at 1T
    return numValue;
  }

  validateCurrency(currency) {
    const validCurrencies = ['USD', 'EUR', 'GBP', 'JPY', 'CAD', 'AUD', 'CHF', 'CNY', 'SEK', 'NOK'];
    if (validCurrencies.includes(currency)) {
      return currency;
    }
    return 'USD'; // Default
  }

  validateTransactionFee(fee) {
    const numFee = parseFloat(fee);
    if (isNaN(numFee) || numFee < 0 || numFee > 15) {
      return 2.5; // Default 2.5%
    }
    return numFee;
  }

  validateDealLTV(ltv) {
    const numLtv = parseFloat(ltv);
    if (isNaN(numLtv) || numLtv < 0 || numLtv > 100) {
      return 70; // Default 70%
    }
    return numLtv;
  }

  // Get fallback deal assumptions when extraction fails
  getFallbackDealAssumptions() {
    const dealValue = 50000000; // $50M
    const ltvPercent = 70;
    const debtFinancing = dealValue * (ltvPercent / 100);
    const equityContribution = dealValue - debtFinancing;
    
    return {
      dealName: 'Sample M&A Transaction',
      dealValue: dealValue,
      dealCurrency: 'USD',
      transactionFee: 2.5,
      dealLTV: ltvPercent,
      equityContribution: equityContribution,
      debtFinancing: debtFinancing,
      confidence: {
        dealName: 0.3,
        dealValue: 0.3,
        transactionFee: 0.3,
        dealLTV: 0.3
      },
      sources: {
        dealName: 'fallback default',
        dealValue: 'fallback default',
        transactionFee: 'fallback default',
        dealLTV: 'fallback default'
      }
    };
  }

  // Apply extracted deal assumptions to the form
  async applyDealAssumptions(dealData) {
    console.log('🎯 Applying deal assumptions to form...');
    console.log('🎯 Deal data:', dealData);
    
    try {
      // Apply basic deal parameters
      this.setInputValue('dealName', dealData.dealName);
      this.setInputValue('dealValue', dealData.dealValue);
      this.setInputValue('transactionFee', dealData.transactionFee);
      this.setInputValue('dealLTV', dealData.dealLTV);
      
      // Wait a moment for calculations to trigger
      await new Promise(resolve => setTimeout(resolve, 500));
      
      // The equity and debt contributions should be calculated automatically by the form
      // but we can verify they match our calculations
      const equityElement = document.getElementById('equityContribution');
      const debtElement = document.getElementById('debtFinancing');
      
      if (equityElement) {
        console.log('🎯 Equity contribution field found:', equityElement.value);
      }
      if (debtElement) {
        console.log('🎯 Debt financing field found:', debtElement.value);
      }
      
      // Trigger form calculations
      if (window.formHandler) {
        window.formHandler.triggerCalculations();
      }
      
      console.log('🎯 Deal assumptions applied successfully');
      return true;
      
    } catch (error) {
      console.error('Error applying deal assumptions:', error);
      return false;
    }
  }

  // Helper method to set input values
  setInputValue(elementId, value) {
    console.log(`🔧 Setting ${elementId} = ${value}`);
    const element = document.getElementById(elementId);
    if (element && value !== null && value !== undefined) {
      element.value = value;
      element.dispatchEvent(new Event('change', { bubbles: true }));
      element.dispatchEvent(new Event('input', { bubbles: true }));
      console.log(`🔧 Successfully set ${elementId}`);
    } else {
      console.log(`🔧 Failed to set ${elementId}: element=${!!element}, value=${value}`);
    }
  }

  // Get extraction summary for user feedback
  getExtractionSummary(dealData) {
    const confidenceText = (conf) => {
      if (conf >= 0.8) return 'High';
      if (conf >= 0.5) return 'Medium';
      return 'Low';
    };
    
    const formatCurrency = (value, currency) => {
      const formatter = new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: currency || 'USD',
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
      });
      return formatter.format(value);
    };
    
    return {
      title: 'Deal Assumptions Extracted',
      items: [
        {
          label: 'Deal Name',
          value: dealData.dealName,
          confidence: confidenceText(dealData.confidence?.dealName || 0),
          source: dealData.sources?.dealName || 'unknown'
        },
        {
          label: 'Deal Value',
          value: formatCurrency(dealData.dealValue, dealData.dealCurrency),
          confidence: confidenceText(dealData.confidence?.dealValue || 0),
          source: dealData.sources?.dealValue || 'unknown'
        },
        {
          label: 'Transaction Fee',
          value: `${dealData.transactionFee}%`,
          confidence: confidenceText(dealData.confidence?.transactionFee || 0),
          source: dealData.sources?.transactionFee || 'unknown'
        },
        {
          label: 'Deal LTV',
          value: `${dealData.dealLTV}%`,
          confidence: confidenceText(dealData.confidence?.dealLTV || 0),
          source: dealData.sources?.dealLTV || 'unknown'
        },
        {
          label: 'Equity Contribution',
          value: formatCurrency(dealData.equityContribution, dealData.dealCurrency),
          confidence: 'Calculated',
          source: 'calculated from Deal Value and LTV'
        },
        {
          label: 'Debt Financing',
          value: formatCurrency(dealData.debtFinancing, dealData.dealCurrency),
          confidence: 'Calculated',
          source: 'calculated from Deal Value and LTV'
        }
      ]
    };
  }
}

// Export for use in main application
window.DealAssumptionsExtractor = DealAssumptionsExtractor;