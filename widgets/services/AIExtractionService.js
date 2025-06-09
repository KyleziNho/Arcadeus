/**
 * AIExtractionService.js - Sophisticated AI extraction service with OpenAI GPT-4
 * Handles document analysis, confidence scoring, and intelligent data extraction
 */

class AIExtractionService {
  constructor() {
    this.apiEndpoint = this.getAPIEndpoint();
    this.maxRetries = 3;
    this.retryDelay = 1000;
    this.confidenceThreshold = 0.5;
    this.extractionCache = new Map();
  }

  getAPIEndpoint() {
    const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
    return isLocal ? 'http://localhost:8888/.netlify/functions/chat' : '/.netlify/functions/chat';
  }

  /**
   * Main extraction method - analyzes documents and returns structured data
   */
  async extractFromDocuments(files, extractionType = 'comprehensive') {
    console.log(`🤖 AI Extraction Service: Processing ${files.length} files for ${extractionType} extraction`);
    
    // Check cache first
    const cacheKey = this.generateCacheKey(files, extractionType);
    if (this.extractionCache.has(cacheKey)) {
      console.log('🤖 Returning cached extraction results');
      return this.extractionCache.get(cacheKey);
    }

    try {
      // Prepare file contents
      const fileContents = this.prepareFileContents(files);
      
      // Select appropriate extraction strategy
      const extractionStrategy = this.getExtractionStrategy(extractionType);
      
      // Call AI with retry logic
      const response = await this.callAIWithRetry(fileContents, extractionStrategy);
      
      // Process and validate response
      const extractedData = this.processAIResponse(response, extractionType);
      
      // Calculate confidence scores
      const dataWithConfidence = this.calculateConfidenceScores(extractedData, fileContents);
      
      // Cache successful extraction
      this.extractionCache.set(cacheKey, dataWithConfidence);
      
      return dataWithConfidence;
      
    } catch (error) {
      console.error('🤖 AI Extraction failed:', error);
      
      // Check if it's an API configuration issue
      if (error.message.includes('500') || error.message.includes('API error')) {
        throw new Error('AI service is currently unavailable. Please check API configuration or try again later.');
      } else {
        throw new Error(`AI extraction failed: ${error.message}`);
      }
    }
  }

  /**
   * Prepare file contents for AI processing
   */
  prepareFileContents(files) {
    return files.map(file => {
      // Add metadata to help AI understand context
      const metadata = {
        filename: file.name,
        type: file.type,
        size: file.size,
        processor: file.processor
      };
      
      return {
        metadata: metadata,
        content: file.content || '',
        extractedText: file.extractedData || null
      };
    });
  }

  /**
   * Get extraction strategy based on type
   */
  getExtractionStrategy(type) {
    const strategies = {
      comprehensive: {
        model: 'gpt-4-turbo-preview',
        temperature: 0.1,
        maxTokens: 4000,
        systemPrompt: this.getComprehensiveExtractionPrompt()
      },
      dealAssumptions: {
        model: 'gpt-4-turbo-preview',
        temperature: 0.1,
        maxTokens: 1500,
        systemPrompt: this.getDealAssumptionsPrompt()
      },
      revenue: {
        model: 'gpt-4-turbo-preview',
        temperature: 0.1,
        maxTokens: 2000,
        systemPrompt: this.getRevenueExtractionPrompt()
      },
      costs: {
        model: 'gpt-4-turbo-preview',
        temperature: 0.1,
        maxTokens: 2000,
        systemPrompt: this.getCostExtractionPrompt()
      },
      debtModel: {
        model: 'gpt-4-turbo-preview',
        temperature: 0.1,
        maxTokens: 1000,
        systemPrompt: this.getDebtModelPrompt()
      },
      exitAssumptions: {
        model: 'gpt-4-turbo-preview',
        temperature: 0.1,
        maxTokens: 1000,
        systemPrompt: this.getExitAssumptionsPrompt()
      }
    };
    
    return strategies[type] || strategies.comprehensive;
  }

  /**
   * Call AI API with retry logic
   */
  async callAIWithRetry(fileContents, strategy, retryCount = 0) {
    try {
      const requestBody = {
        message: 'Extract financial data from these documents',
        fileContents: fileContents.map(f => `
          File: ${f.metadata.filename}
          Type: ${f.metadata.type}
          Content: ${f.content}
        `),
        systemPrompt: strategy.systemPrompt,
        temperature: strategy.temperature,
        maxTokens: strategy.maxTokens,
        autoFillMode: true
      };

      console.log('🤖 Sending request to AI API...');
      
      const response = await fetch(this.apiEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        body: JSON.stringify(requestBody)
      });

      if (!response.ok) {
        throw new Error(`API error: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      
      if (data.error) {
        throw new Error(data.error);
      }

      return data;
      
    } catch (error) {
      console.error(`🤖 AI API call failed (attempt ${retryCount + 1}):`, error);
      
      if (retryCount < this.maxRetries) {
        console.log(`🤖 Retrying in ${this.retryDelay}ms...`);
        await this.sleep(this.retryDelay * (retryCount + 1));
        return this.callAIWithRetry(fileContents, strategy, retryCount + 1);
      }
      
      throw error;
    }
  }

  /**
   * Process AI response and extract structured data
   */
  processAIResponse(response, extractionType) {
    console.log('🤖 Processing AI response...');
    
    try {
      // Handle different response formats
      let extractedData;
      
      if (response.extractedData) {
        extractedData = response.extractedData;
      } else if (response.data) {
        extractedData = response.data;
      } else if (typeof response === 'string') {
        // Try to parse JSON from string response
        extractedData = JSON.parse(response);
      } else {
        extractedData = response;
      }

      // Validate extracted data structure
      this.validateExtractedData(extractedData, extractionType);
      
      return extractedData;
      
    } catch (error) {
      console.error('🤖 Error processing AI response:', error);
      throw new Error('Failed to process AI response: ' + error.message);
    }
  }

  /**
   * Calculate confidence scores for extracted data
   */
  calculateConfidenceScores(data, fileContents) {
    console.log('🤖 Calculating confidence scores...');
    
    const scoredData = {};
    
    for (const [field, value] of Object.entries(data)) {
      if (value === null || value === undefined) {
        scoredData[field] = {
          value: null,
          confidence: 0,
          source: 'not_found'
        };
        continue;
      }

      // Calculate confidence based on multiple factors
      let confidence = 0.5; // Base confidence
      
      // Factor 1: Value found in multiple files
      const occurrences = this.countValueOccurrences(value, fileContents);
      if (occurrences > 1) confidence += 0.2;
      
      // Factor 2: Value matches expected format
      if (this.validateFieldFormat(field, value)) confidence += 0.2;
      
      // Factor 3: Reasonable value range
      if (this.isReasonableValue(field, value)) confidence += 0.1;
      
      // Find source document
      const source = this.findValueSource(value, fileContents);
      
      scoredData[field] = {
        value: value,
        confidence: Math.min(confidence, 1.0),
        source: source
      };
    }
    
    return scoredData;
  }

  /**
   * Validate extracted data structure
   */
  validateExtractedData(data, type) {
    const requiredFields = {
      comprehensive: ['dealName', 'dealValue', 'currency'],
      dealAssumptions: ['dealName', 'dealValue', 'transactionFee', 'dealLTV'],
      revenue: ['revenueItems'],
      costs: ['operatingExpenses', 'capitalExpenses'],
      debtModel: ['loanIssuanceFees', 'interestRate'],
      exitAssumptions: ['disposalCost', 'terminalCapRate']
    };
    
    const required = requiredFields[type] || [];
    const missing = required.filter(field => !data.hasOwnProperty(field));
    
    if (missing.length > 0) {
      console.warn(`🤖 Warning: Missing required fields: ${missing.join(', ')}`);
    }
  }

  /**
   * Helper methods
   */
  
  countValueOccurrences(value, fileContents) {
    if (!value) return 0;
    const valueStr = String(value).toLowerCase();
    return fileContents.filter(f => 
      f.content && f.content.toLowerCase().includes(valueStr)
    ).length;
  }

  validateFieldFormat(field, value) {
    const formatValidators = {
      dealValue: (v) => typeof v === 'number' && v > 0,
      transactionFee: (v) => typeof v === 'number' && v >= 0 && v <= 100,
      dealLTV: (v) => typeof v === 'number' && v >= 0 && v <= 100,
      currency: (v) => ['USD', 'EUR', 'GBP', 'JPY', 'CAD', 'AUD', 'CHF', 'CNY'].includes(v),
      projectStartDate: (v) => !isNaN(Date.parse(v)),
      projectEndDate: (v) => !isNaN(Date.parse(v))
    };
    
    const validator = formatValidators[field];
    return validator ? validator(value) : true;
  }

  isReasonableValue(field, value) {
    const reasonableRanges = {
      dealValue: (v) => v >= 1000000 && v <= 100000000000, // $1M to $100B
      transactionFee: (v) => v >= 0.5 && v <= 5, // 0.5% to 5%
      dealLTV: (v) => v >= 30 && v <= 90, // 30% to 90%
      disposalCost: (v) => v >= 0.5 && v <= 5, // 0.5% to 5%
      terminalCapRate: (v) => v >= 4 && v <= 15 // 4% to 15%
    };
    
    const validator = reasonableRanges[field];
    return validator ? validator(value) : true;
  }

  findValueSource(value, fileContents) {
    if (!value) return 'not_found';
    
    const valueStr = String(value).toLowerCase();
    for (const file of fileContents) {
      if (file.content && file.content.toLowerCase().includes(valueStr)) {
        return file.metadata.filename;
      }
    }
    
    return 'inferred';
  }

  generateCacheKey(files, type) {
    const fileNames = files.map(f => f.name).sort().join('|');
    return `${type}:${fileNames}`;
  }

  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Extraction Prompts
   */
  
  getComprehensiveExtractionPrompt() {
    return `You are an expert M&A analyst extracting data from financial documents.

TASK: Extract ALL relevant financial information for M&A modeling.

CRITICAL INSTRUCTIONS:
1. Extract ONLY real data found in the documents
2. Do NOT make up or estimate values
3. Return null for any field not found
4. Include all revenue items, cost items, and financial metrics
5. Preserve exact values and units as found

REQUIRED OUTPUT FORMAT:
{
  "dealName": "actual company/deal name or null",
  "dealValue": numeric_value_or_null,
  "currency": "USD|EUR|GBP|etc or null",
  "transactionFee": percentage_as_number_or_null,
  "dealLTV": percentage_as_number_or_null,
  "projectStartDate": "YYYY-MM-DD or null",
  "projectEndDate": "YYYY-MM-DD or null",
  "modelPeriods": "daily|monthly|quarterly|yearly or null",
  "revenueItems": [
    {
      "name": "revenue stream name",
      "value": numeric_value,
      "growthType": "linear|compound|custom",
      "growthRate": percentage_as_number
    }
  ],
  "operatingExpenses": [
    {
      "name": "expense name",
      "value": numeric_value,
      "growthType": "linear|compound|custom",
      "growthRate": percentage_as_number
    }
  ],
  "capitalExpenses": [
    {
      "name": "capex name",
      "value": numeric_value,
      "growthType": "linear|compound|custom",
      "growthRate": percentage_as_number
    }
  ],
  "debtFinancing": {
    "loanIssuanceFees": percentage_as_number_or_null,
    "interestRateType": "fixed|floating or null",
    "interestRate": percentage_as_number_or_null,
    "baseRate": percentage_as_number_or_null,
    "creditMargin": percentage_as_number_or_null
  },
  "exitAssumptions": {
    "disposalCost": percentage_as_number_or_null,
    "terminalCapRate": percentage_as_number_or_null
  }
}

IMPORTANT: Return ONLY valid JSON. Extract actual values from documents.`;
  }

  getDealAssumptionsPrompt() {
    return `Extract deal assumptions from M&A documents.

Look for:
- Deal/Transaction/Company name
- Deal/Enterprise/Transaction value (in millions/billions)
- Transaction/Advisory fees (as %)
- LTV/Leverage ratio (as %)
- Equity and debt split

Return JSON with actual extracted values or null.`;
  }

  getRevenueExtractionPrompt() {
    return `Extract all revenue streams and projections.

Look for:
- Revenue line items
- Sales categories
- Income sources
- Growth rates or projections
- Historical revenues

Format each as: name, current value, growth type, growth rate.`;
  }

  getCostExtractionPrompt() {
    return `Extract all operating and capital expenses.

Look for:
- Operating expenses (staff, rent, marketing, etc.)
- Capital expenditures
- Cost inflation rates
- Expense projections

Separate into operatingExpenses and capitalExpenses arrays.`;
  }

  getDebtModelPrompt() {
    return `Extract debt financing details.

Look for:
- Loan issuance/arrangement fees
- Interest rates (fixed or floating)
- Base rates (LIBOR, SOFR, etc.)
- Credit spreads/margins
- Loan terms

Return all debt-related parameters.`;
  }

  getExitAssumptionsPrompt() {
    return `Extract exit and valuation assumptions.

Look for:
- Disposal/Exit costs (as %)
- Terminal cap rates
- Exit multiples
- Terminal values
- Expected returns

Focus on percentage values and multiples.`;
  }
}

// Export for use
window.AIExtractionService = AIExtractionService;