class MasterDataAnalyzer {
  constructor() {
    this.isInitialized = false;
    this.standardizedData = null;
  }

  initialize() {
    if (this.isInitialized) return;
    this.isInitialized = true;
    console.log('✅ MasterDataAnalyzer initialized');
  }

  // Stage 1: Create comprehensive analysis and standardized data table
  async analyzeAndStandardizeData(fileContents) {
    console.log('🔍 Starting master analysis of uploaded files...');
    
    try {
      // Call GPT-4 to create comprehensive standardized analysis
      const standardizedData = await this.callMasterAnalysisAI(fileContents);
      
      if (!standardizedData) {
        console.log('🔍 AI analysis failed, using fallback parsing...');
        const fallbackData = this.createFallbackStandardizedData(fileContents);
        this.standardizedData = fallbackData;
        return fallbackData;
      }
      
      // Store the standardized data for other extractors to use
      if (standardizedData) {
        this.standardizedData = standardizedData;
      }
      
      console.log('🔍 Master analysis completed:', standardizedData);
      return this.standardizedData;
      
    } catch (error) {
      console.error('Error in master analysis:', error);
      const fallbackData = this.createFallbackStandardizedData(fileContents);
      this.standardizedData = fallbackData;
      return fallbackData;
    }
  }

  // Call GPT-4 for comprehensive M&A analysis
  async callMasterAnalysisAI(fileContents) {
    try {
      console.log('🤖 Calling GPT-4 for master M&A analysis...');
      console.log('🤖 File contents to analyze:', fileContents.length, 'files');
      
      // Check if we're running locally or on Netlify
      const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
      const apiUrl = isLocal ? 'http://localhost:8888/.netlify/functions/chat' : '/.netlify/functions/chat';
      
      // Format file contents for the chat API
      const formattedContents = fileContents.map(file => 
        `=== FILE: ${file.name} ===\n${file.content}\n`
      ).join('\n');
      
      console.log('🤖 Sending request to:', apiUrl);
      
      const requestBody = {
        message: 'Analyze these M&A documents and create standardized data table.',
        fileContents: fileContents.map(f => `File: ${f.name}\n${f.content}`),
        autoFillMode: true,
        batchType: 'master_analysis'
      };
      
      console.log('🤖 Request body prepared, calling API...');
      
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(requestBody)
      });
      
      console.log('🤖 API response status:', response.status);
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error('🤖 API error response:', errorText);
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      
      const data = await response.json();
      console.log('🤖 API response data:', data);
      
      if (data.error) {
        console.error('🤖 API returned error:', data.error);
        throw new Error(data.error);
      }
      
      // Parse the standardized data from AI response
      if (data.extractedData && data.extractedData.standardizedData) {
        console.log('🤖 Successfully extracted standardized data');
        return data.extractedData.standardizedData;
      } else {
        console.log('🤖 No standardized data in response, using fallback');
        return null;
      }
      
    } catch (error) {
      console.error('Master analysis AI call failed:', error);
      return null;
    }
  }

  // Create comprehensive M&A analysis prompt
  createMasterAnalysisPrompt(formattedContents) {
    return `You are an expert M&A analyst. Your task is to analyze the provided documents and create a comprehensive, standardized data table that will be used by other AI systems to populate a financial model.

CRITICAL TASK: Create a standardized analysis table that captures ALL relevant information for M&A financial modeling, including P&L, FCF, and IRR calculations.

DOCUMENT CONTENT TO ANALYZE:
${formattedContents}

INSTRUCTIONS:
1. Act as a senior M&A analyst reviewing deal documents
2. Extract and organize ALL financial information systematically
3. Create standardized categories for easy retrieval
4. Include confidence levels and sources for each data point
5. Fill gaps with industry-standard assumptions where needed

REQUIRED STANDARDIZED DATA FORMAT (JSON):
{
  "extractedData": {
    "standardizedData": {
      "companyOverview": {
        "companyName": "string",
        "industry": "string", 
        "businessDescription": "string",
        "keyBusinessMetrics": "string"
      },
      "transactionDetails": {
        "dealName": "string",
        "dealValue": number,
        "currency": "string",
        "transactionType": "string",
        "transactionFees": number,
        "closingDate": "YYYY-MM-DD",
        "expectedExitDate": "YYYY-MM-DD"
      },
      "financingStructure": {
        "totalDealValue": number,
        "debtLTV": number,
        "equityContribution": number,
        "debtFinancing": number,
        "interestRate": number,
        "loanTerms": "string"
      },
      "historicalFinancials": {
        "baseYear": "string",
        "revenueStreams": [
          {
            "name": "string",
            "currentValue": number,
            "growthRate": number,
            "growthType": "linear|compound"
          }
        ],
        "operatingExpenses": [
          {
            "name": "string", 
            "currentValue": number,
            "inflationRate": number,
            "category": "staff|marketing|rent|other"
          }
        ],
        "capitalExpenses": [
          {
            "name": "string",
            "currentValue": number,
            "frequency": "one-time|annual|periodic"
          }
        ]
      },
      "projectionAssumptions": {
        "projectionPeriod": "string",
        "reportingFrequency": "monthly|quarterly|annually",
        "keyGrowthDrivers": "string",
        "marketAssumptions": "string",
        "riskFactors": "string"
      },
      "exitAssumptions": {
        "exitStrategy": "string",
        "exitMultiple": number,
        "terminalValue": number,
        "disposalCosts": number,
        "expectedIRR": number
      },
      "keyMetrics": {
        "currentEBITDA": number,
        "EBITDAMargin": number,
        "currentRevenue": number,
        "revenueGrowthRate": number,
        "paybackPeriod": number
      },
      "dataQuality": {
        "overallConfidence": number,
        "missingCriticalData": ["string"],
        "assumptions": ["string"],
        "dataSourceQuality": "high|medium|low"
      }
    }
  }
}

ANALYSIS GUIDELINES:
- Look for numerical values, percentages, dates, and financial figures
- Identify revenue streams, cost categories, and growth patterns
- Extract deal structure, financing terms, and exit plans
- Note any missing data and suggest reasonable industry assumptions
- Organize data logically for financial modeling purposes
- Include context and explanations for complex items

RETURN ONLY THE JSON STRUCTURE - NO OTHER TEXT.`;
  }

  // Create fallback standardized data when AI fails
  createFallbackStandardizedData(fileContents) {
    console.log('🔍 Creating fallback standardized data...');
    
    // Basic parsing of file contents to extract what we can
    const allText = fileContents.map(f => f.content).join(' ').toLowerCase();
    
    // Try to extract basic information
    const companyName = this.extractCompanyName(fileContents);
    const dealValue = this.extractDealValue(allText);
    const currency = this.extractCurrency(allText);
    
    return {
      companyOverview: {
        companyName: companyName,
        industry: "Technology", // Default assumption
        businessDescription: "Business details to be confirmed",
        keyBusinessMetrics: "Revenue and growth metrics to be analyzed"
      },
      transactionDetails: {
        dealName: companyName + " Acquisition",
        dealValue: dealValue,
        currency: currency,
        transactionType: "Acquisition",
        transactionFees: 2.5,
        closingDate: new Date().toISOString().split('T')[0],
        expectedExitDate: new Date(Date.now() + 5*365*24*60*60*1000).toISOString().split('T')[0]
      },
      financingStructure: {
        totalDealValue: dealValue,
        debtLTV: 70,
        equityContribution: dealValue * 0.3,
        debtFinancing: dealValue * 0.7,
        interestRate: 5.5,
        loanTerms: "5-year term loan"
      },
      historicalFinancials: {
        baseYear: new Date().getFullYear().toString(),
        revenueStreams: [
          {
            name: "Primary Revenue",
            currentValue: dealValue * 0.1, // Estimate 10% of deal value as annual revenue
            growthRate: 15,
            growthType: "compound"
          }
        ],
        operatingExpenses: [
          {
            name: "Staff Costs",
            currentValue: dealValue * 0.05, // Estimate 5% of deal value
            inflationRate: 3,
            category: "staff"
          }
        ],
        capitalExpenses: [
          {
            name: "Equipment & Technology",
            currentValue: dealValue * 0.02, // Estimate 2% of deal value
            frequency: "annual"
          }
        ]
      },
      projectionAssumptions: {
        projectionPeriod: "5 years",
        reportingFrequency: "monthly",
        keyGrowthDrivers: "Market expansion and operational improvements",
        marketAssumptions: "Stable market conditions",
        riskFactors: "Competition and regulatory changes"
      },
      exitAssumptions: {
        exitStrategy: "Trade sale or IPO",
        exitMultiple: 12,
        terminalValue: dealValue * 2, // Estimate 2x return
        disposalCosts: 2.5,
        expectedIRR: 20
      },
      keyMetrics: {
        currentEBITDA: dealValue * 0.03, // Estimate 3% of deal value
        EBITDAMargin: 25,
        currentRevenue: dealValue * 0.1,
        revenueGrowthRate: 15,
        paybackPeriod: 4
      },
      dataQuality: {
        overallConfidence: 0.3,
        missingCriticalData: ["Historical financials", "Detailed revenue breakdown", "Operating expense details"],
        assumptions: ["Used industry benchmarks", "Estimated based on deal size"],
        dataSourceQuality: "low"
      }
    };
  }

  // Helper methods for fallback parsing
  extractCompanyName(fileContents) {
    for (const file of fileContents) {
      const filename = file.name.replace(/\.(csv|pdf|png|jpg|jpeg)$/i, '');
      if (filename.length > 3 && !filename.toLowerCase().includes('data') && !filename.toLowerCase().includes('test')) {
        return filename.replace(/[-_]/g, ' ');
      }
    }
    return 'Target Company';
  }

  extractDealValue(text) {
    const valuePatterns = [
      /deal value[^0-9]*([0-9,]+(?:\.[0-9]+)?)\s*(?:million|billion|m|b)?/gi,
      /\$\s*([0-9,]+(?:\.[0-9]+)?)\s*(?:million|billion|m|b)?/gi
    ];
    
    for (const pattern of valuePatterns) {
      const match = text.match(pattern);
      if (match) {
        let value = parseFloat(match[1].replace(/,/g, ''));
        if (text.includes('billion') || text.includes(' b')) {
          value *= 1000000000;
        } else if (text.includes('million') || text.includes(' m')) {
          value *= 1000000;
        }
        if (value >= 1000000) return value;
      }
    }
    return 50000000; // Default $50M
  }

  extractCurrency(text) {
    if (text.includes('€') || text.includes('eur')) return 'EUR';
    if (text.includes('£') || text.includes('gbp')) return 'GBP';
    if (text.includes('¥') || text.includes('jpy')) return 'JPY';
    return 'USD';
  }

  // Get the standardized data for other extractors to use
  getStandardizedData() {
    return this.standardizedData;
  }

  // Check if standardized data is available
  hasStandardizedData() {
    return this.standardizedData !== null;
  }

  // Clear standardized data (for new file uploads)
  clearStandardizedData() {
    this.standardizedData = null;
  }

  // Get specific section of standardized data
  getSection(sectionName) {
    if (!this.standardizedData) return null;
    return this.standardizedData[sectionName] || null;
  }

  // Get data quality information
  getDataQuality() {
    if (!this.standardizedData || !this.standardizedData.dataQuality) {
      return {
        overallConfidence: 0.3,
        dataSourceQuality: 'low'
      };
    }
    return this.standardizedData.dataQuality;
  }

  // Create summary for user feedback
  getAnalysisSummary() {
    if (!this.standardizedData) return null;
    
    const quality = this.getDataQuality();
    const company = this.standardizedData.companyOverview || {};
    const transaction = this.standardizedData.transactionDetails || {};
    
    return {
      title: 'Master Analysis Complete',
      summary: {
        company: company.companyName || 'Unknown Company',
        dealValue: transaction.dealValue || 0,
        currency: transaction.currency || 'USD',
        confidence: Math.round(quality.overallConfidence * 100) + '%',
        dataQuality: quality.dataSourceQuality || 'low'
      },
      sections: [
        'Company Overview',
        'Transaction Details', 
        'Financing Structure',
        'Historical Financials',
        'Projection Assumptions',
        'Exit Assumptions',
        'Key Metrics'
      ]
    };
  }
}

// Export for use in main application
window.MasterDataAnalyzer = MasterDataAnalyzer;