class MasterDataAnalyzer {
  constructor() {
    this.isInitialized = false;
    this._data = null;
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
      let result = await this.callMasterAnalysisAI(fileContents);
      
      if (!result) {
        console.error('🔍 CRITICAL: AI analysis failed! This indicates a problem with API calls or parsing.');
        console.error('🔍 File contents available:', fileContents ? fileContents.length : 0);
        if (fileContents && fileContents.length > 0) {
          console.error('🔍 Files that should be analyzed:', fileContents.map(f => `${f.name} (${f.content ? f.content.length : 0} chars)`));
        }
        console.log('🔍 Using fallback parsing...');
        result = this.createFallbackStandardizedData(fileContents);
      }
      
      // Store the result
      this._data = result;
      
      console.log('🔍 Master analysis completed:', result);
      return result;
      
    } catch (error) {
      console.error('Error in master analysis:', error);
      const fallbackData = this.createFallbackStandardizedData(fileContents);
      this._data = fallbackData;
      return fallbackData;
    }
  }

  // Call GPT-4 for comprehensive M&A analysis
  async callMasterAnalysisAI(fileContents) {
    try {
      console.log('🤖 Calling GPT-4 for master M&A analysis...');
      console.log('🤖 File contents to analyze:', fileContents.length, 'files');
      
      // DEBUG: Log actual file content being sent to AI
      fileContents.forEach((file, index) => {
        console.log(`🤖 DEBUG - File ${index + 1}: ${file.name}`);
        console.log(`🤖 DEBUG - Content preview: ${file.content ? file.content.substring(0, 500) : 'NO CONTENT'}`);
        console.log(`🤖 DEBUG - Content length: ${file.content ? file.content.length : 0} characters`);
      });
      
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
      
      console.log('🤖 DEBUG: Request body structure:', {
        messageLength: requestBody.message.length,
        fileContentsCount: requestBody.fileContents.length,
        autoFillMode: requestBody.autoFillMode,
        batchType: requestBody.batchType,
        firstFilePreview: requestBody.fileContents[0] ? requestBody.fileContents[0].substring(0, 200) : 'No files'
      });
      
      console.log('🤖 Request body prepared, calling API...');
      
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        body: JSON.stringify(requestBody)
      });
      
      console.log('🤖 API response status:', response.status);
      console.log('🤖 API response headers:', response.headers);
      
      if (!response.ok) {
        let errorText;
        try {
          errorText = await response.text();
        } catch (e) {
          errorText = 'Could not read error response';
        }
        console.error('🤖 API error response:', errorText);
        console.error('🤖 Request URL:', apiUrl);
        console.error('🤖 Request body preview:', JSON.stringify(requestBody).substring(0, 1000));
        throw new Error(`HTTP error! status: ${response.status} - ${errorText}`);
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
      } else if (data.extractedData) {
        console.log('🤖 Got extractedData but no standardizedData, checking structure:', data.extractedData);
        // Sometimes the AI might return the data directly under extractedData
        return data.extractedData;
      } else {
        console.log('🤖 CRITICAL: AI extraction failed - no extractedData in response');
        console.log('🤖 Full response data:', data);
        alert('AI extraction failed. Check console for details. Using fallback parsing...');
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
    console.log('🔍 File contents length:', fileContents ? fileContents.length : 0);
    
    // CRITICAL: Check if we have actual file content to extract from
    if (!fileContents || fileContents.length === 0) {
      console.warn('🔍 No file contents provided - using minimal defaults');
    } else {
      console.log('🔍 ALERT: AI extraction failed but we have file content! This should be investigated.');
      console.log('🔍 Files available for extraction:', fileContents.map(f => f.name));
    }
    
    // DEBUG: Show what files we received
    if (fileContents && fileContents.length > 0) {
      fileContents.forEach((file, index) => {
        console.log(`🔍 File ${index + 1}: ${file.name} (${file.content ? file.content.length : 0} chars)`);
        if (file.content) {
          console.log(`🔍 Content preview: ${file.content.substring(0, 200)}...`);
        }
      });
    }
    
    // Basic parsing of file contents to extract what we can
    let allText = '';
    let companyName = null;
    let dealValue = null;
    let currency = null;
    let transactionFee = null;
    let ltvRatio = null;
    let startDate = null;
    let endDate = null;
    let reportingPeriod = null;
    
    try {
      if (fileContents && fileContents.length > 0) {
        allText = fileContents.map(f => f.content || '').join(' ');
        companyName = this.extractCompanyName(fileContents);
        dealValue = this.extractDealValue(allText);
        currency = this.extractCurrency(allText);
        
        // Extract all values and show comprehensive summary
        transactionFee = this.extractTransactionFee(allText);
        ltvRatio = this.extractLTVRatio(allText);
        startDate = this.extractStartDate(allText);
        endDate = this.extractEndDate(allText);
        reportingPeriod = this.extractReportingPeriod(allText);
        
        const extractionSummary = {
          companyName,
          dealValue,
          currency,
          transactionFee,
          ltvRatio,
          startDate,
          endDate,
          reportingPeriod
        };
        
        console.log('📊 ========== FULL EXTRACTION SUMMARY ==========');
        console.log('📊 Company Name:', companyName || '(blank)');
        console.log('📊 Deal Value:', dealValue || '(blank)');
        console.log('📊 Currency:', currency || '(blank)');
        console.log('📊 Transaction Fee:', transactionFee || '(blank)');
        console.log('📊 LTV Ratio:', ltvRatio || '(blank)');
        console.log('📊 Start Date:', startDate || '(blank)');
        console.log('📊 End Date:', endDate || '(blank)');
        console.log('📊 Reporting Period:', reportingPeriod || '(blank)');
        console.log('📊 ============================================');
        
        // Show this summary to user via UI
        this.showExtractionSummaryToUser(extractionSummary);
      }
    } catch (extractError) {
      console.warn('🔍 Error in basic extraction:', extractError);
    }
    
    const fallbackData = {
      companyOverview: {
        companyName: companyName || 'Unknown Company',
        industry: 'Unknown',
        businessDescription: 'Extracted from uploaded files',
        keyBusinessMetrics: 'Not available'
      },
      transactionDetails: {
        dealName: companyName ? companyName + ' Acquisition' : 'Unnamed Deal',
        dealValue: dealValue,
        currency: currency,
        transactionType: 'Acquisition',
        transactionFees: transactionFee,
        closingDate: startDate,
        expectedExitDate: endDate
      },
      financingStructure: {
        totalDealValue: dealValue,
        debtLTV: ltvRatio,
        equityContribution: dealValue && ltvRatio ? dealValue * (1 - ltvRatio / 100) : null,
        debtFinancing: dealValue && ltvRatio ? dealValue * (ltvRatio / 100) : null,
        interestRate: 5.5,
        loanTerms: "5-year term loan"
      },
      historicalFinancials: {
        baseYear: new Date().getFullYear().toString(),
        revenueStreams: [
          {
            name: "Primary Revenue",
            currentValue: dealValue * 0.1, // Estimate 10% of deal value as annual revenue
            growthRate: 0, // No fallback - must be provided by user
            growthType: "linear"
          }
        ],
        operatingExpenses: [
          {
            name: "Staff Costs",
            currentValue: dealValue * 0.05, // Estimate 5% of deal value
            inflationRate: 0, // No fallback - must be provided by user
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
    
    console.log('🔍 Fallback data created:', fallbackData);
    return fallbackData;
  }

  // Helper methods for fallback parsing
  extractCompanyName(fileContents) {
    try {
      if (!fileContents || fileContents.length === 0) return 'Target Company';
      
      // First try to extract from CSV content
      for (const file of fileContents) {
        if (!file || !file.content) continue;
        
        // Look for company name in CSV content
        const lines = file.content.split('\n');
        for (const line of lines) {
          // Handle different CSV formats: "Label,Value" or "Label,,,,,Value"
          if (line.toLowerCase().includes('company') || line.toLowerCase().includes('assumptions')) {
            // Try to extract company name from header line like "Sample Company Ltd. - Key Assumptions"
            if (line.includes(' - ') || line.includes('Ltd') || line.includes('Inc') || line.includes('Corp')) {
              const parts = line.split(',');
              const companyName = parts[0].trim().replace(' - Key Assumptions', '').replace(' - Assumptions', '');
              if (companyName.length > 2 && !companyName.toLowerCase().includes('sample')) {
                console.log(`📊 Found company name in CSV header: ${companyName}`);
                return companyName;
              }
            }
          }
        }
      }
      
      // Fallback to filename
      for (const file of fileContents) {
        if (!file || !file.name) continue;
        const filename = file.name.replace(/\.(csv|pdf|png|jpg|jpeg)$/i, '');
        if (filename.length > 3 && !filename.toLowerCase().includes('data') && !filename.toLowerCase().includes('test')) {
          return filename.replace(/[-_]/g, ' ');
        }
      }
    } catch (error) {
      console.warn('Error extracting company name:', error);
    }
    return 'Target Company';
  }

  extractDealValue(text) {
    try {
      if (!text || typeof text !== 'string') return 50000000;
      
      console.log('📊 Extracting deal value from text length:', text.length);
      console.log('📊 Text preview:', text.substring(0, 500));
      
      // Look for both Equity Contribution and Debt Financing to calculate total deal value
      const lines = text.split('\n');
      let equityContribution = 0;
      let debtFinancing = 0;
      
      for (const line of lines) {
        if (line.toLowerCase().includes('equity contribution')) {
          const parts = line.split(',');
          // Find the last non-empty part (handles multiple commas)
          for (let i = parts.length - 1; i >= 0; i--) {
            if (parts[i].trim()) {
              let valueStr = parts[i].trim().replace(/["£€¥\$,]/g, ''); // Remove quotes, currency symbols, commas
              const value = parseFloat(valueStr);
              if (!isNaN(value) && value > 0) {
                equityContribution = value;
                console.log(`📊 Found equity contribution in CSV: ${value}`);
              }
              break;
            }
          }
        }
        
        if (line.toLowerCase().includes('debt financing')) {
          const parts = line.split(',');
          // Find the last non-empty part (handles multiple commas)
          for (let i = parts.length - 1; i >= 0; i--) {
            if (parts[i].trim()) {
              let valueStr = parts[i].trim().replace(/["£€¥\$,]/g, ''); // Remove quotes, currency symbols, commas
              const value = parseFloat(valueStr);
              if (!isNaN(value) && value > 0) {
                debtFinancing = value;
                console.log(`📊 Found debt financing in CSV: ${value}`);
              }
              break;
            }
          }
        }
      }
      
      // Calculate total deal value
      if (equityContribution > 0 && debtFinancing > 0) {
        const totalDealValue = equityContribution + debtFinancing;
        console.log(`📊 Calculated total deal value: ${totalDealValue} (Equity: ${equityContribution} + Debt: ${debtFinancing})`);
        return totalDealValue;
      } else if (equityContribution > 0) {
        console.log(`📊 Using equity contribution as deal value: ${equityContribution}`);
        return equityContribution;
      } else if (debtFinancing > 0) {
        console.log(`📊 Using debt financing as deal value: ${debtFinancing}`);
        return debtFinancing;
      }
      
      // Fallback to regex patterns
      const valuePatterns = [
        /deal value[^0-9]*([0-9,]+(?:\.[0-9]+)?)\s*(?:million|billion|m|b)?/gi,
        /\$\s*([0-9,]+(?:\.[0-9]+)?)\s*(?:million|billion|m|b)?/gi
      ];
      
      for (const pattern of valuePatterns) {
        const match = text.match(pattern);
        if (match && match[1]) {
          let value = parseFloat(match[1].replace(/,/g, ''));
          if (!isNaN(value)) {
            if (text.includes('billion') || text.includes(' b')) {
              value *= 1000000000;
            } else if (text.includes('million') || text.includes(' m')) {
              value *= 1000000;
            }
            if (value >= 1000000) {
              console.log(`📊 Found deal value via regex: ${value}`);
              return value;
            }
          }
        }
      }
    } catch (error) {
      console.warn('Error extracting deal value:', error);
    }
    console.log('📊 No deal value found in CSV!');
    return null; // No default - force user to see what's wrong
  }

  extractCurrency(text) {
    try {
      if (!text || typeof text !== 'string') return 'USD';
      
      console.log('📊 Extracting currency from text length:', text.length);
      console.log('📊 Text preview for currency:', text.substring(0, 200));
      
      // Look for CSV format: "Currency,USD" or "Currency·,,,,,£"
      const lines = text.split('\n');
      for (const line of lines) {
        if (line.toLowerCase().includes('currency')) {
          const parts = line.split(',');
          // Find the last non-empty part
          for (let i = parts.length - 1; i >= 0; i--) {
            if (parts[i].trim()) {
              const currencyText = parts[i].trim();
              if (currencyText.includes('£') || currencyText.toLowerCase().includes('gbp')) {
                console.log(`📊 Found currency in CSV: GBP`);
                return 'GBP';
              }
              if (currencyText.includes('€') || currencyText.toLowerCase().includes('eur')) {
                console.log(`📊 Found currency in CSV: EUR`);
                return 'EUR';
              }
              if (currencyText.includes('$') || currencyText.toLowerCase().includes('usd')) {
                console.log(`📊 Found currency in CSV: USD`);
                return 'USD';
              }
              break;
            }
          }
        }
      }
      
      // Fallback to text patterns
      const lowerText = text.toLowerCase();
      if (lowerText.includes('€') || lowerText.includes('eur')) return 'EUR';
      if (lowerText.includes('£') || lowerText.includes('gbp')) return 'GBP';
      if (lowerText.includes('¥') || lowerText.includes('jpy')) return 'JPY';
    } catch (error) {
      console.warn('Error extracting currency:', error);
    }
    console.log('📊 No currency found in CSV!');
    return null; // No default
  }

  // Extract transaction fee from CSV
  extractTransactionFee(text) {
    try {
      if (!text || typeof text !== 'string') return 2.5;
      
      console.log('📊 Extracting transaction fee from text');
      
      // Look for CSV format: "Transaction Fees,,,,,1.50%"
      const lines = text.split('\n');
      for (const line of lines) {
        if (line.toLowerCase().includes('transaction fee') || (line.toLowerCase().includes('fee') && !line.toLowerCase().includes('issuance'))) {
          const parts = line.split(',');
          // Find the last non-empty part
          for (let i = parts.length - 1; i >= 0; i--) {
            if (parts[i].trim()) {
              const feeStr = parts[i].trim().replace(/[^0-9.]/g, '');
              const fee = parseFloat(feeStr);
              if (!isNaN(fee) && fee >= 0 && fee <= 10) {
                console.log(`📊 Found transaction fee in CSV: ${fee}%`);
                return fee;
              }
              break;
            }
          }
        }
      }
    } catch (error) {
      console.warn('Error extracting transaction fee:', error);
    }
    console.log('📊 No transaction fee found in CSV!');
    return null; // No default
  }

  // Extract LTV ratio from CSV
  extractLTVRatio(text) {
    try {
      if (!text || typeof text !== 'string') return 70;
      
      console.log('📊 Extracting LTV ratio from text');
      
      // Look for CSV format: "Acquisition LTV,,,,,80.00%"
      const lines = text.split('\n');
      for (const line of lines) {
        if (line.toLowerCase().includes('ltv') || line.toLowerCase().includes('loan to value')) {
          const parts = line.split(',');
          // Find the last non-empty part
          for (let i = parts.length - 1; i >= 0; i--) {
            if (parts[i].trim()) {
              const ltvStr = parts[i].trim().replace(/[^0-9.]/g, '');
              const ltv = parseFloat(ltvStr);
              if (!isNaN(ltv) && ltv >= 0 && ltv <= 100) {
                console.log(`📊 Found LTV ratio in CSV: ${ltv}%`);
                return ltv;
              }
              break;
            }
          }
        }
      }
    } catch (error) {
      console.warn('Error extracting LTV ratio:', error);
    }
    console.log('📊 No LTV ratio found in CSV!');
    return null; // No default
  }

  // Extract start date from CSV
  extractStartDate(text) {
    try {
      if (!text || typeof text !== 'string') return new Date().toISOString().split('T')[0];
      
      console.log('📊 Extracting start date from text');
      
      // Look for CSV format: "Acquisition date,,,,,31/03/2025"
      const lines = text.split('\n');
      for (const line of lines) {
        if (line.toLowerCase().includes('acquisition date') || line.toLowerCase().includes('start date')) {
          const parts = line.split(',');
          // Find the last non-empty part
          for (let i = parts.length - 1; i >= 0; i--) {
            if (parts[i].trim()) {
              const dateStr = parts[i].trim();
              // Handle DD/MM/YYYY format
              let date;
              if (dateStr.includes('/')) {
                const dateParts = dateStr.split('/');
                if (dateParts.length === 3) {
                  // Assume DD/MM/YYYY format
                  date = new Date(dateParts[2], dateParts[1] - 1, dateParts[0]);
                }
              } else {
                date = new Date(dateStr);
              }
              if (date && !isNaN(date.getTime())) {
                const isoDate = date.toISOString().split('T')[0];
                console.log(`📊 Found start date in CSV: ${isoDate}`);
                return isoDate;
              }
              break;
            }
          }
        }
      }
    } catch (error) {
      console.warn('Error extracting start date:', error);
    }
    console.log('📊 No start date found in CSV!');
    return null; // No default
  }

  // Extract end date from CSV
  extractEndDate(text) {
    try {
      if (!text || typeof text !== 'string') return new Date(Date.now() + 5*365*24*60*60*1000).toISOString().split('T')[0];
      
      console.log('📊 Extracting end date from text');
      
      // Look for CSV format or calculate from acquisition date + holding period
      const lines = text.split('\n');
      let acquisitionDate = null;
      let holdingPeriodMonths = null;
      
      // First try to find holding period and acquisition date
      for (const line of lines) {
        if (line.toLowerCase().includes('holding period')) {
          const parts = line.split(',');
          for (let i = parts.length - 1; i >= 0; i--) {
            if (parts[i].trim()) {
              const monthsStr = parts[i].trim().replace(/[^0-9.]/g, '');
              const months = parseFloat(monthsStr);
              if (!isNaN(months) && months > 0) {
                holdingPeriodMonths = months;
                console.log(`📊 Found holding period: ${months} months`);
                break;
              }
            }
          }
        }
        if (line.toLowerCase().includes('acquisition date')) {
          const parts = line.split(',');
          for (let i = parts.length - 1; i >= 0; i--) {
            if (parts[i].trim()) {
              const dateStr = parts[i].trim();
              if (dateStr.includes('/')) {
                const dateParts = dateStr.split('/');
                if (dateParts.length === 3) {
                  acquisitionDate = new Date(dateParts[2], dateParts[1] - 1, dateParts[0]);
                }
              } else {
                acquisitionDate = new Date(dateStr);
              }
              if (acquisitionDate && !isNaN(acquisitionDate.getTime())) {
                console.log(`📊 Found acquisition date: ${acquisitionDate.toISOString().split('T')[0]}`);
                break;
              }
            }
          }
        }
      }
      
      // Calculate end date from acquisition date + holding period
      if (acquisitionDate && holdingPeriodMonths) {
        const endDate = new Date(acquisitionDate);
        endDate.setMonth(endDate.getMonth() + holdingPeriodMonths);
        const isoDate = endDate.toISOString().split('T')[0];
        console.log(`📊 Calculated end date: ${isoDate} (${holdingPeriodMonths} months after acquisition)`);
        return isoDate;
      }
    } catch (error) {
      console.warn('Error extracting end date:', error);
    }
    console.log('📊 No end date found in CSV!');
    return null; // No default
  }

  // Extract reporting period from CSV
  extractReportingPeriod(text) {
    try {
      if (!text || typeof text !== 'string') return 'monthly';
      
      console.log('📊 Extracting reporting period from text');
      
      // For this CSV format, default to monthly for real estate deals
      console.log('📊 No specific reporting period found, defaulting to monthly for real estate');
      return 'monthly';
    } catch (error) {
      console.warn('Error extracting reporting period:', error);
    }
    console.log('📊 Using default reporting period: monthly');
    return 'monthly';
  }
  
  // Override reporting period method
  extractReportingPeriod(text) {
    try {
      if (!text || typeof text !== 'string') return 'monthly';
      
      console.log('📊 Extracting reporting period from text');
      
      // Look for CSV format: "Reporting Period,Monthly" 
      const lines = text.split('\n');
      for (const line of lines) {
        if (line.toLowerCase().includes('reporting period') || line.toLowerCase().includes('frequency')) {
          const parts = line.split(',');
          // Find the last non-empty part
          for (let i = parts.length - 1; i >= 0; i--) {
            if (parts[i].trim()) {
              const period = parts[i].trim().toLowerCase();
              if (['daily', 'monthly', 'quarterly', 'yearly', 'annual'].includes(period)) {
                const mapped = period === 'annual' ? 'yearly' : period;
                console.log(`📊 Found reporting period in CSV: ${mapped}`);
                return mapped;
              }
              break;
            }
          }
        }
      }
      
      // For real estate deals, typically monthly
      if (text.toLowerCase().includes('real estate') || text.toLowerCase().includes('office')) {
        console.log('📊 Real estate deal detected, using monthly reporting');
        return 'monthly';
      }
    } catch (error) {
      console.warn('Error extracting reporting period:', error);
    }
    console.log('📊 Using default reporting period: monthly');
    return 'monthly';
  }

  // Show extraction summary to user via UI alert/console
  showExtractionSummaryToUser(summary) {
    const summaryText = `
🔍 AI EXTRACTION RESULTS:

Company: ${summary.companyName || ''}
Deal Value: ${summary.dealValue || ''}
Currency: ${summary.currency || ''}
Transaction Fee: ${summary.transactionFee ? summary.transactionFee + '%' : ''}
LTV Ratio: ${summary.ltvRatio ? summary.ltvRatio + '%' : ''}
Start Date: ${summary.startDate || ''}
End Date: ${summary.endDate || ''}
Reporting: ${summary.reportingPeriod || ''}

Check console for detailed parsing logs.`;
    
    // Show to user via alert (you can change this to a better UI element)
    setTimeout(() => {
      alert(summaryText);
    }, 1000);
    
    return summary;
  }

  // Get the standardized data for other extractors to use
  getStandardizedData() {
    return this._data;
  }

  // Check if standardized data is available
  hasStandardizedData() {
    return this._data !== null;
  }

  // Clear standardized data (for new file uploads)
  clearStandardizedData() {
    this._data = null;
  }

  // Get specific section of standardized data
  getSection(sectionName) {
    if (!this._data) return null;
    return this._data[sectionName] || null;
  }

  // Get data quality information
  getDataQuality() {
    if (!this._data || !this._data.dataQuality) {
      return {
        overallConfidence: 0.3,
        dataSourceQuality: 'low'
      };
    }
    return this._data.dataQuality;
  }

  // Create summary for user feedback
  getAnalysisSummary() {
    if (!this._data) return null;
    
    const quality = this.getDataQuality();
    const company = this._data.companyOverview || {};
    const transaction = this._data.transactionDetails || {};
    
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