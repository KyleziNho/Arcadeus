/**
 * DealAssumptionsExtractor.js - Extract M&A deal assumptions
 * Handles: Deal value, LTV, transaction fees, deal name, equity/debt split
 */

class DealAssumptionsExtractor {
  constructor() {
    this.extractionService = null;
    this.standardizer = null;
    this.mappingEngine = null;
    this.confidence = {
      high: 0.8,
      medium: 0.5,
      low: 0.3
    };
  }

  initialize(services) {
    this.extractionService = services.extractionService;
    this.standardizer = services.standardizer;
    this.mappingEngine = services.mappingEngine;
    console.log('✅ DealAssumptionsExtractor initialized');
  }

  /**
   * Extract deal assumptions from documents
   */
  async extract(files) {
    console.log('💼 Extracting deal assumptions from', files.length, 'files');
    
    try {
      // Step 1: Use AI to extract deal-specific data
      const aiExtraction = await this.extractWithAI(files);
      
      // Step 2: Enhance with intelligent parsing
      const enhancedData = this.enhanceWithParsing(aiExtraction, files);
      
      // Step 3: Validate deal assumptions
      const validatedData = this.validateDealAssumptions(enhancedData);
      
      // Step 4: Calculate derived values
      const completeData = this.calculateDerivedValues(validatedData);
      
      // Step 5: Score confidence
      const scoredData = this.scoreConfidence(completeData, files);
      
      // Step 6: Standardize the data
      const standardized = await this.standardizer.standardize(scoredData);
      
      console.log('💼 Deal assumptions extraction complete:', standardized);
      return standardized;
      
    } catch (error) {
      console.error('💼 Error extracting deal assumptions:', error);
      return this.getIntelligentDefaults(files);
    }
  }

  /**
   * Use AI service to extract deal assumptions
   */
  async extractWithAI(files) {
    const prompt = `Extract M&A deal assumptions from these financial documents.

Focus on identifying:

1. DEAL IDENTIFICATION:
   - Company/Target name
   - Transaction/Deal name
   - Deal description

2. FINANCIAL STRUCTURE:
   - Deal Value / Enterprise Value / Transaction Value (in millions/billions)
   - Equity contribution amount
   - Debt financing amount
   - Total consideration

3. TRANSACTION COSTS:
   - Transaction fees (as percentage)
   - Advisory fees
   - Investment banking fees
   - Due diligence costs

4. LEVERAGE STRUCTURE:
   - LTV (Loan-to-Value) ratio
   - Debt-to-equity ratio
   - Leverage multiple
   - Financing split

Look for explicit numbers, percentages, and clear deal terms.
Extract actual values only - do not estimate or assume.

Return ONLY these fields with actual values found or null:
{
  "dealName": "actual company/deal name or null",
  "dealValue": numeric_value_or_null,
  "transactionFee": percentage_as_number_or_null,
  "dealLTV": percentage_as_number_or_null,
  "equityContribution": numeric_value_or_null,
  "debtFinancing": numeric_value_or_null
}`;

    try {
      const extraction = await this.extractionService.extractFromDocuments(
        files,
        'dealAssumptions'
      );
      
      return extraction;
    } catch (error) {
      console.error('AI extraction failed:', error);
      return {};
    }
  }

  /**
   * Enhance AI extraction with pattern matching
   */
  enhanceWithParsing(aiData, files) {
    const enhanced = { ...aiData };
    
    // Combine all file contents
    const allContent = files
      .map(f => f.content || '')
      .join('\n');
    
    // Extract deal name if not found
    if (!enhanced.dealName?.value) {
      const dealName = this.extractDealName(files);
      if (dealName) {
        enhanced.dealName = dealName;
      }
    }
    
    // Extract deal value if not found
    if (!enhanced.dealValue?.value) {
      const dealValue = this.extractDealValue(allContent);
      if (dealValue) {
        enhanced.dealValue = dealValue;
      }
    }
    
    // Extract transaction fee if not found
    if (!enhanced.transactionFee?.value) {
      const transactionFee = this.extractTransactionFee(allContent);
      if (transactionFee) {
        enhanced.transactionFee = transactionFee;
      }
    }
    
    // Extract LTV if not found
    if (!enhanced.dealLTV?.value) {
      const ltv = this.extractLTV(allContent);
      if (ltv) {
        enhanced.dealLTV = ltv;
      }
    }
    
    // Extract equity/debt split if not found
    if (!enhanced.equityContribution?.value || !enhanced.debtFinancing?.value) {
      const financing = this.extractFinancingSplit(allContent);
      if (financing.equity && !enhanced.equityContribution?.value) {
        enhanced.equityContribution = financing.equity;
      }
      if (financing.debt && !enhanced.debtFinancing?.value) {
        enhanced.debtFinancing = financing.debt;
      }
    }
    
    return enhanced;
  }

  /**
   * Extract deal name from file names and content
   */
  extractDealName(files) {
    // Try filename first (most reliable)
    for (const file of files) {
      const filename = file.name;
      
      // Clean filename
      const cleaned = filename
        .replace(/\.(csv|pdf|xlsx?|png|jpe?g)$/i, '')
        .replace(/[-_]/g, ' ')
        .trim();
      
      // Skip generic names
      const genericNames = ['data', 'model', 'financial', 'analysis', 'report', 'document', 'file'];
      const isGeneric = genericNames.some(name => 
        cleaned.toLowerCase().includes(name) && cleaned.length < 15
      );
      
      if (!isGeneric && cleaned.length > 3 && cleaned.length < 100) {
        return {
          value: cleaned,
          confidence: 0.8,
          source: 'filename'
        };
      }
    }
    
    // Try content extraction
    const allContent = files.map(f => f.content || '').join('\n');
    
    // Look for company patterns
    const companyPatterns = [
      // Header patterns
      /^([A-Z][A-Za-z\s&.,]+(?:Inc|Corp|LLC|Ltd|Company|Co\.?|Corporation|Limited))/m,
      
      // Deal patterns
      /(?:acquisition of|target company|company name)\s*:?\s*([A-Za-z\s&.,]+)/i,
      /(?:transaction|deal)\s*:?\s*([A-Za-z\s&.,]+(?:acquisition|purchase|merger))/i,
      
      // Title patterns
      /^([A-Z][A-Za-z\s&.,]+ (?:Acquisition|Purchase|Merger|Transaction))/m
    ];
    
    for (const pattern of companyPatterns) {
      const match = allContent.match(pattern);
      if (match) {
        const name = match[1].trim();
        if (name.length > 3 && name.length < 100) {
          return {
            value: name,
            confidence: 0.7,
            source: 'content_pattern'
          };
        }
      }
    }
    
    return null;
  }

  /**
   * Extract deal value with multiple formats
   */
  extractDealValue(text) {
    const patterns = [
      // Explicit deal value mentions
      {
        regex: /(?:deal value|transaction value|enterprise value|purchase price|total consideration)\s*:?\s*(?:\$|USD)?\s*([\d,]+(?:\.\d+)?)\s*(million|billion|m|b)?/i,
        weight: 1.0
      },
      
      // Large monetary amounts with context
      {
        regex: /(?:for|of|worth)\s+(?:\$|USD)?\s*([\d,]+(?:\.\d+)?)\s*(million|billion|m|b)/i,
        weight: 0.8
      },
      
      // Investment amounts
      {
        regex: /(?:invest|investing|investment of)\s+(?:\$|USD)?\s*([\d,]+(?:\.\d+)?)\s*(million|billion|m|b)/i,
        weight: 0.7
      },
      
      // CSV-style extraction (label,value)
      {
        regex: /(?:deal value|transaction value|enterprise value)\s*,\s*(?:\$|USD)?\s*([\d,]+(?:\.\d+)?)/i,
        weight: 0.9
      }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const pattern of patterns) {
      const match = text.match(pattern.regex);
      if (match) {
        let value = parseFloat(match[1].replace(/,/g, ''));
        const unit = (match[2] || '').toLowerCase();
        
        // Apply multipliers
        if (unit === 'billion' || unit === 'b') {
          value *= 1000000000;
        } else if (unit === 'million' || unit === 'm') {
          value *= 1000000;
        }
        
        // Validate reasonable deal size (1M to 1T)
        if (value >= 1000000 && value <= 1000000000000) {
          if (pattern.weight > highestWeight) {
            highestWeight = pattern.weight;
            bestMatch = {
              value: value,
              confidence: pattern.weight,
              source: 'pattern_matching'
            };
          }
        }
      }
    }
    
    return bestMatch;
  }

  /**
   * Extract transaction fees
   */
  extractTransactionFee(text) {
    const patterns = [
      // Direct percentage mentions
      {
        regex: /(?:transaction fee|advisory fee|investment banking fee|fees?)\s*:?\s*([\d.]+)\s*%/i,
        weight: 1.0
      },
      
      // Fee descriptions
      {
        regex: /(?:fees? of|fee is)\s*([\d.]+)\s*(?:percent|%)/i,
        weight: 0.9
      },
      
      // CSV format
      {
        regex: /(?:transaction fee|advisory fee)\s*,\s*([\d.]+)%?/i,
        weight: 0.8
      }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const pattern of patterns) {
      const match = text.match(pattern.regex);
      if (match) {
        const fee = parseFloat(match[1]);
        
        // Validate reasonable fee range (0.1% to 10%)
        if (fee >= 0.1 && fee <= 10) {
          if (pattern.weight > highestWeight) {
            highestWeight = pattern.weight;
            bestMatch = {
              value: fee,
              confidence: pattern.weight,
              source: 'pattern_matching'
            };
          }
        }
      }
    }
    
    return bestMatch;
  }

  /**
   * Extract LTV (Loan-to-Value) ratio
   */
  extractLTV(text) {
    const patterns = [
      // Direct LTV mentions
      {
        regex: /(?:ltv|loan[- ]to[- ]value|leverage ratio)\s*:?\s*([\d.]+)\s*%/i,
        weight: 1.0
      },
      
      // Debt ratio mentions
      {
        regex: /(?:debt ratio|debt[- ]to[- ]value)\s*:?\s*([\d.]+)\s*%/i,
        weight: 0.9
      },
      
      // CSV format
      {
        regex: /(?:ltv|loan to value|leverage)\s*,\s*([\d.]+)%?/i,
        weight: 0.8
      },
      
      // Financing descriptions
      {
        regex: /([\d.]+)\s*%\s*(?:debt financing|leverage)/i,
        weight: 0.7
      }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const pattern of patterns) {
      const match = text.match(pattern.regex);
      if (match) {
        const ltv = parseFloat(match[1]);
        
        // Validate reasonable LTV range (10% to 95%)
        if (ltv >= 10 && ltv <= 95) {
          if (pattern.weight > highestWeight) {
            highestWeight = pattern.weight;
            bestMatch = {
              value: ltv,
              confidence: pattern.weight,
              source: 'pattern_matching'
            };
          }
        }
      }
    }
    
    return bestMatch;
  }

  /**
   * Extract equity and debt financing amounts
   */
  extractFinancingSplit(text) {
    const result = {
      equity: null,
      debt: null
    };
    
    // Equity patterns
    const equityPatterns = [
      /(?:equity contribution|equity investment|equity financing)\s*:?\s*(?:\$|USD)?\s*([\d,]+(?:\.\d+)?)\s*(million|billion|m|b)?/i,
      /(?:sponsor equity|initial equity)\s*:?\s*(?:\$|USD)?\s*([\d,]+(?:\.\d+)?)\s*(million|billion|m|b)?/i
    ];
    
    // Debt patterns
    const debtPatterns = [
      /(?:debt financing|debt amount|loan amount)\s*:?\s*(?:\$|USD)?\s*([\d,]+(?:\.\d+)?)\s*(million|billion|m|b)?/i,
      /(?:senior debt|total debt|borrowing)\s*:?\s*(?:\$|USD)?\s*([\d,]+(?:\.\d+)?)\s*(million|billion|m|b)?/i
    ];
    
    // Process equity
    for (const pattern of equityPatterns) {
      const match = text.match(pattern);
      if (match) {
        let value = parseFloat(match[1].replace(/,/g, ''));
        const unit = (match[2] || '').toLowerCase();
        
        if (unit === 'billion' || unit === 'b') {
          value *= 1000000000;
        } else if (unit === 'million' || unit === 'm') {
          value *= 1000000;
        }
        
        if (value >= 100000) { // Minimum $100k
          result.equity = {
            value: value,
            confidence: 0.8,
            source: 'pattern_matching'
          };
          break;
        }
      }
    }
    
    // Process debt
    for (const pattern of debtPatterns) {
      const match = text.match(pattern);
      if (match) {
        let value = parseFloat(match[1].replace(/,/g, ''));
        const unit = (match[2] || '').toLowerCase();
        
        if (unit === 'billion' || unit === 'b') {
          value *= 1000000000;
        } else if (unit === 'million' || unit === 'm') {
          value *= 1000000;
        }
        
        if (value >= 100000) { // Minimum $100k
          result.debt = {
            value: value,
            confidence: 0.8,
            source: 'pattern_matching'
          };
          break;
        }
      }
    }
    
    return result;
  }

  /**
   * Validate deal assumptions for consistency
   */
  validateDealAssumptions(data) {
    const validated = { ...data };
    
    // Check deal value consistency
    if (validated.dealValue?.value && validated.equityContribution?.value && validated.debtFinancing?.value) {
      const totalFinancing = validated.equityContribution.value + validated.debtFinancing.value;
      const dealValue = validated.dealValue.value;
      
      // If totals don't match within 5%, adjust confidence
      const variance = Math.abs(totalFinancing - dealValue) / dealValue;
      if (variance > 0.05) {
        console.warn('💼 Deal value and financing split inconsistent');
        validated.dealValue.confidence *= 0.8;
        validated.equityContribution.confidence *= 0.8;
        validated.debtFinancing.confidence *= 0.8;
      }
    }
    
    // Validate LTV consistency
    if (validated.dealLTV?.value && validated.dealValue?.value && validated.debtFinancing?.value) {
      const calculatedLTV = (validated.debtFinancing.value / validated.dealValue.value) * 100;
      const variance = Math.abs(calculatedLTV - validated.dealLTV.value) / validated.dealLTV.value;
      
      if (variance > 0.1) {
        console.warn('💼 LTV and debt financing inconsistent');
        validated.dealLTV.confidence *= 0.8;
      }
    }
    
    return validated;
  }

  /**
   * Calculate derived values
   */
  calculateDerivedValues(data) {
    const derived = { ...data };
    
    // Calculate missing equity/debt from deal value and LTV
    if (derived.dealValue?.value && derived.dealLTV?.value) {
      const dealValue = derived.dealValue.value;
      const ltvPercent = derived.dealLTV.value;
      
      if (!derived.debtFinancing?.value) {
        derived.debtFinancing = {
          value: dealValue * (ltvPercent / 100),
          confidence: Math.min(derived.dealValue.confidence, derived.dealLTV.confidence),
          source: 'calculated'
        };
      }
      
      if (!derived.equityContribution?.value) {
        derived.equityContribution = {
          value: dealValue * (1 - ltvPercent / 100),
          confidence: Math.min(derived.dealValue.confidence, derived.dealLTV.confidence),
          source: 'calculated'
        };
      }
    }
    
    // Calculate missing deal value from equity + debt
    if (!derived.dealValue?.value && derived.equityContribution?.value && derived.debtFinancing?.value) {
      derived.dealValue = {
        value: derived.equityContribution.value + derived.debtFinancing.value,
        confidence: Math.min(derived.equityContribution.confidence, derived.debtFinancing.confidence),
        source: 'calculated'
      };
    }
    
    // Calculate missing LTV from debt and deal value
    if (!derived.dealLTV?.value && derived.debtFinancing?.value && derived.dealValue?.value) {
      derived.dealLTV = {
        value: (derived.debtFinancing.value / derived.dealValue.value) * 100,
        confidence: Math.min(derived.debtFinancing.confidence, derived.dealValue.confidence),
        source: 'calculated'
      };
    }
    
    return derived;
  }

  /**
   * Score confidence based on cross-validation
   */
  scoreConfidence(data, files) {
    const scored = {};
    
    for (const [field, value] of Object.entries(data)) {
      if (!value || !value.value) {
        scored[field] = value;
        continue;
      }
      
      let confidence = value.confidence || 0.5;
      
      // Boost confidence for values found in multiple files
      const occurrences = this.countOccurrences(value.value, files, field);
      if (occurrences > 1) {
        confidence = Math.min(confidence + 0.1 * (occurrences - 1), 1.0);
      }
      
      // Boost confidence for realistic values
      if (this.isRealisticValue(field, value.value)) {
        confidence = Math.min(confidence + 0.1, 1.0);
      }
      
      scored[field] = {
        ...value,
        confidence: confidence
      };
    }
    
    return scored;
  }

  /**
   * Count occurrences with field-specific logic
   */
  countOccurrences(value, files, field) {
    if (field === 'dealName') {
      return files.filter(file => 
        file.name.toLowerCase().includes(String(value).toLowerCase()) ||
        (file.content && file.content.toLowerCase().includes(String(value).toLowerCase()))
      ).length;
    }
    
    const searchValue = String(value).toLowerCase();
    return files.filter(file => 
      file.content && file.content.toLowerCase().includes(searchValue)
    ).length;
  }

  /**
   * Check if value is realistic for the field
   */
  isRealisticValue(field, value) {
    const realistic = {
      dealValue: (v) => v >= 1000000 && v <= 1000000000000, // $1M to $1T
      transactionFee: (v) => v >= 0.1 && v <= 10, // 0.1% to 10%
      dealLTV: (v) => v >= 10 && v <= 95, // 10% to 95%
      equityContribution: (v) => v >= 100000 && v <= 500000000000, // $100K to $500B
      debtFinancing: (v) => v >= 100000 && v <= 500000000000 // $100K to $500B
    };
    
    const validator = realistic[field];
    return validator ? validator(value) : true;
  }

  /**
   * Get intelligent defaults
   */
  getIntelligentDefaults(files) {
    console.log('💼 Using intelligent defaults for deal assumptions');
    
    return {
      dealName: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      dealValue: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      transactionFee: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      dealLTV: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      equityContribution: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      debtFinancing: {
        value: null,
        confidence: 0,
        source: 'not_found'
      }
    };
  }

  /**
   * Apply extracted data to form
   */
  async applyToForm(extractedData) {
    console.log('💼 Applying deal assumptions to form');
    
    return await this.mappingEngine.applyDataToForm(extractedData, {
      section: 'dealAssumptions',
      showConfidence: true,
      animateChanges: true
    });
  }
}

// Export for use
window.DealAssumptionsExtractor = DealAssumptionsExtractor;