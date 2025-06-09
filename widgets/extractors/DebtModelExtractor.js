/**
 * DebtModelExtractor.js - Extract debt financing parameters and loan details
 * Handles: Interest rates, fees, loan terms, and debt structure
 */

class DebtModelExtractor {
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
    console.log('✅ DebtModelExtractor initialized');
  }

  /**
   * Extract debt model parameters from documents
   */
  async extract(files) {
    console.log('🏦 Extracting debt model parameters from', files.length, 'files');
    
    try {
      // Step 1: Use AI to extract debt data
      const aiExtraction = await this.extractWithAI(files);
      
      // Step 2: Enhance with pattern matching
      const enhancedData = this.enhanceWithParsing(aiExtraction, files);
      
      // Step 3: Validate and normalize debt parameters
      const validatedData = this.validateDebtParameters(enhancedData);
      
      // Step 4: Calculate derived metrics
      const enrichedData = this.enrichWithCalculations(validatedData);
      
      // Step 5: Score confidence
      const scoredData = this.scoreConfidence(enrichedData, files);
      
      // Step 6: Standardize the data
      const standardized = await this.standardizer.standardize(scoredData);
      
      console.log('🏦 Debt model extraction complete:', standardized);
      return standardized;
      
    } catch (error) {
      console.error('🏦 Error extracting debt model:', error);
      return this.getIntelligentDefaults(files);
    }
  }

  /**
   * Use AI service to extract debt parameters
   */
  async extractWithAI(files) {
    const prompt = `Extract debt financing and loan parameters from these financial documents.

Focus on identifying:

1. INTEREST RATE STRUCTURE:
   - Fixed interest rates (%)
   - Floating/Variable rates
   - Base rates (LIBOR, SOFR, Prime, etc.)
   - Credit spreads/margins (%)
   - Rate adjustment mechanisms

2. LOAN FEES:
   - Loan issuance/arrangement fees (%)
   - Origination fees
   - Commitment fees
   - Agency fees
   - Legal and due diligence costs

3. LOAN TERMS:
   - Loan amount/facility size
   - Loan term/maturity (years)
   - Amortization schedule
   - Prepayment penalties
   - Financial covenants

4. DEBT STRUCTURE:
   - Senior debt vs. subordinated
   - Term loans vs. revolving credit
   - Secured vs. unsecured
   - Currency of borrowing
   - Draw-down schedule

Look for loan agreements, term sheets, and financing proposals.
Extract actual values only - do not estimate or assume.

Return ONLY this structure with actual values found or null:
{
  "loanIssuanceFees": percentage_as_number_or_null,
  "interestRateType": "fixed|floating|null",
  "interestRate": percentage_as_number_or_null,
  "baseRate": percentage_as_number_or_null,
  "creditMargin": percentage_as_number_or_null,
  "loanTerm": numeric_years_or_null,
  "loanAmount": numeric_amount_or_null,
  "commitmentFee": percentage_as_number_or_null,
  "prepaymentPenalty": percentage_as_number_or_null,
  "debtType": "senior|subordinated|revolving|term|null",
  "debtCurrency": "USD|EUR|GBP|etc or null",
  "amortizationType": "bullet|linear|custom|null"
}`;

    try {
      const extraction = await this.extractionService.extractFromDocuments(
        files,
        'debtModel'
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
    
    // Extract interest rate information
    if (!enhanced.interestRateType || !enhanced.interestRate) {
      const rateInfo = this.extractInterestRates(allContent);
      if (rateInfo.type && !enhanced.interestRateType) {
        enhanced.interestRateType = rateInfo.type;
      }
      if (rateInfo.rate && !enhanced.interestRate) {
        enhanced.interestRate = rateInfo.rate;
      }
      if (rateInfo.baseRate && !enhanced.baseRate) {
        enhanced.baseRate = rateInfo.baseRate;
      }
      if (rateInfo.margin && !enhanced.creditMargin) {
        enhanced.creditMargin = rateInfo.margin;
      }
    }
    
    // Extract loan fees
    if (!enhanced.loanIssuanceFees) {
      const fees = this.extractLoanFees(allContent);
      if (fees) {
        enhanced.loanIssuanceFees = fees;
      }
    }
    
    // Extract loan terms
    if (!enhanced.loanTerm) {
      const term = this.extractLoanTerm(allContent);
      if (term) {
        enhanced.loanTerm = term;
      }
    }
    
    // Extract loan amount
    if (!enhanced.loanAmount) {
      const amount = this.extractLoanAmount(allContent);
      if (amount) {
        enhanced.loanAmount = amount;
      }
    }
    
    // Extract debt type
    if (!enhanced.debtType) {
      const type = this.extractDebtType(allContent);
      if (type) {
        enhanced.debtType = type;
      }
    }
    
    return enhanced;
  }

  /**
   * Extract interest rate information
   */
  extractInterestRates(text) {
    const rateInfo = {
      type: null,
      rate: null,
      baseRate: null,
      margin: null
    };
    
    // Fixed rate patterns
    const fixedRatePatterns = [
      /fixed\\s+(?:interest\\s+)?rate\\s*:?\\s*([0-9.]+)\\s*%/gi,
      /interest\\s+rate\\s*:?\\s*([0-9.]+)\\s*%\\s*(?:fixed|per\\s+annum)/gi,
      /([0-9.]+)\\s*%\\s*fixed\\s*(?:rate|interest)/gi
    ];
    
    for (const pattern of fixedRatePatterns) {
      const match = text.match(pattern);
      if (match) {
        const rate = parseFloat(match[1]);
        if (rate >= 0.1 && rate <= 25) {
          rateInfo.type = { value: 'fixed', confidence: 0.9, source: 'pattern_matching' };
          rateInfo.rate = { value: rate, confidence: 0.9, source: 'pattern_matching' };
          break;
        }
      }
    }
    
    // Floating rate patterns
    const floatingRatePatterns = [
      /(?:LIBOR|SOFR|prime)\\s*\\+\\s*([0-9.]+)\\s*%/gi,
      /floating\\s+rate\\s*:?\\s*([0-9.]+)\\s*%/gi,
      /variable\\s+rate\\s*:?\\s*([0-9.]+)\\s*%/gi,
      /(LIBOR|SOFR|prime)\\s*:?\\s*([0-9.]+)\\s*%/gi
    ];
    
    for (const pattern of floatingRatePatterns) {
      const match = text.match(pattern);
      if (match) {
        rateInfo.type = { value: 'floating', confidence: 0.9, source: 'pattern_matching' };
        
        if (pattern.source.includes('+')) {
          // Base rate + margin format
          const margin = parseFloat(match[1]);
          rateInfo.margin = { value: margin, confidence: 0.8, source: 'pattern_matching' };
        } else if (match[2]) {
          // Base rate identified
          const baseRate = parseFloat(match[2]);
          rateInfo.baseRate = { value: baseRate, confidence: 0.8, source: 'pattern_matching' };
        }
        break;
      }
    }
    
    // Credit spread/margin patterns
    const marginPatterns = [
      /credit\\s+(?:spread|margin)\\s*:?\\s*([0-9.]+)\\s*%/gi,
      /margin\\s*:?\\s*([0-9.]+)\\s*%/gi,
      /spread\\s+over\\s+(?:LIBOR|SOFR|prime)\\s*:?\\s*([0-9.]+)\\s*%/gi
    ];
    
    for (const pattern of marginPatterns) {
      const match = text.match(pattern);
      if (match) {
        const margin = parseFloat(match[1]);
        if (margin >= 0.1 && margin <= 15) {
          rateInfo.margin = { value: margin, confidence: 0.8, source: 'pattern_matching' };
          if (!rateInfo.type) {
            rateInfo.type = { value: 'floating', confidence: 0.7, source: 'inferred' };
          }
          break;
        }
      }
    }
    
    return rateInfo;
  }

  /**
   * Extract loan fees
   */
  extractLoanFees(text) {
    const feePatterns = [
      // Arrangement/Issuance fees
      {
        regex: /(?:loan\\s+)?(?:issuance|arrangement|origination)\\s+fees?\\s*:?\\s*([0-9.]+)\\s*%/gi,
        weight: 1.0
      },
      
      // Upfront fees
      {
        regex: /upfront\\s+fees?\\s*:?\\s*([0-9.]+)\\s*%/gi,
        weight: 0.9
      },
      
      // Commitment fees
      {
        regex: /commitment\\s+fees?\\s*:?\\s*([0-9.]+)\\s*%/gi,
        weight: 0.8
      },
      
      // General loan fees
      {
        regex: /loan\\s+fees?\\s*:?\\s*([0-9.]+)\\s*%/gi,
        weight: 0.7
      }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const pattern of feePatterns) {
      const match = text.match(pattern.regex);
      if (match) {
        const fee = parseFloat(match[1]);
        
        // Validate reasonable fee range (0.1% to 5%)
        if (fee >= 0.1 && fee <= 5 && pattern.weight > highestWeight) {
          highestWeight = pattern.weight;
          bestMatch = {
            value: fee,
            confidence: pattern.weight,
            source: 'pattern_matching'
          };
        }
      }
    }
    
    return bestMatch;
  }

  /**
   * Extract loan term
   */
  extractLoanTerm(text) {
    const termPatterns = [
      // Direct term mentions
      /(?:loan\\s+)?term\\s*:?\\s*([0-9]+)\\s*years?/gi,
      /maturity\\s*:?\\s*([0-9]+)\\s*years?/gi,
      /([0-9]+)\\s*year\\s*(?:loan|term|facility)/gi,
      
      // Tenor mentions
      /tenor\\s*:?\\s*([0-9]+)\\s*years?/gi
    ];
    
    for (const pattern of termPatterns) {
      const match = text.match(pattern);
      if (match) {
        const term = parseInt(match[1]);
        
        // Validate reasonable term (1-30 years)
        if (term >= 1 && term <= 30) {
          return {
            value: term,
            confidence: 0.8,
            source: 'pattern_matching'
          };
        }
      }
    }
    
    return null;
  }

  /**
   * Extract loan amount
   */
  extractLoanAmount(text) {
    const amountPatterns = [
      // Direct loan amount
      {
        regex: /(?:loan\\s+)?(?:amount|facility)\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?/gi,
        weight: 1.0
      },
      
      // Credit facility
      {
        regex: /credit\\s+facility\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?/gi,
        weight: 0.9
      },
      
      // Borrowing amount
      {
        regex: /borrowing\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?/gi,
        weight: 0.8
      }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const pattern of amountPatterns) {
      const match = text.match(pattern.regex);
      if (match) {
        let amount = parseFloat(match[1].replace(/,/g, ''));
        const unit = (match[2] || '').toLowerCase();
        
        // Apply multipliers
        if (unit === 'billion' || unit === 'b') {
          amount *= 1000000000;
        } else if (unit === 'million' || unit === 'm') {
          amount *= 1000000;
        }
        
        // Validate reasonable amount
        if (amount >= 100000 && amount <= 100000000000 && pattern.weight > highestWeight) {
          highestWeight = pattern.weight;
          bestMatch = {
            value: amount,
            confidence: pattern.weight,
            source: 'pattern_matching'
          };
        }
      }
    }
    
    return bestMatch;
  }

  /**
   * Extract debt type
   */
  extractDebtType(text) {
    const typePatterns = [
      { regex: /senior\\s+(?:debt|loan|facility)/gi, type: 'senior', weight: 1.0 },
      { regex: /subordinated\\s+(?:debt|loan)/gi, type: 'subordinated', weight: 1.0 },
      { regex: /revolving\\s+(?:credit|facility)/gi, type: 'revolving', weight: 0.9 },
      { regex: /term\\s+loan/gi, type: 'term', weight: 0.9 },
      { regex: /bridge\\s+(?:loan|facility)/gi, type: 'bridge', weight: 0.8 },
      { regex: /acquisition\\s+(?:loan|facility)/gi, type: 'acquisition', weight: 0.8 }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const pattern of typePatterns) {
      const match = text.match(pattern.regex);
      if (match && pattern.weight > highestWeight) {
        highestWeight = pattern.weight;
        bestMatch = {
          value: pattern.type,
          confidence: pattern.weight,
          source: 'pattern_matching'
        };
      }
    }
    
    return bestMatch;
  }

  /**
   * Validate debt parameters
   */
  validateDebtParameters(data) {
    const validated = { ...data };
    
    // Validate interest rates
    if (validated.interestRate?.value) {
      const rate = validated.interestRate.value;
      if (rate < 0.1 || rate > 25) {
        console.warn('🏦 Interest rate outside reasonable range:', rate);
        validated.interestRate.confidence *= 0.5;
      }
    }
    
    // Validate fees
    if (validated.loanIssuanceFees?.value) {
      const fee = validated.loanIssuanceFees.value;
      if (fee < 0.1 || fee > 10) {
        console.warn('🏦 Loan fees outside reasonable range:', fee);
        validated.loanIssuanceFees.confidence *= 0.5;
      }
    }
    
    // Validate loan term
    if (validated.loanTerm?.value) {
      const term = validated.loanTerm.value;
      if (term < 1 || term > 30) {
        console.warn('🏦 Loan term outside reasonable range:', term);
        validated.loanTerm.confidence *= 0.5;
      }
    }
    
    // Validate consistency between rate components
    if (validated.interestRateType?.value === 'floating' && 
        validated.baseRate?.value && validated.creditMargin?.value) {
      const totalRate = validated.baseRate.value + validated.creditMargin.value;
      if (validated.interestRate?.value && 
          Math.abs(totalRate - validated.interestRate.value) > 0.5) {
        console.warn('🏦 Inconsistent floating rate components');
        validated.interestRate.confidence *= 0.8;
      }
    }
    
    return validated;
  }

  /**
   * Enrich with calculated metrics
   */
  enrichWithCalculations(data) {
    const enriched = { ...data };
    
    // Calculate all-in interest rate for floating rates
    if (enriched.interestRateType?.value === 'floating' && 
        enriched.baseRate?.value && enriched.creditMargin?.value) {
      
      const allInRate = enriched.baseRate.value + enriched.creditMargin.value;
      
      if (!enriched.interestRate?.value) {
        enriched.interestRate = {
          value: allInRate,
          confidence: Math.min(enriched.baseRate.confidence, enriched.creditMargin.confidence),
          source: 'calculated'
        };
      }
      
      // Add calculated all-in rate as separate field
      enriched.allInRate = {
        value: allInRate,
        confidence: Math.min(enriched.baseRate.confidence, enriched.creditMargin.confidence),
        source: 'calculated'
      };
    }
    
    // Calculate annual debt service if we have rate and amount
    if (enriched.interestRate?.value && enriched.loanAmount?.value) {
      const annualInterest = enriched.loanAmount.value * (enriched.interestRate.value / 100);
      
      enriched.annualDebtService = {
        value: annualInterest,
        confidence: Math.min(enriched.interestRate.confidence, enriched.loanAmount.confidence),
        source: 'calculated'
      };
    }
    
    // Determine amortization type if not specified
    if (!enriched.amortizationType?.value && enriched.loanTerm?.value) {
      // Default assumptions based on loan type
      const defaultAmortization = enriched.debtType?.value === 'revolving' ? 'bullet' : 'linear';
      
      enriched.amortizationType = {
        value: defaultAmortization,
        confidence: 0.3,
        source: 'inferred'
      };
    }
    
    return enriched;
  }

  /**
   * Score confidence based on cross-validation
   */
  scoreConfidence(data, files) {
    const scored = {};
    
    for (const [field, value] of Object.entries(data)) {
      if (!value || value.value === null || value.value === undefined) {
        scored[field] = value;
        continue;
      }
      
      let confidence = value.confidence || 0.5;
      
      // Boost confidence for multiple file occurrences
      const occurrences = this.countOccurrences(value.value, files, field);
      if (occurrences > 1) {
        confidence = Math.min(confidence + 0.1 * (occurrences - 1), 1.0);
      }
      
      // Boost confidence for realistic values
      if (this.isRealisticValue(field, value.value)) {
        confidence = Math.min(confidence + 0.1, 1.0);
      }
      
      // Boost confidence for complete rate structure
      if (field === 'interestRate' && data.interestRateType?.value && 
          ((data.interestRateType.value === 'fixed') || 
           (data.interestRateType.value === 'floating' && data.baseRate?.value && data.creditMargin?.value))) {
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
    if (field === 'interestRateType') {
      const searchTerms = value === 'fixed' ? ['fixed rate', 'fixed interest'] : 
                         ['floating', 'variable', 'LIBOR', 'SOFR'];
      
      return files.filter(file => 
        file.content && searchTerms.some(term => 
          file.content.toLowerCase().includes(term.toLowerCase())
        )
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
      interestRate: (v) => v >= 0.5 && v <= 20, // 0.5% to 20%
      loanIssuanceFees: (v) => v >= 0.1 && v <= 5, // 0.1% to 5%
      baseRate: (v) => v >= 0.1 && v <= 10, // 0.1% to 10%
      creditMargin: (v) => v >= 0.1 && v <= 15, // 0.1% to 15%
      loanTerm: (v) => v >= 1 && v <= 25, // 1 to 25 years
      loanAmount: (v) => v >= 100000 && v <= 100000000000, // $100K to $100B
      commitmentFee: (v) => v >= 0.1 && v <= 2, // 0.1% to 2%
      prepaymentPenalty: (v) => v >= 0.1 && v <= 5 // 0.1% to 5%
    };
    
    const validator = realistic[field];
    return validator ? validator(value) : true;
  }

  /**
   * Get intelligent defaults
   */
  getIntelligentDefaults(files) {
    console.log('🏦 Using intelligent defaults for debt model');
    
    return {
      loanIssuanceFees: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      interestRateType: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      interestRate: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      baseRate: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      creditMargin: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      loanTerm: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      loanAmount: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      debtType: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      debtCurrency: {
        value: null,
        confidence: 0,
        source: 'not_found'
      }
    };
  }

  /**
   * Apply extracted debt model to form
   */
  async applyToForm(extractedData) {
    console.log('🏦 Applying debt model to form');
    
    return await this.mappingEngine.applyDataToForm(extractedData, {
      section: 'debtModel',
      showConfidence: true,
      animateChanges: true
    });
  }
}

// Export for use
window.DebtModelExtractor = DebtModelExtractor;