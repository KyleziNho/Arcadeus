/**
 * ExitAssumptionsExtractor.js - Extract exit and valuation assumptions
 * Handles: Exit costs, terminal cap rates, exit multiples, and expected returns
 */

class ExitAssumptionsExtractor {
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
    console.log('✅ ExitAssumptionsExtractor initialized');
  }

  /**
   * Extract exit assumptions from documents
   */
  async extract(files) {
    console.log('🚪 Extracting exit assumptions from', files.length, 'files');
    
    try {
      // Step 1: Use AI to extract exit data
      const aiExtraction = await this.extractWithAI(files);
      
      // Step 2: Enhance with pattern matching
      const enhancedData = this.enhanceWithParsing(aiExtraction, files);
      
      // Step 3: Validate and normalize exit parameters
      const validatedData = this.validateExitParameters(enhancedData);
      
      // Step 4: Calculate derived metrics
      const enrichedData = this.enrichWithCalculations(validatedData);
      
      // Step 5: Score confidence
      const scoredData = this.scoreConfidence(enrichedData, files);
      
      // Step 6: Standardize the data
      const standardized = await this.standardizer.standardize(scoredData);
      
      console.log('🚪 Exit assumptions extraction complete:', standardized);
      return standardized;
      
    } catch (error) {
      console.error('🚪 Error extracting exit assumptions:', error);
      return this.getIntelligentDefaults(files);
    }
  }

  /**
   * Use AI service to extract exit assumptions
   */
  async extractWithAI(files) {
    const prompt = `Extract exit and valuation assumptions from these financial documents.

Focus on identifying:

1. EXIT COSTS:
   - Disposal costs (%)
   - Transaction fees on exit
   - Legal and advisory fees
   - Tax implications
   - Break-up fees

2. VALUATION METRICS:
   - Terminal capitalization rates (%)
   - Exit multiples (EV/EBITDA, P/E, etc.)
   - Terminal growth rates
   - Discount rates for DCF
   - Risk premiums

3. EXIT STRATEGY:
   - Expected exit date/timing
   - Exit route (IPO, trade sale, secondary buyout)
   - Target buyers or market
   - Exit value expectations
   - IRR targets

4. MARKET ASSUMPTIONS:
   - Market conditions at exit
   - Comparable transactions
   - Industry multiples
   - Market growth rates
   - Sector outlook

Look for investment memos, valuation models, and exit strategy documents.
Extract actual values only - do not estimate or assume.

Return ONLY this structure with actual values found or null:
{
  "disposalCost": percentage_as_number_or_null,
  "terminalCapRate": percentage_as_number_or_null,
  "exitMultiple": numeric_multiple_or_null,
  "exitMultipleType": "EV/EBITDA|P/E|EV/Revenue|other|null",
  "terminalGrowthRate": percentage_as_number_or_null,
  "discountRate": percentage_as_number_or_null,
  "expectedExitDate": "YYYY-MM-DD or null",
  "exitRoute": "IPO|trade_sale|secondary_buyout|other|null",
  "targetIRR": percentage_as_number_or_null,
  "exitValue": numeric_value_or_null,
  "holdingPeriod": numeric_years_or_null
}`;

    try {
      const extraction = await this.extractionService.extractFromDocuments(
        files,
        'exitAssumptions'
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
    
    // Extract disposal costs
    if (!enhanced.disposalCost) {
      const disposalCost = this.extractDisposalCost(allContent);
      if (disposalCost) {
        enhanced.disposalCost = disposalCost;
      }
    }
    
    // Extract terminal cap rate
    if (!enhanced.terminalCapRate) {
      const capRate = this.extractTerminalCapRate(allContent);
      if (capRate) {
        enhanced.terminalCapRate = capRate;
      }
    }
    
    // Extract exit multiples
    if (!enhanced.exitMultiple || !enhanced.exitMultipleType) {
      const multiple = this.extractExitMultiple(allContent);
      if (multiple.value && !enhanced.exitMultiple) {
        enhanced.exitMultiple = multiple.value;
      }
      if (multiple.type && !enhanced.exitMultipleType) {
        enhanced.exitMultipleType = multiple.type;
      }
    }
    
    // Extract IRR targets
    if (!enhanced.targetIRR) {
      const irr = this.extractTargetIRR(allContent);
      if (irr) {
        enhanced.targetIRR = irr;
      }
    }
    
    // Extract exit dates
    if (!enhanced.expectedExitDate) {
      const exitDate = this.extractExitDate(allContent);
      if (exitDate) {
        enhanced.expectedExitDate = exitDate;
      }
    }
    
    // Extract exit route
    if (!enhanced.exitRoute) {
      const route = this.extractExitRoute(allContent);
      if (route) {
        enhanced.exitRoute = route;
      }
    }
    
    return enhanced;
  }

  /**
   * Extract disposal costs
   */
  extractDisposalCost(text) {
    const patterns = [
      // Direct disposal cost mentions
      {
        regex: /disposal\s+costs?\s*:?\s*([0-9.]+)\s*%/gi,
        weight: 1.0
      },
      
      // Exit costs
      {
        regex: /exit\s+costs?\s*:?\s*([0-9.]+)\s*%/gi,
        weight: 0.9
      },
      
      // Transaction costs on exit
      {
        regex: /transaction\s+costs?\s+on\s+exit\s*:?\s*([0-9.]+)\s*%/gi,
        weight: 0.9
      },
      
      // Selling costs
      {
        regex: /selling\s+costs?\s*:?\s*([0-9.]+)\s*%/gi,
        weight: 0.8
      }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const pattern of patterns) {
      const match = text.match(pattern.regex);
      if (match) {
        const cost = parseFloat(match[1]);
        
        // Validate reasonable disposal cost (0.1% to 10%)
        if (cost >= 0.1 && cost <= 10 && pattern.weight > highestWeight) {
          highestWeight = pattern.weight;
          bestMatch = {
            value: cost,
            confidence: pattern.weight,
            source: 'pattern_matching'
          };
        }
      }
    }
    
    return bestMatch;
  }

  /**
   * Extract terminal capitalization rate
   */
  extractTerminalCapRate(text) {
    const patterns = [
      // Direct terminal cap rate
      {
        regex: /terminal\s+(?:cap\s+rate|capitali[sz]ation\s+rate)\s*:?\s*([0-9.]+)\s*%/gi,
        weight: 1.0
      },
      
      // Exit cap rate
      {
        regex: /exit\s+cap\s+rate\s*:?\s*([0-9.]+)\s*%/gi,
        weight: 0.9
      },
      
      // Cap rate on exit
      {
        regex: /cap\s+rate\s+on\s+exit\s*:?\s*([0-9.]+)\s*%/gi,
        weight: 0.9
      },
      
      // Terminal yield
      {
        regex: /terminal\s+yield\s*:?\s*([0-9.]+)\s*%/gi,
        weight: 0.8
      }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const pattern of patterns) {
      const match = text.match(pattern.regex);
      if (match) {
        const rate = parseFloat(match[1]);
        
        // Validate reasonable cap rate (3% to 20%)
        if (rate >= 3 && rate <= 20 && pattern.weight > highestWeight) {
          highestWeight = pattern.weight;
          bestMatch = {
            value: rate,
            confidence: pattern.weight,
            source: 'pattern_matching'
          };
        }
      }
    }
    
    return bestMatch;
  }

  /**
   * Extract exit multiples
   */
  extractExitMultiple(text) {
    const result = {
      value: null,
      type: null
    };
    
    const multiplePatterns = [
      // EV/EBITDA
      {
        regex: /(?:exit\s+)?(?:EV\/EBITDA|enterprise\s+value\s+to\s+EBITDA)\s*:?\s*([0-9.]+)x?/gi,
        type: 'EV/EBITDA',
        weight: 1.0
      },
      
      // P/E ratio
      {
        regex: /(?:exit\s+)?(?:P\/E|price\s+to\s+earnings)\s*:?\s*([0-9.]+)x?/gi,
        type: 'P/E',
        weight: 0.9
      },
      
      // EV/Revenue
      {
        regex: /(?:exit\s+)?(?:EV\/Revenue|enterprise\s+value\s+to\s+revenue)\s*:?\s*([0-9.]+)x?/gi,
        type: 'EV/Revenue',
        weight: 0.9
      },
      
      // Generic exit multiple
      {
        regex: /exit\s+multiple\s*:?\s*([0-9.]+)x?/gi,
        type: 'other',
        weight: 0.7
      }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const pattern of multiplePatterns) {
      const match = text.match(pattern.regex);
      if (match) {
        const multiple = parseFloat(match[1]);
        
        // Validate reasonable multiple range
        const isReasonable = this.isReasonableMultiple(pattern.type, multiple);
        
        if (isReasonable && pattern.weight > highestWeight) {
          highestWeight = pattern.weight;
          bestMatch = {
            value: multiple,
            type: pattern.type,
            confidence: pattern.weight
          };
        }
      }
    }
    
    if (bestMatch) {
      result.value = {
        value: bestMatch.value,
        confidence: bestMatch.confidence,
        source: 'pattern_matching'
      };
      result.type = {
        value: bestMatch.type,
        confidence: bestMatch.confidence,
        source: 'pattern_matching'
      };
    }
    
    return result;
  }

  /**
   * Extract target IRR
   */
  extractTargetIRR(text) {
    const patterns = [
      // Direct IRR target
      {
        regex: /target\s+IRR\s*:?\s*([0-9.]+)\s*%/gi,
        weight: 1.0
      },
      
      // Expected IRR
      {
        regex: /expected\s+IRR\s*:?\s*([0-9.]+)\s*%/gi,
        weight: 0.9
      },
      
      // IRR assumption
      {
        regex: /IRR\s+assumption\s*:?\s*([0-9.]+)\s*%/gi,
        weight: 0.9
      },
      
      // Return target
      {
        regex: /return\s+target\s*:?\s*([0-9.]+)\s*%\s*IRR/gi,
        weight: 0.8
      }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const pattern of patterns) {
      const match = text.match(pattern.regex);
      if (match) {
        const irr = parseFloat(match[1]);
        
        // Validate reasonable IRR (5% to 50%)
        if (irr >= 5 && irr <= 50 && pattern.weight > highestWeight) {
          highestWeight = pattern.weight;
          bestMatch = {
            value: irr,
            confidence: pattern.weight,
            source: 'pattern_matching'
          };
        }
      }
    }
    
    return bestMatch;
  }

  /**
   * Extract exit date
   */
  extractExitDate(text) {
    const patterns = [
      // Expected exit date
      /expected\s+exit\s+(?:date)?\s*:?\s*(\d{4}-\d{2}-\d{2}|\d{1,2}\/\d{1,2}\/\d{4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d{4})/gi,
      
      // Exit in [year]
      /exit\s+in\s+(\d{4})/gi,
      
      // Sale date
      /(?:anticipated\s+)?sale\s+date\s*:?\s*(\d{4}-\d{2}-\d{2}|\d{1,2}\/\d{1,2}\/\d{4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d{4})/gi
    ];
    
    for (const pattern of patterns) {
      const match = text.match(pattern);
      if (match) {
        const dateStr = match[1];
        
        // Validate and standardize date
        const standardizedDate = this.standardizeExitDate(dateStr);
        if (standardizedDate) {
          return {
            value: standardizedDate,
            confidence: 0.8,
            source: 'pattern_matching'
          };
        }
      }
    }
    
    return null;
  }

  /**
   * Extract exit route
   */
  extractExitRoute(text) {
    const routes = [
      { pattern: /IPO|initial\\s+public\\s+offering|public\\s+listing/gi, value: 'IPO', weight: 1.0 },
      { pattern: /trade\\s+sale|strategic\\s+sale|acquisition\\s+by/gi, value: 'trade_sale', weight: 1.0 },
      { pattern: /secondary\\s+buyout|private\\s+equity\\s+sale/gi, value: 'secondary_buyout', weight: 1.0 },
      { pattern: /management\\s+buyout|MBO/gi, value: 'management_buyout', weight: 0.9 },
      { pattern: /refinancing|recap/gi, value: 'refinancing', weight: 0.8 }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const route of routes) {
      const match = text.match(route.pattern);
      if (match && route.weight > highestWeight) {
        highestWeight = route.weight;
        bestMatch = {
          value: route.value,
          confidence: route.weight,
          source: 'pattern_matching'
        };
      }
    }
    
    return bestMatch;
  }

  /**
   * Check if multiple is reasonable for its type
   */
  isReasonableMultiple(type, value) {
    const ranges = {
      'EV/EBITDA': [3, 25],
      'P/E': [5, 40],
      'EV/Revenue': [0.5, 10],
      'other': [1, 50]
    };
    
    const range = ranges[type] || ranges['other'];
    return value >= range[0] && value <= range[1];
  }

  /**
   * Standardize exit date format
   */
  standardizeExitDate(dateStr) {
    // Handle year-only format
    if (/^\d{4}$/.test(dateStr)) {
      return `${dateStr}-12-31`; // Assume end of year
    }
    
    // Try to parse and format as YYYY-MM-DD
    try {
      const date = new Date(dateStr);
      if (!isNaN(date.getTime())) {
        return date.toISOString().split('T')[0];
      }
    } catch (error) {
      console.warn('Could not parse exit date:', dateStr);
    }
    
    return null;
  }

  /**
   * Validate exit parameters
   */
  validateExitParameters(data) {
    const validated = { ...data };
    
    // Validate disposal cost
    if (validated.disposalCost?.value) {
      const cost = validated.disposalCost.value;
      if (cost < 0.1 || cost > 15) {
        console.warn('🚪 Disposal cost outside reasonable range:', cost);
        validated.disposalCost.confidence *= 0.5;
      }
    }
    
    // Validate terminal cap rate
    if (validated.terminalCapRate?.value) {
      const rate = validated.terminalCapRate.value;
      if (rate < 2 || rate > 25) {
        console.warn('🚪 Terminal cap rate outside reasonable range:', rate);
        validated.terminalCapRate.confidence *= 0.5;
      }
    }
    
    // Validate IRR target
    if (validated.targetIRR?.value) {
      const irr = validated.targetIRR.value;
      if (irr < 5 || irr > 50) {
        console.warn('🚪 Target IRR outside reasonable range:', irr);
        validated.targetIRR.confidence *= 0.5;
      }
    }
    
    // Validate exit date is in the future
    if (validated.expectedExitDate?.value) {
      const exitDate = new Date(validated.expectedExitDate.value);
      const now = new Date();
      if (exitDate <= now) {
        console.warn('🚪 Exit date should be in the future');
        validated.expectedExitDate.confidence *= 0.7;
      }
    }
    
    return validated;
  }

  /**
   * Enrich with calculated metrics
   */
  enrichWithCalculations(data) {
    const enriched = { ...data };
    
    // Calculate holding period from current date to exit date
    if (enriched.expectedExitDate?.value && !enriched.holdingPeriod?.value) {
      const exitDate = new Date(enriched.expectedExitDate.value);
      const now = new Date();
      const yearsDiff = (exitDate - now) / (1000 * 60 * 60 * 24 * 365.25);
      
      if (yearsDiff > 0) {
        enriched.holdingPeriod = {
          value: Math.round(yearsDiff * 10) / 10, // Round to 1 decimal
          confidence: enriched.expectedExitDate.confidence,
          source: 'calculated'
        };
      }
    }
    
    // Calculate implied exit value from IRR and holding period
    // This would require more complex calculations with cash flows
    
    // Add exit strategy summary
    if (enriched.exitRoute?.value || enriched.targetIRR?.value || enriched.holdingPeriod?.value) {
      enriched.exitStrategy = {
        summary: this.createExitSummary(enriched),
        confidence: 0.6,
        source: 'calculated'
      };
    }
    
    return enriched;
  }

  /**
   * Create exit strategy summary
   */
  createExitSummary(data) {
    const parts = [];
    
    if (data.exitRoute?.value) {
      const routeLabel = {
        'IPO': 'IPO',
        'trade_sale': 'Trade Sale',
        'secondary_buyout': 'Secondary Buyout',
        'management_buyout': 'Management Buyout',
        'refinancing': 'Refinancing'
      };
      parts.push(`Exit via ${routeLabel[data.exitRoute.value] || data.exitRoute.value}`);
    }
    
    if (data.holdingPeriod?.value) {
      parts.push(`${data.holdingPeriod.value} year holding period`);
    }
    
    if (data.targetIRR?.value) {
      parts.push(`Target IRR: ${data.targetIRR.value}%`);
    }
    
    return parts.join(', ');
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
      
      // Boost confidence for consistent exit strategy
      if (field === 'exitRoute' && data.targetIRR?.value && data.holdingPeriod?.value) {
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
    let searchTerms = [String(value).toLowerCase()];
    
    // Add field-specific search terms
    if (field === 'exitRoute') {
      const routeTerms = {
        'IPO': ['ipo', 'public offering', 'listing'],
        'trade_sale': ['trade sale', 'strategic sale', 'acquisition'],
        'secondary_buyout': ['secondary', 'buyout', 'private equity sale']
      };
      searchTerms = routeTerms[value] || searchTerms;
    }
    
    return files.filter(file => 
      file.content && searchTerms.some(term => 
        file.content.toLowerCase().includes(term)
      )
    ).length;
  }

  /**
   * Check if value is realistic for the field
   */
  isRealisticValue(field, value) {
    const realistic = {
      disposalCost: (v) => v >= 0.5 && v <= 10, // 0.5% to 10%
      terminalCapRate: (v) => v >= 4 && v <= 15, // 4% to 15%
      exitMultiple: (v) => v >= 1 && v <= 50, // 1x to 50x
      terminalGrowthRate: (v) => v >= 0 && v <= 10, // 0% to 10%
      discountRate: (v) => v >= 5 && v <= 25, // 5% to 25%
      targetIRR: (v) => v >= 10 && v <= 40, // 10% to 40%
      holdingPeriod: (v) => v >= 1 && v <= 15, // 1 to 15 years
      exitValue: (v) => v >= 1000000 && v <= 1000000000000 // $1M to $1T
    };
    
    const validator = realistic[field];
    return validator ? validator(value) : true;
  }

  /**
   * Get intelligent defaults
   */
  getIntelligentDefaults(files) {
    console.log('🚪 Using intelligent defaults for exit assumptions');
    
    return {
      disposalCost: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      terminalCapRate: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      exitMultiple: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      exitMultipleType: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      targetIRR: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      expectedExitDate: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      exitRoute: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      holdingPeriod: {
        value: null,
        confidence: 0,
        source: 'not_found'
      }
    };
  }

  /**
   * Apply extracted exit assumptions to form
   */
  async applyToForm(extractedData) {
    console.log('🚪 Applying exit assumptions to form');
    
    return await this.mappingEngine.applyDataToForm(extractedData, {
      section: 'exitAssumptions',
      showConfidence: true,
      animateChanges: true
    });
  }
}

// Export for use
window.ExitAssumptionsExtractor = ExitAssumptionsExtractor;