/**
 * HighLevelParametersExtractor.js - Extract high-level model parameters
 * Handles: Currency, dates, reporting periods, and project timeline
 */

class HighLevelParametersExtractor {
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
    console.log('✅ HighLevelParametersExtractor initialized');
  }

  /**
   * Extract high-level parameters from documents
   */
  async extract(files) {
    console.log('🎯 Extracting high-level parameters from', files.length, 'files');
    
    try {
      // Step 1: Use AI to extract raw data
      const aiExtraction = await this.extractWithAI(files);
      
      // Step 2: Apply intelligent parsing fallbacks
      const enhancedData = this.enhanceWithParsing(aiExtraction, files);
      
      // Step 3: Validate and score confidence
      const scoredData = this.scoreConfidence(enhancedData, files);
      
      // Step 4: Standardize the data
      const standardized = await this.standardizer.standardize(scoredData);
      
      console.log('🎯 High-level parameters extraction complete:', standardized);
      return standardized;
      
    } catch (error) {
      console.error('🎯 Error extracting high-level parameters:', error);
      // Return intelligent defaults based on file content
      return this.getIntelligentDefaults(files);
    }
  }

  /**
   * Use AI service to extract parameters
   */
  async extractWithAI(files) {
    const prompt = `Extract high-level financial model parameters from these documents.

Focus on:
1. CURRENCY: Transaction currency (USD, EUR, GBP, etc.)
2. DATES: Project/acquisition start date, end/exit date, closing date
3. PERIODS: Reporting frequency (daily, monthly, quarterly, yearly)
4. TIMELINE: Investment horizon, holding period

Look for explicit mentions and infer from context when needed.
For dates, look for "closing", "acquisition", "start", "exit", "maturity" keywords.
For currency, check monetary values, currency symbols, or explicit mentions.

Return ONLY these fields with actual values found or null:
{
  "currency": "USD|EUR|GBP|etc or null",
  "projectStartDate": "date string or null",
  "projectEndDate": "date string or null",
  "modelPeriods": "daily|monthly|quarterly|yearly or null"
}`;

    try {
      const extraction = await this.extractionService.extractFromDocuments(
        files,
        'highLevelParameters'
      );
      
      return extraction;
    } catch (error) {
      console.error('AI extraction failed:', error);
      return {};
    }
  }

  /**
   * Enhance AI extraction with intelligent parsing
   */
  enhanceWithParsing(aiData, files) {
    const enhanced = { ...aiData };
    
    // Combine all file contents
    const allContent = files
      .map(f => f.content || '')
      .join('\n')
      .toLowerCase();
    
    // Extract currency if not found by AI
    if (!enhanced.currency?.value) {
      const currency = this.extractCurrency(allContent);
      if (currency) {
        enhanced.currency = {
          value: currency.value,
          confidence: currency.confidence,
          source: currency.source
        };
      }
    }
    
    // Extract dates if not found by AI
    if (!enhanced.projectStartDate?.value || !enhanced.projectEndDate?.value) {
      const dates = this.extractDates(allContent, files);
      
      if (!enhanced.projectStartDate?.value && dates.startDate) {
        enhanced.projectStartDate = dates.startDate;
      }
      
      if (!enhanced.projectEndDate?.value && dates.endDate) {
        enhanced.projectEndDate = dates.endDate;
      }
    }
    
    // Extract reporting period if not found
    if (!enhanced.modelPeriods?.value) {
      const period = this.extractReportingPeriod(allContent);
      if (period) {
        enhanced.modelPeriods = period;
      }
    }
    
    return enhanced;
  }

  /**
   * Extract currency from text
   */
  extractCurrency(text) {
    // Currency patterns with context
    const patterns = [
      // Explicit currency mentions
      { regex: /(?:currency|denomination|in)\s*:?\s*(USD|EUR|GBP|JPY|CAD|AUD|CHF|CNY)/i, weight: 1.0 },
      { regex: /(?:all amounts? (?:are )?in|expressed in)\s*(USD|EUR|GBP|JPY|CAD|AUD|CHF|CNY)/i, weight: 0.9 },
      
      // Currency symbols in amounts
      { regex: /\$[\d,]+(?:\.\d+)?(?:\s*(?:million|billion|m|b))?/i, currency: 'USD', weight: 0.7 },
      { regex: /€[\d,]+(?:\.\d+)?(?:\s*(?:million|billion|m|b))?/i, currency: 'EUR', weight: 0.8 },
      { regex: /£[\d,]+(?:\.\d+)?(?:\s*(?:million|billion|m|b))?/i, currency: 'GBP', weight: 0.8 },
      { regex: /¥[\d,]+(?:\.\d+)?(?:\s*(?:million|billion|m|b))?/i, currency: 'JPY', weight: 0.8 },
      
      // Currency codes in context
      { regex: /[\d,]+(?:\.\d+)?\s*(USD|EUR|GBP|JPY|CAD|AUD|CHF|CNY)(?:\s*(?:million|billion|m|b))?/i, weight: 0.8 }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const pattern of patterns) {
      const match = text.match(pattern.regex);
      if (match) {
        const currency = pattern.currency || match[1];
        if (pattern.weight > highestWeight) {
          highestWeight = pattern.weight;
          bestMatch = {
            value: currency.toUpperCase(),
            confidence: pattern.weight,
            source: 'pattern_matching'
          };
        }
      }
    }
    
    return bestMatch;
  }

  /**
   * Extract dates from text
   */
  extractDates(text, files) {
    const dates = {
      startDate: null,
      endDate: null
    };
    
    // Date patterns with context keywords
    const datePatterns = [
      // ISO format: YYYY-MM-DD
      /(\d{4}-\d{2}-\d{2})/g,
      // US format: MM/DD/YYYY
      /(\d{1,2}\/\d{1,2}\/\d{4})/g,
      // European format: DD.MM.YYYY
      /(\d{1,2}\.\d{1,2}\.\d{4})/g,
      // Long format: January 15, 2024
      /((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4})/gi,
      // Alternative: 15 January 2024
      /(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{4})/gi
    ];
    
    // Keywords for start dates
    const startKeywords = [
      'closing date', 'acquisition date', 'transaction date', 'completion date',
      'effective date', 'start date', 'commencement', 'beginning'
    ];
    
    // Keywords for end dates
    const endKeywords = [
      'exit date', 'maturity date', 'end date', 'termination date',
      'expected exit', 'investment horizon', 'holding period ends'
    ];
    
    // Extract all dates with context
    const foundDates = [];
    
    for (const pattern of datePatterns) {
      let match;
      pattern.lastIndex = 0; // Reset regex
      
      while ((match = pattern.exec(text)) !== null) {
        const dateStr = match[1];
        const position = match.index;
        
        // Get surrounding context (100 chars before and after)
        const contextStart = Math.max(0, position - 100);
        const contextEnd = Math.min(text.length, position + dateStr.length + 100);
        const context = text.substring(contextStart, contextEnd).toLowerCase();
        
        // Determine if this is a start or end date based on context
        let isStartDate = startKeywords.some(keyword => context.includes(keyword));
        let isEndDate = endKeywords.some(keyword => context.includes(keyword));
        
        foundDates.push({
          date: dateStr,
          isStartDate,
          isEndDate,
          context: context.substring(0, 50) + '...',
          confidence: (isStartDate || isEndDate) ? 0.8 : 0.5
        });
      }
    }
    
    // Select best dates
    const startDates = foundDates.filter(d => d.isStartDate);
    const endDates = foundDates.filter(d => d.isEndDate);
    
    if (startDates.length > 0) {
      const best = startDates.reduce((prev, curr) => 
        curr.confidence > prev.confidence ? curr : prev
      );
      dates.startDate = {
        value: best.date,
        confidence: best.confidence,
        source: 'context_matching'
      };
    }
    
    if (endDates.length > 0) {
      const best = endDates.reduce((prev, curr) => 
        curr.confidence > prev.confidence ? curr : prev
      );
      dates.endDate = {
        value: best.date,
        confidence: best.confidence,
        source: 'context_matching'
      };
    }
    
    // If no contextual dates found, try to infer from all dates
    if (!dates.startDate && !dates.endDate && foundDates.length >= 2) {
      // Sort dates chronologically
      const sortedDates = foundDates
        .map(d => ({ ...d, parsed: this.parseDate(d.date) }))
        .filter(d => d.parsed)
        .sort((a, b) => a.parsed - b.parsed);
      
      if (sortedDates.length >= 2) {
        dates.startDate = {
          value: sortedDates[0].date,
          confidence: 0.5,
          source: 'inferred_earliest'
        };
        dates.endDate = {
          value: sortedDates[sortedDates.length - 1].date,
          confidence: 0.5,
          source: 'inferred_latest'
        };
      }
    }
    
    return dates;
  }

  /**
   * Parse date string to Date object
   */
  parseDate(dateStr) {
    // Try different parsing strategies
    const parsed = new Date(dateStr);
    if (!isNaN(parsed.getTime())) return parsed;
    
    // Handle DD/MM/YYYY or MM/DD/YYYY
    const parts = dateStr.match(/(\d{1,2})[\/\.](\d{1,2})[\/\.](\d{4})/);
    if (parts) {
      // Try US format first (MM/DD/YYYY)
      let date = new Date(parts[3], parts[1] - 1, parts[2]);
      if (date.getMonth() === parts[1] - 1) return date;
      
      // Try European format (DD/MM/YYYY)
      date = new Date(parts[3], parts[2] - 1, parts[1]);
      if (date.getMonth() === parts[2] - 1) return date;
    }
    
    return null;
  }

  /**
   * Extract reporting period
   */
  extractReportingPeriod(text) {
    const patterns = [
      { regex: /report(?:ing|s)?\s+(?:on\s+)?(?:a\s+)?(daily|monthly|quarterly|annual|yearly)\s+basis/i, weight: 1.0 },
      { regex: /(daily|monthly|quarterly|annual|yearly)\s+(?:financial\s+)?(?:report|statement|projection)/i, weight: 0.9 },
      { regex: /(?:period|frequency|interval)s?\s*:?\s*(daily|monthly|quarterly|annual|yearly)/i, weight: 0.8 },
      { regex: /\b(daily|monthly|quarterly|annual|yearly)\s+(?:revenue|expense|cash\s*flow)/i, weight: 0.7 }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const pattern of patterns) {
      const match = text.match(pattern.regex);
      if (match) {
        let period = match[1].toLowerCase();
        // Normalize annual to yearly
        if (period === 'annual') period = 'yearly';
        
        if (pattern.weight > highestWeight) {
          highestWeight = pattern.weight;
          bestMatch = {
            value: period,
            confidence: pattern.weight,
            source: 'pattern_matching'
          };
        }
      }
    }
    
    return bestMatch;
  }

  /**
   * Score confidence based on multiple factors
   */
  scoreConfidence(data, files) {
    const scored = {};
    
    for (const [field, value] of Object.entries(data)) {
      if (!value || !value.value) {
        scored[field] = value;
        continue;
      }
      
      let confidence = value.confidence || 0.5;
      
      // Boost confidence if value appears in multiple files
      const occurrences = this.countOccurrences(value.value, files);
      if (occurrences > 1) {
        confidence = Math.min(confidence + 0.1 * (occurrences - 1), 1.0);
      }
      
      // Boost confidence for well-formatted values
      if (field === 'currency' && ['USD', 'EUR', 'GBP'].includes(value.value)) {
        confidence = Math.min(confidence + 0.1, 1.0);
      }
      
      if (field.includes('Date') && this.isValidDate(value.value)) {
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
   * Count occurrences across files
   */
  countOccurrences(value, files) {
    const searchValue = String(value).toLowerCase();
    return files.filter(file => 
      file.content && file.content.toLowerCase().includes(searchValue)
    ).length;
  }

  /**
   * Validate date format
   */
  isValidDate(dateStr) {
    const date = this.parseDate(dateStr);
    if (!date) return false;
    
    const year = date.getFullYear();
    // Check reasonable date range (2000-2050)
    return year >= 2000 && year <= 2050;
  }

  /**
   * Get intelligent defaults based on file content
   */
  getIntelligentDefaults(files) {
    console.log('🎯 Using intelligent defaults based on file analysis');
    
    // Analyze file names and content for clues
    const currentYear = new Date().getFullYear();
    const currentMonth = new Date().getMonth() + 1;
    
    return {
      currency: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      projectStartDate: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      projectEndDate: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      modelPeriods: {
        value: null,
        confidence: 0,
        source: 'not_found'
      }
    };
  }

  /**
   * Apply extracted parameters to form
   */
  async applyToForm(extractedData) {
    console.log('🎯 Applying high-level parameters to form');
    
    return await this.mappingEngine.applyDataToForm(extractedData, {
      section: 'highLevelParameters',
      showConfidence: true,
      animateChanges: true
    });
  }
}

// Export for use
window.HighLevelParametersExtractor = HighLevelParametersExtractor;