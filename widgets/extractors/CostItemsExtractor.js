/**
 * CostItemsExtractor.js - Extract operating and capital expenses
 * Handles: OpEx items, CapEx items, cost inflation rates, and expense projections
 */

class CostItemsExtractor {
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
    console.log('✅ CostItemsExtractor initialized');
  }

  /**
   * Extract cost items from documents
   */
  async extract(files) {
    console.log('💸 Extracting cost items from', files.length, 'files');
    
    try {
      // Step 1: Use AI to extract cost data
      const aiExtraction = await this.extractWithAI(files);
      
      // Step 2: Enhance with pattern matching
      const enhancedData = this.enhanceWithParsing(aiExtraction, files);
      
      // Step 3: Classify and validate cost items
      const classifiedData = this.classifyAndValidate(enhancedData);
      
      // Step 4: Calculate cost metrics
      const enrichedData = this.enrichWithMetrics(classifiedData);
      
      // Step 5: Score confidence
      const scoredData = this.scoreConfidence(enrichedData, files);
      
      // Step 6: Standardize the data
      const standardized = await this.standardizer.standardize(scoredData);
      
      console.log('💸 Cost items extraction complete:', standardized);
      return standardized;
      
    } catch (error) {
      console.error('💸 Error extracting cost items:', error);
      return this.getIntelligentDefaults(files);
    }
  }

  /**
   * Use AI service to extract cost items
   */
  async extractWithAI(files) {
    const prompt = `Extract operating expenses and capital expenditures from these financial documents.

Focus on identifying:

1. OPERATING EXPENSES (OpEx):
   - Personnel costs (salaries, benefits, bonuses)
   - Office expenses (rent, utilities, insurance)
   - Marketing and advertising costs
   - Professional services (legal, accounting, consulting)
   - Technology expenses (software, IT services)
   - Travel and entertainment
   - General administrative expenses

2. CAPITAL EXPENDITURES (CapEx):
   - Equipment purchases
   - Technology investments
   - Property improvements
   - Vehicle purchases
   - Major software implementations
   - Infrastructure investments

3. COST TRENDS:
   - Historical cost data
   - Cost inflation rates
   - Growth projections
   - Cost optimization initiatives

Look for expense breakdowns, cost centers, and budget line items.
Extract actual values only - do not estimate or assume.

Return ONLY this structure with actual values found or null:
{
  "operatingExpenses": [
    {
      "name": "expense category name",
      "value": numeric_annual_value_or_null,
      "growthType": "linear|compound|custom|null",
      "growthRate": percentage_as_number_or_null,
      "category": "personnel|office|marketing|professional|technology|other",
      "isFixed": true_or_false_or_null
    }
  ],
  "capitalExpenses": [
    {
      "name": "capex item name", 
      "value": numeric_value_or_null,
      "growthType": "linear|compound|custom|null",
      "growthRate": percentage_as_number_or_null,
      "category": "equipment|technology|property|vehicles|other",
      "depreciationYears": numeric_years_or_null
    }
  ],
  "totalOpEx": numeric_total_or_null,
  "totalCapEx": numeric_total_or_null,
  "costInflationRate": percentage_as_number_or_null,
  "costCurrency": "USD|EUR|GBP|etc or null"
}`;

    try {
      const extraction = await this.extractionService.extractFromDocuments(
        files,
        'costs'
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
    
    // Extract operating expenses if not found
    if (!enhanced.operatingExpenses?.length) {
      const opEx = this.extractOperatingExpenses(allContent);
      if (opEx.length > 0) {
        enhanced.operatingExpenses = opEx;
      }
    }
    
    // Extract capital expenses if not found
    if (!enhanced.capitalExpenses?.length) {
      const capEx = this.extractCapitalExpenses(allContent);
      if (capEx.length > 0) {
        enhanced.capitalExpenses = capEx;
      }
    }
    
    // Extract totals if not found
    if (!enhanced.totalOpEx) {
      const totalOpEx = this.extractTotalOperatingExpenses(allContent);
      if (totalOpEx) {
        enhanced.totalOpEx = totalOpEx;
      }
    }
    
    if (!enhanced.totalCapEx) {
      const totalCapEx = this.extractTotalCapitalExpenses(allContent);
      if (totalCapEx) {
        enhanced.totalCapEx = totalCapEx;
      }
    }
    
    // Extract cost inflation rate
    if (!enhanced.costInflationRate) {
      const inflationRate = this.extractCostInflationRate(allContent);
      if (inflationRate) {
        enhanced.costInflationRate = inflationRate;
      }
    }
    
    return enhanced;
  }

  /**
   * Extract operating expenses using pattern matching
   */
  extractOperatingExpenses(text) {
    const expenses = [];
    
    // Common operating expense categories
    const opExCategories = {
      'personnel': ['salaries', 'wages', 'benefits', 'bonuses', 'payroll', 'staff costs', 'employee costs'],
      'office': ['rent', 'utilities', 'office expenses', 'facilities', 'insurance', 'office space'],
      'marketing': ['marketing', 'advertising', 'promotion', 'sales expense', 'customer acquisition'],
      'professional': ['legal fees', 'accounting', 'consulting', 'professional services', 'audit fees'],
      'technology': ['software', 'IT services', 'cloud costs', 'technology expenses', 'software licenses'],
      'other': ['travel', 'entertainment', 'general administrative', 'miscellaneous', 'other expenses']
    };
    
    // Pattern 1: Direct expense mentions
    const expensePatterns = [
      // [Category] expenses: $X
      /([A-Za-z\s&]+)\s+(?:expenses?|costs?):\s*(?:\$|USD)?([0-9,]+(?:\.[0-9]+)?)\s*(million|billion|m|b|k|thousand)?/gi,
      
      // $X for [category]
      /(?:\$|USD)?([0-9,]+(?:\.[0-9]+)?)\s*(million|billion|m|b|k|thousand)?\s+(?:for|on)\s+([A-Za-z\s&]+)/gi,
      
      // Table format: Category | Amount
      /([A-Za-z\s&]+)\s*[\|,]\s*(?:\$|USD)?([0-9,]+(?:\.[0-9]+)?)\s*(million|billion|m|b|k|thousand)?/gi
    ];
    
    for (const pattern of expensePatterns) {
      let match;
      pattern.lastIndex = 0;
      
      while ((match = pattern.exec(text)) !== null) {
        let name, valueStr, unit;
        
        if (pattern.source.includes('for|on')) {
          // $X for [category] format
          valueStr = match[1];
          unit = match[2];
          name = match[3];
        } else {
          // [Category]: $X format
          name = match[1];
          valueStr = match[2];
          unit = match[3];
        }
        
        name = name.trim();
        let value = parseFloat(valueStr.replace(/,/g, ''));
        unit = (unit || '').toLowerCase();
        
        // Apply multipliers
        if (unit === 'billion' || unit === 'b') {
          value *= 1000000000;
        } else if (unit === 'million' || unit === 'm') {
          value *= 1000000;
        } else if (unit === 'thousand' || unit === 'k') {
          value *= 1000;
        }
        
        // Validate and categorize
        if (value >= 1000 && value <= 10000000000 && name.length > 2 && name.length < 100) {
          const category = this.categorizeExpense(name, opExCategories);
          const isFixed = this.isFixedCost(name);
          
          expenses.push({
            name: name,
            value: value,
            growthType: 'linear',
            growthRate: null,
            category: category,
            isFixed: isFixed,
            source: 'pattern_matching',
            confidence: 0.7
          });
        }
      }
    }
    
    // Pattern 2: Specific expense extractions
    expenses.push(...this.extractSpecificExpenses(text));
    
    return this.deduplicateExpenses(expenses);
  }

  /**
   * Extract capital expenses using pattern matching
   */
  extractCapitalExpenses(text) {
    const expenses = [];
    
    // Common CapEx categories
    const capExKeywords = [
      'equipment purchase', 'machinery', 'computer equipment', 'software implementation',
      'property improvement', 'building renovation', 'vehicle purchase', 'technology investment',
      'infrastructure', 'capital investment', 'fixed assets', 'plant and equipment'
    ];
    
    // CapEx patterns
    const capExPatterns = [
      // Capital expenditure on [item]: $X
      /capital\\s+expenditure\\s+on\\s+([A-Za-z\\s&]+):\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b|k|thousand)?/gi,
      
      // CapEx: [item] $X
      /capex\\s*:?\\s*([A-Za-z\\s&]+)\\s+(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b|k|thousand)?/gi,
      
      // Investment in [item]: $X
      /investment\\s+in\\s+([A-Za-z\\s&]+):\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b|k|thousand)?/gi
    ];
    
    for (const pattern of capExPatterns) {
      let match;
      pattern.lastIndex = 0;
      
      while ((match = pattern.exec(text)) !== null) {
        const name = match[1].trim();
        let value = parseFloat(match[2].replace(/,/g, ''));
        const unit = (match[3] || '').toLowerCase();
        
        // Apply multipliers
        if (unit === 'billion' || unit === 'b') {
          value *= 1000000000;
        } else if (unit === 'million' || unit === 'm') {
          value *= 1000000;
        } else if (unit === 'thousand' || unit === 'k') {
          value *= 1000;
        }
        
        if (value >= 10000 && value <= 100000000000 && name.length > 2) {
          expenses.push({
            name: name,
            value: value,
            growthType: 'custom',
            growthRate: null,
            category: this.categorizeCapEx(name),
            depreciationYears: this.estimateDepreciationYears(name),
            source: 'pattern_matching',
            confidence: 0.7
          });
        }
      }
    }
    
    // Look for CapEx keywords
    for (const keyword of capExKeywords) {
      const patterns = [
        new RegExp(`${keyword}\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b|k|thousand)?`, 'gi'),
        new RegExp(`(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b|k|thousand)?\\s+for\\s+${keyword}`, 'gi')
      ];
      
      for (const pattern of patterns) {
        const match = text.match(pattern);
        if (match) {
          let value = parseFloat(match[1].replace(/,/g, ''));
          const unit = (match[2] || '').toLowerCase();
          
          if (unit === 'billion' || unit === 'b') {
            value *= 1000000000;
          } else if (unit === 'million' || unit === 'm') {
            value *= 1000000;
          } else if (unit === 'thousand' || unit === 'k') {
            value *= 1000;
          }
          
          if (value >= 10000) {
            expenses.push({
              name: keyword,
              value: value,
              growthType: 'custom',
              growthRate: null,
              category: this.categorizeCapEx(keyword),
              depreciationYears: this.estimateDepreciationYears(keyword),
              source: 'keyword_pattern',
              confidence: 0.6
            });
          }
        }
      }
    }
    
    return this.deduplicateExpenses(expenses);
  }

  /**
   * Extract specific expense types
   */
  extractSpecificExpenses(text) {
    const expenses = [];
    
    // Salary/Personnel patterns
    const salaryPatterns = [
      /(?:total\\s+)?(?:salary|salaries|payroll)\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b|k|thousand)?/gi,
      /employee\\s+costs?\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b|k|thousand)?/gi
    ];
    
    for (const pattern of salaryPatterns) {
      const match = text.match(pattern);
      if (match) {
        let value = parseFloat(match[1].replace(/,/g, ''));
        const unit = (match[2] || '').toLowerCase();
        
        if (unit === 'billion' || unit === 'b') {
          value *= 1000000000;
        } else if (unit === 'million' || unit === 'm') {
          value *= 1000000;
        } else if (unit === 'thousand' || unit === 'k') {
          value *= 1000;
        }
        
        if (value >= 10000) {
          expenses.push({
            name: 'Personnel Costs',
            value: value,
            growthType: 'linear',
            growthRate: null,
            category: 'personnel',
            isFixed: false,
            source: 'specific_pattern',
            confidence: 0.8
          });
        }
      }
    }
    
    return expenses;
  }

  /**
   * Extract total operating expenses
   */
  extractTotalOperatingExpenses(text) {
    const patterns = [
      /total\\s+operating\\s+expenses?\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?/gi,
      /operating\\s+costs?\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?/gi,
      /opex\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?/gi
    ];
    
    for (const pattern of patterns) {
      const match = text.match(pattern);
      if (match) {
        let value = parseFloat(match[1].replace(/,/g, ''));
        const unit = (match[2] || '').toLowerCase();
        
        if (unit === 'billion' || unit === 'b') {
          value *= 1000000000;
        } else if (unit === 'million' || unit === 'm') {
          value *= 1000000;
        }
        
        if (value >= 10000) {
          return {
            value: value,
            confidence: 0.8,
            source: 'pattern_matching'
          };
        }
      }
    }
    
    return null;
  }

  /**
   * Extract total capital expenses
   */
  extractTotalCapitalExpenses(text) {
    const patterns = [
      /total\\s+capital\\s+expenditures?\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?/gi,
      /total\\s+capex\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?/gi,
      /capital\\s+investments?\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?/gi
    ];
    
    for (const pattern of patterns) {
      const match = text.match(pattern);
      if (match) {
        let value = parseFloat(match[1].replace(/,/g, ''));
        const unit = (match[2] || '').toLowerCase();
        
        if (unit === 'billion' || unit === 'b') {
          value *= 1000000000;
        } else if (unit === 'million' || unit === 'm') {
          value *= 1000000;
        }
        
        if (value >= 10000) {
          return {
            value: value,
            confidence: 0.8,
            source: 'pattern_matching'
          };
        }
      }
    }
    
    return null;
  }

  /**
   * Extract cost inflation rate
   */
  extractCostInflationRate(text) {
    const patterns = [
      /cost\\s+inflation\\s*:?\\s*([0-9.]+)\\s*%/gi,
      /expense\\s+growth\\s*:?\\s*([0-9.]+)\\s*%/gi,
      /([0-9.]+)\\s*%\\s+cost\\s+increase/gi
    ];
    
    for (const pattern of patterns) {
      const match = text.match(pattern);
      if (match) {
        const rate = parseFloat(match[1]);
        
        if (rate >= 0 && rate <= 50) {
          return {
            value: rate,
            confidence: 0.7,
            source: 'pattern_matching'
          };
        }
      }
    }
    
    return null;
  }

  /**
   * Categorize expense type
   */
  categorizeExpense(name, categories) {
    const nameL = name.toLowerCase();
    
    for (const [category, keywords] of Object.entries(categories)) {
      if (keywords.some(keyword => nameL.includes(keyword))) {
        return category;
      }
    }
    
    return 'other';
  }

  /**
   * Categorize CapEx type
   */
  categorizeCapEx(name) {
    const nameL = name.toLowerCase();
    
    if (nameL.includes('equipment') || nameL.includes('machinery') || nameL.includes('tools')) {
      return 'equipment';
    }
    if (nameL.includes('technology') || nameL.includes('software') || nameL.includes('computer') || nameL.includes('it')) {
      return 'technology';
    }
    if (nameL.includes('property') || nameL.includes('building') || nameL.includes('renovation') || nameL.includes('facility')) {
      return 'property';
    }
    if (nameL.includes('vehicle') || nameL.includes('car') || nameL.includes('truck') || nameL.includes('fleet')) {
      return 'vehicles';
    }
    
    return 'other';
  }

  /**
   * Determine if cost is fixed
   */
  isFixedCost(name) {
    const nameL = name.toLowerCase();
    const fixedKeywords = ['rent', 'insurance', 'license', 'subscription', 'fixed'];
    const variableKeywords = ['commission', 'bonus', 'variable', 'travel', 'marketing'];
    
    if (fixedKeywords.some(keyword => nameL.includes(keyword))) {
      return true;
    }
    if (variableKeywords.some(keyword => nameL.includes(keyword))) {
      return false;
    }
    
    return null; // Unknown
  }

  /**
   * Estimate depreciation years for CapEx
   */
  estimateDepreciationYears(name) {
    const nameL = name.toLowerCase();
    
    if (nameL.includes('computer') || nameL.includes('software') || nameL.includes('technology')) {
      return 3;
    }
    if (nameL.includes('equipment') || nameL.includes('machinery') || nameL.includes('vehicle')) {
      return 5;
    }
    if (nameL.includes('building') || nameL.includes('property') || nameL.includes('facility')) {
      return 10;
    }
    
    return 5; // Default
  }

  /**
   * Remove duplicate expenses
   */
  deduplicateExpenses(expenses) {
    const unique = [];
    const seen = new Set();
    
    for (const expense of expenses) {
      const key = `${expense.name.toLowerCase()}_${expense.value}`;
      if (!seen.has(key)) {
        seen.add(key);
        unique.push(expense);
      }
    }
    
    return unique.sort((a, b) => b.value - a.value);
  }

  /**
   * Classify and validate expense items
   */
  classifyAndValidate(data) {
    const validated = { ...data };
    
    // Validate operating expenses
    if (validated.operatingExpenses) {
      validated.operatingExpenses = validated.operatingExpenses.filter(expense => {
        if (!expense.name || !expense.value || expense.value <= 0) {
          return false;
        }
        
        if (expense.value < 100 || expense.value > 50000000000) {
          return false;
        }
        
        if (expense.growthRate && (expense.growthRate < -50 || expense.growthRate > 100)) {
          expense.growthRate = null;
        }
        
        return true;
      });
    }
    
    // Validate capital expenses
    if (validated.capitalExpenses) {
      validated.capitalExpenses = validated.capitalExpenses.filter(expense => {
        if (!expense.name || !expense.value || expense.value <= 0) {
          return false;
        }
        
        if (expense.value < 1000 || expense.value > 100000000000) {
          return false;
        }
        
        return true;
      });
    }
    
    return validated;
  }

  /**
   * Enrich with calculated metrics
   */
  enrichWithMetrics(data) {
    const enriched = { ...data };
    
    // Calculate totals from items if not provided
    if (enriched.operatingExpenses?.length > 0 && !enriched.totalOpEx) {
      const total = enriched.operatingExpenses.reduce((sum, expense) => sum + expense.value, 0);
      enriched.totalOpEx = {
        value: total,
        confidence: 0.6,
        source: 'calculated'
      };
    }
    
    if (enriched.capitalExpenses?.length > 0 && !enriched.totalCapEx) {
      const total = enriched.capitalExpenses.reduce((sum, expense) => sum + expense.value, 0);
      enriched.totalCapEx = {
        value: total,
        confidence: 0.6,
        source: 'calculated'
      };
    }
    
    // Calculate cost breakdown by category
    if (enriched.operatingExpenses?.length > 0) {
      enriched.opExBreakdown = this.calculateCostBreakdown(enriched.operatingExpenses);
    }
    
    if (enriched.capitalExpenses?.length > 0) {
      enriched.capExBreakdown = this.calculateCostBreakdown(enriched.capitalExpenses);
    }
    
    return enriched;
  }

  /**
   * Calculate cost breakdown by category
   */
  calculateCostBreakdown(expenses) {
    const breakdown = {};
    const total = expenses.reduce((sum, expense) => sum + expense.value, 0);
    
    for (const expense of expenses) {
      const category = expense.category || 'other';
      if (!breakdown[category]) {
        breakdown[category] = { value: 0, percentage: 0, items: [] };
      }
      breakdown[category].value += expense.value;
      breakdown[category].items.push(expense.name);
    }
    
    // Calculate percentages
    for (const category of Object.keys(breakdown)) {
      breakdown[category].percentage = (breakdown[category].value / total) * 100;
    }
    
    return breakdown;
  }

  /**
   * Score confidence based on validation
   */
  scoreConfidence(data, files) {
    const scored = {};
    
    for (const [field, value] of Object.entries(data)) {
      if (!value || (typeof value === 'object' && value.value === null)) {
        scored[field] = value;
        continue;
      }
      
      if ((field === 'operatingExpenses' || field === 'capitalExpenses') && Array.isArray(value)) {
        scored[field] = value.map(item => ({
          ...item,
          confidence: this.scoreItemConfidence(item, files)
        }));
      } else if (typeof value === 'object' && value.value !== undefined) {
        let confidence = value.confidence || 0.5;
        
        const occurrences = this.countOccurrences(value.value, files);
        if (occurrences > 1) {
          confidence = Math.min(confidence + 0.1 * (occurrences - 1), 1.0);
        }
        
        scored[field] = {
          ...value,
          confidence: confidence
        };
      } else {
        scored[field] = value;
      }
    }
    
    return scored;
  }

  /**
   * Score individual item confidence
   */
  scoreItemConfidence(item, files) {
    let confidence = item.confidence || 0.5;
    
    // Boost for realistic values
    if (item.value >= 1000 && item.value <= 1000000000) {
      confidence += 0.1;
    }
    
    // Boost for complete categorization
    if (item.category && item.category !== 'other') {
      confidence += 0.1;
    }
    
    // Boost for CapEx with depreciation info
    if (item.depreciationYears) {
      confidence += 0.1;
    }
    
    return Math.min(confidence, 1.0);
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
   * Get intelligent defaults
   */
  getIntelligentDefaults(files) {
    console.log('💸 Using intelligent defaults for cost items');
    
    return {
      operatingExpenses: [],
      capitalExpenses: [],
      totalOpEx: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      totalCapEx: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      costInflationRate: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      costCurrency: {
        value: null,
        confidence: 0,
        source: 'not_found'
      }
    };
  }

  /**
   * Apply extracted cost items to form
   */
  async applyToForm(extractedData) {
    console.log('💸 Applying cost items to form');
    
    return await this.mappingEngine.applyDataToForm(extractedData, {
      section: 'costItems',
      showConfidence: true,
      animateChanges: true
    });
  }
}

// Export for use
window.CostItemsExtractor = CostItemsExtractor;