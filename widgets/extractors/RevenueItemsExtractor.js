/**
 * RevenueItemsExtractor.js - Extract revenue streams, projections, and growth rates
 * Handles: Revenue line items, sales categories, historical data, and growth projections
 */

class RevenueItemsExtractor {
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
    console.log('✅ RevenueItemsExtractor initialized');
  }

  /**
   * Extract revenue items from documents
   */
  async extract(files) {
    console.log('💰 Extracting revenue items from', files.length, 'files');
    
    try {
      // Step 1: Use AI to extract revenue data
      const aiExtraction = await this.extractWithAI(files);
      
      // Step 2: Enhance with pattern matching
      const enhancedData = this.enhanceWithParsing(aiExtraction, files);
      
      // Step 3: Validate and normalize revenue items
      const validatedData = this.validateRevenueItems(enhancedData);
      
      // Step 4: Calculate derived metrics
      const enrichedData = this.enrichWithMetrics(validatedData);
      
      // Step 5: Score confidence
      const scoredData = this.scoreConfidence(enrichedData, files);
      
      // Step 6: Standardize the data
      const standardized = await this.standardizer.standardize(scoredData);
      
      console.log('💰 Revenue items extraction complete:', standardized);
      return standardized;
      
    } catch (error) {
      console.error('💰 Error extracting revenue items:', error);
      return this.getIntelligentDefaults(files);
    }
  }

  /**
   * Use AI service to extract revenue items
   */
  async extractWithAI(files) {
    const prompt = `Extract revenue streams and financial projections from these documents.

Focus on identifying:

1. REVENUE LINE ITEMS:
   - Product/service revenue categories
   - Sales divisions or segments
   - Revenue by geography or market
   - Subscription vs. one-time revenue
   - Recurring vs. non-recurring income

2. HISTORICAL REVENUE DATA:
   - Current annual revenue
   - Previous year revenue
   - Monthly/quarterly revenue trends
   - Revenue by product line

3. GROWTH PROJECTIONS:
   - Revenue growth rates (%)
   - Linear vs. compound growth
   - Growth by segment
   - Market expansion plans

4. REVENUE DRIVERS:
   - Units sold × price per unit
   - Customer count × average revenue per customer
   - Subscription pricing models
   - Volume discounts

Look for tables, charts, and explicit revenue breakdowns.
Extract actual values only - do not estimate or assume.

Return ONLY this structure with actual values found or null:
{
  "revenueItems": [
    {
      "name": "actual revenue stream name",
      "value": numeric_current_annual_value_or_null,
      "growthType": "linear|compound|custom|null",
      "growthRate": percentage_as_number_or_null,
      "category": "product|service|subscription|other|null",
      "historicalData": [
        {
          "period": "YYYY or YYYY-MM",
          "value": numeric_value
        }
      ]
    }
  ],
  "totalRevenue": numeric_total_or_null,
  "revenueGrowthRate": percentage_as_number_or_null,
  "revenueCurrency": "USD|EUR|GBP|etc or null"
}`;

    try {
      const extraction = await this.extractionService.extractFromDocuments(
        files,
        'revenue'
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
    
    // Extract revenue items if not found
    if (!enhanced.revenueItems?.length) {
      const revenueItems = this.extractRevenueItems(allContent);
      if (revenueItems.length > 0) {
        enhanced.revenueItems = revenueItems;
      }
    }
    
    // Extract total revenue if not found
    if (!enhanced.totalRevenue) {
      const totalRevenue = this.extractTotalRevenue(allContent);
      if (totalRevenue) {
        enhanced.totalRevenue = totalRevenue;
      }
    }
    
    // Extract overall growth rate if not found
    if (!enhanced.revenueGrowthRate) {
      const growthRate = this.extractOverallGrowthRate(allContent);
      if (growthRate) {
        enhanced.revenueGrowthRate = growthRate;
      }
    }
    
    return enhanced;
  }

  /**
   * Extract revenue items using pattern matching
   */
  extractRevenueItems(text) {
    const revenueItems = [];
    
    // Pattern 1: Table-style revenue breakdown
    const tablePatterns = [
      // Product/Service | Revenue | Growth
      /([A-Za-z\s&]+)\s*[\|,]\s*(?:\$|USD)?([0-9,]+(?:\.[0-9]+)?)\s*(?:million|billion|m|b)?\s*[\|,]?\s*([0-9.]+)?\s*%?/gi,
      
      // Revenue from [category]: $X million
      /revenue\s+from\s+([A-Za-z\s&]+):\s*(?:\$|USD)?([0-9,]+(?:\.[0-9]+)?)\s*(million|billion|m|b)?/gi,
      
      // [Category] sales: $X
      /([A-Za-z\s&]+)\s+(?:sales|revenue|income):\s*(?:\$|USD)?([0-9,]+(?:\.[0-9]+)?)\s*(million|billion|m|b)?/gi
    ];
    
    for (const pattern of tablePatterns) {
      let match;
      pattern.lastIndex = 0;
      
      while ((match = pattern.exec(text)) !== null) {
        const name = match[1].trim();
        let value = parseFloat(match[2].replace(/,/g, ''));
        const unit = (match[3] || '').toLowerCase();
        const growthRate = match[4] ? parseFloat(match[4]) : null;
        
        // Apply multipliers
        if (unit === 'billion' || unit === 'b') {
          value *= 1000000000;
        } else if (unit === 'million' || unit === 'm') {
          value *= 1000000;
        }
        
        // Validate reasonable values
        if (value >= 10000 && value <= 100000000000 && name.length > 2 && name.length < 100) {
          revenueItems.push({
            name: name,
            value: value,
            growthType: 'linear',
            growthRate: growthRate,
            category: this.categorizeRevenue(name),
            source: 'pattern_matching',
            confidence: 0.7
          });
        }
      }
    }
    
    // Pattern 2: CSV-style data
    const csvPatterns = this.extractCSVRevenue(text);
    revenueItems.push(...csvPatterns);
    
    // Pattern 3: Common revenue categories
    const categoryPatterns = this.extractCategoryRevenue(text);
    revenueItems.push(...categoryPatterns);
    
    // Remove duplicates and validate
    return this.deduplicateRevenueItems(revenueItems);
  }

  /**
   * Extract CSV-style revenue data
   */
  extractCSVRevenue(text) {
    const items = [];
    
    // Look for CSV patterns
    const csvLines = text.split('\n').filter(line => 
      line.includes(',') && 
      /revenue|sales|income/i.test(line) &&
      /[0-9,]+/.test(line)
    );
    
    for (const line of csvLines) {
      const parts = line.split(',').map(p => p.trim());
      
      if (parts.length >= 2) {
        const name = parts[0];
        const valueStr = parts[1];
        
        // Extract numeric value
        const valueMatch = valueStr.match(/([0-9,]+(?:\.[0-9]+)?)/);
        if (valueMatch) {
          const value = parseFloat(valueMatch[1].replace(/,/g, ''));
          
          if (value >= 10000 && name.length > 2) {
            items.push({
              name: name,
              value: value,
              growthType: 'linear',
              growthRate: null,
              category: this.categorizeRevenue(name),
              source: 'csv_pattern',
              confidence: 0.6
            });
          }
        }
      }
    }
    
    return items;
  }

  /**
   * Extract revenue by common categories
   */
  extractCategoryRevenue(text) {
    const items = [];
    
    const commonCategories = [
      'product sales', 'service revenue', 'subscription revenue', 'licensing fees',
      'consulting revenue', 'software sales', 'hardware sales', 'maintenance revenue',
      'support revenue', 'training revenue', 'advertising revenue', 'commission income'
    ];
    
    for (const category of commonCategories) {
      const patterns = [
        new RegExp(`${category}\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?`, 'gi'),
        new RegExp(`([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?\\s*(?:from\\s+)?${category}`, 'gi')
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
            items.push({
              name: category,
              value: value,
              growthType: 'linear',
              growthRate: null,
              category: this.categorizeRevenue(category),
              source: 'category_pattern',
              confidence: 0.5
            });
          }
        }
      }
    }
    
    return items;
  }

  /**
   * Extract total revenue
   */
  extractTotalRevenue(text) {
    const patterns = [
      // Direct total revenue mentions
      {
        regex: /(?:total revenue|total sales|gross revenue|annual revenue)\\s*:?\\s*(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?/gi,
        weight: 1.0
      },
      
      // Revenue of X
      {
        regex: /revenue\\s+of\\s+(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?/gi,
        weight: 0.9
      },
      
      // Generated X in revenue
      {
        regex: /generated\\s+(?:\\$|USD)?([0-9,]+(?:\\.[0-9]+)?)\\s*(million|billion|m|b)?\\s+in\\s+revenue/gi,
        weight: 0.8
      }
    ];
    
    let bestMatch = null;
    let highestWeight = 0;
    
    for (const pattern of patterns) {
      const match = text.match(pattern.regex);
      if (match) {
        let value = parseFloat(match[1].replace(/,/g, ''));
        const unit = (match[2] || '').toLowerCase();
        
        if (unit === 'billion' || unit === 'b') {
          value *= 1000000000;
        } else if (unit === 'million' || unit === 'm') {
          value *= 1000000;
        }
        
        if (value >= 100000 && pattern.weight > highestWeight) {
          highestWeight = pattern.weight;
          bestMatch = {
            value: value,
            confidence: pattern.weight,
            source: 'pattern_matching'
          };
        }
      }
    }
    
    return bestMatch;
  }

  /**
   * Extract overall revenue growth rate
   */
  extractOverallGrowthRate(text) {
    const patterns = [
      // Revenue growth of X%
      /revenue\\s+growth\\s+(?:of\\s+)?([0-9.]+)\\s*%/gi,
      
      // X% revenue growth
      /([0-9.]+)\\s*%\\s+revenue\\s+growth/gi,
      
      // Growing at X%
      /growing\\s+(?:at\\s+)?([0-9.]+)\\s*%/gi,
      
      // Growth rate: X%
      /growth\\s+rate\\s*:?\\s*([0-9.]+)\\s*%/gi
    ];
    
    for (const pattern of patterns) {
      const match = text.match(pattern);
      if (match) {
        const rate = parseFloat(match[1]);
        
        // Validate reasonable growth rate (0% to 100%)
        if (rate >= 0 && rate <= 100) {
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
   * Categorize revenue type
   */
  categorizeRevenue(name) {
    const nameL = name.toLowerCase();
    
    if (nameL.includes('subscription') || nameL.includes('recurring') || nameL.includes('saas')) {
      return 'subscription';
    }
    if (nameL.includes('service') || nameL.includes('consulting') || nameL.includes('support')) {
      return 'service';
    }
    if (nameL.includes('product') || nameL.includes('hardware') || nameL.includes('software') || nameL.includes('sales')) {
      return 'product';
    }
    
    return 'other';
  }

  /**
   * Remove duplicate revenue items
   */
  deduplicateRevenueItems(items) {
    const unique = [];
    const seen = new Set();
    
    for (const item of items) {
      const key = `${item.name.toLowerCase()}_${item.value}`;
      if (!seen.has(key)) {
        seen.add(key);
        unique.push(item);
      }
    }
    
    // Sort by value (highest first)
    return unique.sort((a, b) => b.value - a.value);
  }

  /**
   * Validate revenue items
   */
  validateRevenueItems(data) {
    const validated = { ...data };
    
    // Ensure revenueItems is always an array
    if (!validated.revenueItems || !Array.isArray(validated.revenueItems)) {
      console.warn('💰 Invalid revenueItems format, using empty array');
      validated.revenueItems = [];
    }
    
    if (validated.revenueItems.length > 0) {
      validated.revenueItems = validated.revenueItems.filter(item => {
        // Basic validation
        if (!item.name || !item.value || item.value <= 0) {
          return false;
        }
        
        // Reasonable value range
        if (item.value < 1000 || item.value > 1000000000000) {
          return false;
        }
        
        // Reasonable growth rate
        if (item.growthRate && (item.growthRate < -50 || item.growthRate > 200)) {
          item.growthRate = null;
        }
        
        return true;
      });
    }
    
    // Validate total against sum of items
    if (validated.totalRevenue && validated.revenueItems?.length > 0) {
      const sumOfItems = validated.revenueItems.reduce((sum, item) => sum + item.value, 0);
      const variance = Math.abs(sumOfItems - validated.totalRevenue.value) / validated.totalRevenue.value;
      
      // If variance > 20%, reduce confidence
      if (variance > 0.2) {
        console.warn('💰 Total revenue and sum of items inconsistent');
        validated.totalRevenue.confidence *= 0.8;
      }
    }
    
    return validated;
  }

  /**
   * Enrich with calculated metrics
   */
  enrichWithMetrics(data) {
    const enriched = { ...data };
    
    // Ensure revenueItems is always an array
    if (!enriched.revenueItems || !Array.isArray(enriched.revenueItems)) {
      enriched.revenueItems = [];
    }
    
    if (enriched.revenueItems.length > 0) {
      // Calculate total from items if not provided
      if (!enriched.totalRevenue) {
        const total = enriched.revenueItems.reduce((sum, item) => sum + item.value, 0);
        enriched.totalRevenue = {
          value: total,
          confidence: 0.6,
          source: 'calculated'
        };
      }
      
      // Calculate weighted average growth rate
      if (!enriched.revenueGrowthRate && enriched.revenueItems.some(item => item.growthRate)) {
        const itemsWithGrowth = enriched.revenueItems.filter(item => item.growthRate);
        const totalValue = itemsWithGrowth.reduce((sum, item) => sum + item.value, 0);
        
        if (totalValue > 0) {
          const weightedGrowth = itemsWithGrowth.reduce((sum, item) => 
            sum + (item.growthRate * item.value / totalValue), 0
          );
          
          enriched.revenueGrowthRate = {
            value: weightedGrowth,
            confidence: 0.5,
            source: 'calculated'
          };
        }
      }
      
      // Add revenue mix analysis
      enriched.revenueMix = this.calculateRevenueMix(enriched.revenueItems);
    }
    
    return enriched;
  }

  /**
   * Calculate revenue mix by category
   */
  calculateRevenueMix(items) {
    const mix = {};
    const total = items.reduce((sum, item) => sum + item.value, 0);
    
    for (const item of items) {
      const category = item.category || 'other';
      if (!mix[category]) {
        mix[category] = { value: 0, percentage: 0, items: [] };
      }
      mix[category].value += item.value;
      mix[category].items.push(item.name);
    }
    
    // Calculate percentages
    for (const category of Object.keys(mix)) {
      mix[category].percentage = (mix[category].value / total) * 100;
    }
    
    return mix;
  }

  /**
   * Score confidence based on cross-validation
   */
  scoreConfidence(data, files) {
    const scored = {};
    
    for (const [field, value] of Object.entries(data)) {
      if (!value || (typeof value === 'object' && value.value === null)) {
        scored[field] = value;
        continue;
      }
      
      if (field === 'revenueItems' && Array.isArray(value)) {
        scored[field] = value.map(item => ({
          ...item,
          confidence: this.scoreItemConfidence(item, files)
        }));
      } else if (typeof value === 'object' && value.value !== undefined) {
        let confidence = value.confidence || 0.5;
        
        // Boost confidence for multiple file occurrences
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
    if (item.value >= 100000 && item.value <= 10000000000) {
      confidence += 0.1;
    }
    
    // Boost for complete data
    if (item.name && item.value && item.category) {
      confidence += 0.1;
    }
    
    // Boost for reasonable growth rates
    if (item.growthRate && item.growthRate >= 0 && item.growthRate <= 50) {
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
    console.log('💰 Using intelligent defaults for revenue items');
    
    return {
      revenueItems: [],
      totalRevenue: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      revenueGrowthRate: {
        value: null,
        confidence: 0,
        source: 'not_found'
      },
      revenueCurrency: {
        value: null,
        confidence: 0,
        source: 'not_found'
      }
    };
  }

  /**
   * Apply extracted revenue items to form
   */
  async applyToForm(extractedData) {
    console.log('💰 Applying revenue items to form');
    
    return await this.mappingEngine.applyDataToForm(extractedData, {
      section: 'revenueItems',
      showConfidence: true,
      animateChanges: true
    });
  }
}

// Export for use
window.RevenueItemsExtractor = RevenueItemsExtractor;