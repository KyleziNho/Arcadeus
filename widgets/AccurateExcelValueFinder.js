/**
 * Accurate Excel Value Finder
 * Intelligently finds and extracts actual values from Excel worksheets
 * Handles various cell layouts and formats to ensure AI uses real data
 */

class AccurateExcelValueFinder {
  constructor() {
    this.cachedValues = {};
    this.lastUpdate = null;
    console.log('✅ AccurateExcelValueFinder initialized');
  }

  /**
   * Find all financial metrics with their actual values
   * Uses multiple strategies to locate the correct values
   */
  async findAllFinancialMetrics() {
    console.log('🔍 Starting comprehensive financial metrics search...');
    
    try {
      const metrics = {};
      
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load('items');
        await context.sync();

        // Define search patterns with variations
        const searchPatterns = {
          'IRR': {
            terms: ['irr', 'internal rate of return', 'return rate', 'project irr'],
            variations: ['unlevered irr', 'levered irr', 'equity irr', 'project irr']
          },
          'MOIC': {
            terms: ['moic', 'multiple on invested capital', 'money multiple', 'total moic'],
            variations: ['gross moic', 'net moic', 'realized moic']
          },
          'Revenue': {
            terms: ['revenue', 'sales', 'total revenue', 'net revenue'],
            variations: ['annual revenue', 'monthly revenue', 'projected revenue']
          },
          'EBITDA': {
            terms: ['ebitda', 'earnings before', 'operating income'],
            variations: ['adjusted ebitda', 'normalized ebitda']
          },
          'Exit Value': {
            terms: ['exit value', 'terminal value', 'enterprise value', 'sale price'],
            variations: ['exit enterprise value', 'exit equity value']
          },
          'Deal Value': {
            terms: ['deal value', 'purchase price', 'acquisition price', 'transaction value'],
            variations: ['total deal value', 'enterprise value']
          },
          'Equity': {
            terms: ['equity', 'equity investment', 'equity contribution', 'sponsor equity'],
            variations: ['initial equity', 'total equity']
          },
          'Debt': {
            terms: ['debt', 'debt financing', 'leverage', 'total debt'],
            variations: ['senior debt', 'subordinated debt', 'term loan']
          }
        };

        // Process each worksheet
        for (const worksheet of worksheets.items) {
          worksheet.load('name');
          const usedRange = worksheet.getUsedRangeOrNullObject();
          usedRange.load(['values', 'formulas', 'numberFormat']);
          
          await context.sync();

          if (!usedRange.isNullObject && usedRange.values) {
            console.log(`📊 Scanning worksheet: ${worksheet.name}`);
            
            // Scan for metrics using multiple strategies
            for (let row = 0; row < usedRange.values.length; row++) {
              for (let col = 0; col < usedRange.values[row].length; col++) {
                const cellValue = usedRange.values[row][col];
                const cellText = String(cellValue).toLowerCase().trim();
                
                // Check each metric pattern
                for (const [metricName, pattern] of Object.entries(searchPatterns)) {
                  const allTerms = [...pattern.terms, ...pattern.variations];
                  
                  if (this.matchesPattern(cellText, allTerms)) {
                    // Found a label, now find the associated value
                    const foundValue = await this.findAssociatedValue(
                      usedRange,
                      row,
                      col,
                      worksheet.name,
                      metricName
                    );
                    
                    if (foundValue) {
                      // Store the best match or update if this is better
                      if (!metrics[metricName] || this.isBetterMatch(foundValue, metrics[metricName])) {
                        metrics[metricName] = foundValue;
                        console.log(`✅ Found ${metricName}: ${foundValue.value} at ${foundValue.location}`);
                      }
                    }
                  }
                }
              }
            }
          }
        }
      });

      // Cache the results
      this.cachedValues = metrics;
      this.lastUpdate = new Date().toISOString();
      
      console.log('📊 Metrics found:', Object.keys(metrics).length);
      return metrics;
      
    } catch (error) {
      console.error('❌ Error finding financial metrics:', error);
      return {};
    }
  }

  /**
   * Check if text matches any of the patterns
   */
  matchesPattern(text, patterns) {
    return patterns.some(pattern => {
      // Exact match or contains
      return text === pattern || text.includes(pattern);
    });
  }

  /**
   * Find the actual value associated with a label
   * Tries multiple strategies to find the correct value
   */
  async findAssociatedValue(usedRange, labelRow, labelCol, sheetName, metricName) {
    const values = usedRange.values;
    const formulas = usedRange.formulas;
    const numberFormats = usedRange.numberFormat;
    
    // Strategy 1: Look to the right (same row)
    for (let offset = 1; offset <= 5; offset++) {
      if (labelCol + offset < values[labelRow].length) {
        const value = values[labelRow][labelCol + offset];
        if (this.isValidValue(value, metricName)) {
          return {
            value: this.formatValue(value, metricName),
            rawValue: value,
            location: `${sheetName}!${this.getCellAddress(labelRow, labelCol + offset)}`,
            formula: formulas[labelRow][labelCol + offset],
            numberFormat: numberFormats[labelRow][labelCol + offset],
            confidence: 'high',
            strategy: 'right_adjacent'
          };
        }
      }
    }
    
    // Strategy 2: Look below (same column)
    for (let offset = 1; offset <= 3; offset++) {
      if (labelRow + offset < values.length) {
        const value = values[labelRow + offset][labelCol];
        if (this.isValidValue(value, metricName)) {
          return {
            value: this.formatValue(value, metricName),
            rawValue: value,
            location: `${sheetName}!${this.getCellAddress(labelRow + offset, labelCol)}`,
            formula: formulas[labelRow + offset][labelCol],
            numberFormat: numberFormats[labelRow + offset][labelCol],
            confidence: 'high',
            strategy: 'below'
          };
        }
      }
    }
    
    // Strategy 3: Look for colon pattern (Label: Value)
    const labelText = String(values[labelRow][labelCol]);
    if (labelText.includes(':')) {
      const parts = labelText.split(':');
      if (parts.length === 2) {
        const potentialValue = parts[1].trim();
        const numericValue = this.parseNumericValue(potentialValue);
        if (numericValue !== null) {
          return {
            value: this.formatValue(numericValue, metricName),
            rawValue: numericValue,
            location: `${sheetName}!${this.getCellAddress(labelRow, labelCol)}`,
            formula: formulas[labelRow][labelCol],
            numberFormat: numberFormats[labelRow][labelCol],
            confidence: 'medium',
            strategy: 'inline_colon'
          };
        }
      }
    }
    
    // Strategy 4: Look in the row above or below for a year/period, then right
    // (Common in financial models where metrics are in a grid)
    if (labelRow > 0) {
      const aboveRow = values[labelRow - 1];
      for (let col = labelCol + 1; col < Math.min(labelCol + 10, aboveRow.length); col++) {
        const headerValue = aboveRow[col];
        if (this.looksLikeYearOrPeriod(headerValue)) {
          const value = values[labelRow][col];
          if (this.isValidValue(value, metricName)) {
            return {
              value: this.formatValue(value, metricName),
              rawValue: value,
              location: `${sheetName}!${this.getCellAddress(labelRow, col)}`,
              formula: formulas[labelRow][col],
              numberFormat: numberFormats[labelRow][col],
              period: String(headerValue),
              confidence: 'high',
              strategy: 'grid_with_period'
            };
          }
        }
      }
    }
    
    return null;
  }

  /**
   * Check if a value is valid for the given metric type
   */
  isValidValue(value, metricName) {
    if (value === null || value === undefined || value === '') {
      return false;
    }
    
    // Check if it's a number
    if (typeof value === 'number') {
      // IRR and percentages should be between -1 and 10 (allowing for -100% to 1000%)
      if (metricName === 'IRR') {
        return value >= -1 && value <= 10;
      }
      
      // MOIC should be positive and typically between 0 and 20
      if (metricName === 'MOIC') {
        return value > 0 && value <= 20;
      }
      
      // Other metrics should just be non-zero numbers
      return value !== 0;
    }
    
    // Check if it's a formula result
    if (typeof value === 'string' && value.startsWith('=')) {
      return false; // This is a formula, not a value
    }
    
    // Check if it's a string that looks like a number
    const numericValue = this.parseNumericValue(value);
    return numericValue !== null && this.isValidValue(numericValue, metricName);
  }

  /**
   * Parse a numeric value from various formats
   */
  parseNumericValue(value) {
    if (typeof value === 'number') {
      return value;
    }
    
    if (typeof value === 'string') {
      // Remove common formatting
      let cleaned = value
        .replace(/[$,]/g, '')
        .replace(/\s/g, '')
        .replace(/[()]/g, match => match === '(' ? '-' : '');
      
      // Handle percentages
      if (cleaned.includes('%')) {
        cleaned = cleaned.replace('%', '');
        const num = parseFloat(cleaned);
        return isNaN(num) ? null : num / 100;
      }
      
      // Handle 'x' notation for multiples
      if (cleaned.toLowerCase().endsWith('x')) {
        cleaned = cleaned.slice(0, -1);
      }
      
      const num = parseFloat(cleaned);
      return isNaN(num) ? null : num;
    }
    
    return null;
  }

  /**
   * Format value for display based on metric type
   */
  formatValue(value, metricName) {
    if (metricName === 'IRR') {
      return `${(value * 100).toFixed(1)}%`;
    }
    
    if (metricName === 'MOIC') {
      return `${value.toFixed(2)}x`;
    }
    
    if (value >= 1000000) {
      return `$${(value / 1000000).toFixed(1)}M`;
    }
    
    if (value >= 1000) {
      return `$${(value / 1000).toFixed(1)}K`;
    }
    
    return `$${value.toFixed(2)}`;
  }

  /**
   * Check if a value looks like a year or period header
   */
  looksLikeYearOrPeriod(value) {
    if (typeof value === 'number') {
      // Check if it's a year (2020-2030)
      return value >= 2020 && value <= 2030;
    }
    
    if (typeof value === 'string') {
      // Check for year patterns
      if (/20\d{2}/.test(value)) {
        return true;
      }
      
      // Check for period patterns (Year 1, Y1, Period 1, etc.)
      if (/^(year|yr|y|period|p)\s*\d+$/i.test(value)) {
        return true;
      }
    }
    
    return false;
  }

  /**
   * Determine if one match is better than another
   */
  isBetterMatch(newMatch, existingMatch) {
    // Prefer high confidence over lower confidence
    if (newMatch.confidence === 'high' && existingMatch.confidence !== 'high') {
      return true;
    }
    
    // Prefer matches with formulas (calculated values)
    if (newMatch.formula && !existingMatch.formula) {
      return true;
    }
    
    // Prefer grid matches with periods (more specific)
    if (newMatch.strategy === 'grid_with_period' && existingMatch.strategy !== 'grid_with_period') {
      return true;
    }
    
    return false;
  }

  /**
   * Get cell address from row/col indices
   */
  getCellAddress(row, col) {
    let columnLetter = '';
    let temp = col;
    while (temp >= 0) {
      columnLetter = String.fromCharCode(65 + (temp % 26)) + columnLetter;
      temp = Math.floor(temp / 26) - 1;
    }
    return `${columnLetter}${row + 1}`;
  }

  /**
   * Get a specific metric value with details
   */
  async getMetricValue(metricName) {
    // Check cache first
    if (this.cachedValues[metricName] && this.lastUpdate) {
      const cacheAge = Date.now() - new Date(this.lastUpdate).getTime();
      if (cacheAge < 30000) { // 30 second cache
        console.log(`📦 Using cached value for ${metricName}`);
        return this.cachedValues[metricName];
      }
    }
    
    // Refresh values
    console.log(`🔄 Refreshing values to get ${metricName}`);
    const metrics = await this.findAllFinancialMetrics();
    return metrics[metricName] || null;
  }

  /**
   * Get formatted summary for AI context
   */
  async getFormattedSummaryForAI() {
    const metrics = await this.findAllFinancialMetrics();
    
    if (Object.keys(metrics).length === 0) {
      return 'No financial metrics found in the Excel workbook.';
    }
    
    let summary = 'ACTUAL EXCEL VALUES (verified from workbook):\n\n';
    
    for (const [metricName, data] of Object.entries(metrics)) {
      summary += `${metricName}:\n`;
      summary += `  • Value: ${data.value}\n`;
      summary += `  • Location: ${data.location}\n`;
      if (data.period) {
        summary += `  • Period: ${data.period}\n`;
      }
      if (data.formula) {
        summary += `  • Formula: ${data.formula}\n`;
      }
      summary += `  • Confidence: ${data.confidence}\n\n`;
    }
    
    return summary;
  }
}

// Initialize and make globally available
if (typeof window !== 'undefined') {
  window.AccurateExcelValueFinder = AccurateExcelValueFinder;
  
  // Auto-initialize when Excel is ready
  if (typeof Office !== 'undefined') {
    Office.onReady(() => {
      window.excelValueFinder = new AccurateExcelValueFinder();
      console.log('✅ AccurateExcelValueFinder ready for use');
    });
  }
}