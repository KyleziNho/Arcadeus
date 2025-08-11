/**
 * Excel Structure Fetcher - Phase 1 Implementation
 * Native Office.js integration for professional M&A analysis
 */

class ExcelStructureFetcher {
  constructor() {
    this.cache = new Map();
    this.cacheTimeout = 30000; // 30 seconds cache
  }

  /**
   * Main function: Fetch workbook structure for AI analysis
   */
  async fetchWorkbookStructure(query) {
    const cacheKey = `structure_${query}_${Date.now() - (Date.now() % this.cacheTimeout)}`;
    
    if (this.cache.has(cacheKey)) {
      console.log('📋 Using cached workbook structure');
      return this.cache.get(cacheKey);
    }

    console.log('📊 Fetching fresh workbook structure for:', query);

    try {
      const structure = await Excel.run(async (context) => {
        const workbook = context.workbook;
        const sheets = workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        let structure = { 
          sheets: {},
          keyMetrics: {},
          warnings: [],
          metadata: {
            totalSheets: sheets.items.length,
            timestamp: new Date().toISOString(),
            query: query
          }
        };

        // Get relevant sheets based on query
        const relevantSheets = this.identifyRelevantSheets(query, sheets.items);
        console.log(`🎯 Identified ${relevantSheets.length} relevant sheets:`, relevantSheets.map(s => s.name));
        
        // Process each relevant sheet
        for (let sheet of relevantSheets) {
          try {
            const sheetData = await this.processSheet(sheet, context);
            structure.sheets[sheet.name] = sheetData;

            // Extract key metrics for M&A models
            const metrics = this.extractKeyMetrics(sheet.name, sheetData);
            if (metrics.length > 0) {
              structure.keyMetrics[sheet.name] = metrics;
            }

          } catch (error) {
            console.error(`❌ Error processing sheet ${sheet.name}:`, error);
            if (error.code === "ItemNotFound") {
              structure.sheets[sheet.name] = { empty: true };
            } else {
              structure.warnings.push(`Error reading ${sheet.name}: ${error.message}`);
            }
          }
        }

        // Add M&A model validation
        structure.validation = this.validateMandAModel(structure);
        
        // Add metadata about advanced features
        structure.advancedFeatures = {
          formulaDependencies: Object.keys(structure.sheets).some(sheet => 
            Object.keys(structure.sheets[sheet].formulaDependencies || {}).length > 0),
          namedRanges: Object.values(structure.sheets).some(sheet => 
            sheet.namedRanges && sheet.namedRanges.length > 0),
          tables: Object.values(structure.sheets).some(sheet => 
            sheet.tables && sheet.tables.length > 0),
          maModelPatterns: Object.values(structure.sheets).some(sheet => 
            sheet.formulaAnalysis && sheet.formulaAnalysis.maModelPatterns.length > 0)
        };

        return structure;
      });

      // Serialize and cache
      const serialized = JSON.stringify(structure);
      this.cache.set(cacheKey, serialized);
      
      console.log('✅ Workbook structure fetched successfully');
      return serialized;

    } catch (error) {
      console.error('❌ Failed to fetch workbook structure:', error);
      throw new Error(`Excel access failed: ${error.message}`);
    }
  }

  /**
   * Process individual sheet data with advanced Excel API features
   */
  async processSheet(sheet, context) {
    const usedRange = sheet.getUsedRange();
    
    usedRange.load("address, values, formulas, rowCount, columnCount, numberFormat");
    
    // Load named ranges from the workbook
    const workbook = context.workbook;
    const namedItems = workbook.names;
    namedItems.load("items/name, items/formula");
    
    // Load tables in this worksheet
    const tables = sheet.tables;
    tables.load("items/name, items/range");
    
    await context.sync();

    const sheetData = {
      usedRange: usedRange.address,
      values: usedRange.values,
      formulas: usedRange.formulas,
      numberFormat: usedRange.numberFormat,
      rowCount: usedRange.rowCount,
      columnCount: usedRange.columnCount,
      isEmpty: false,
      namedRanges: [],
      tables: [],
      formulaDependencies: {}
    };

    // Extract named ranges that reference this sheet
    for (const namedItem of namedItems.items) {
      if (namedItem.formula.includes(`'${sheet.name}'!`) || 
          namedItem.formula.includes(`${sheet.name}!`)) {
        sheetData.namedRanges.push({
          name: namedItem.name,
          formula: namedItem.formula
        });
      }
    }

    // Extract table information
    for (const table of tables.items) {
      sheetData.tables.push({
        name: table.name,
        range: table.range.address
      });
    }

    // Add enhanced formula analysis
    sheetData.formulaAnalysis = this.analyzeFormulas(usedRange.formulas);
    
    // Add formula dependency mapping for key cells
    sheetData.formulaDependencies = await this.mapFormulaDependencies(sheet, usedRange, context);
    
    return sheetData;
  }

  /**
   * Smart sheet identification based on query and M&A patterns
   */
  identifyRelevantSheets(query, allSheets) {
    const queryLower = query.toLowerCase();
    const relevantSheets = [];
    
    // M&A model priority sheets (always include if they exist)
    const prioritySheets = [
      'fcf', 'free cash flow', 'cash flow',
      'revenue', 'revenues', 'sales',
      'assumptions', 'inputs', 'parameters',
      'dashboard', 'summary', 'overview',
      'p&l', 'pl', 'income', 'profit',
      'balance', 'bs', 'balance sheet',
      'valuation', 'dcf', 'irr'
    ];
    
    // Score sheets based on relevance
    const sheetScores = allSheets.map(sheet => {
      const sheetName = sheet.name.toLowerCase();
      let score = 0;
      
      // High score for priority M&A sheets
      if (prioritySheets.some(priority => sheetName.includes(priority))) {
        score += 10;
      }
      
      // Medium score for query keyword matches
      if (queryLower.includes(sheetName) || sheetName.includes(queryLower.split(' ')[0])) {
        score += 5;
      }
      
      // Query-specific scoring
      if (queryLower.includes('irr') && (sheetName.includes('fcf') || sheetName.includes('cash'))) {
        score += 8;
      }
      if (queryLower.includes('moic') && sheetName.includes('fcf')) {
        score += 8;
      }
      if (queryLower.includes('revenue') && sheetName.includes('revenue')) {
        score += 8;
      }
      
      return { sheet, score, name: sheet.name };
    });
    
    // Sort by score and take top sheets
    sheetScores.sort((a, b) => b.score - a.score);
    
    // Always include sheets with score > 0, limit to top 5 for performance
    const topSheets = sheetScores
      .filter(item => item.score > 0)
      .slice(0, 5)
      .map(item => item.sheet);
    
    // If no relevant sheets found, include active sheet + first few sheets
    if (topSheets.length === 0) {
      console.log('⚠️ No relevant sheets found, using first 3 sheets');
      return allSheets.slice(0, 3);
    }
    
    return topSheets;
  }

  /**
   * Extract key M&A metrics from sheet data
   */
  extractKeyMetrics(sheetName, sheetData) {
    const metrics = [];
    const values = sheetData.values;
    const formulas = sheetData.formulas;
    
    // M&A model patterns to detect
    const patterns = {
      moic: /moic|multiple.*invested|money.*multiple/i,
      irr: /irr|internal.*rate/i,
      npv: /npv|net.*present/i,
      revenue: /revenue|sales|income(?!.*tax)/i,
      ebitda: /ebitda|operating.*income/i,
      equity: /equity|contribution/i,
      debt: /debt|loan/i,
      exit: /exit|terminal|disposal/i
    };
    
    for (let row = 0; row < values.length; row++) {
      for (let col = 0; col < values[row].length; col++) {
        const cellValue = values[row][col];
        const cellFormula = formulas[row][col];
        
        if (typeof cellValue === 'string') {
          for (const [metricType, pattern] of Object.entries(patterns)) {
            if (pattern.test(cellValue)) {
              // Look for the actual metric value in adjacent cells
              const metricValue = this.findAdjacentValue(values, row, col);
              if (metricValue !== null) {
                metrics.push({
                  type: metricType,
                  label: cellValue.trim(),
                  value: metricValue,
                  location: `${sheetName}!${this.getColumnLetter(col)}${row + 1}`,
                  formula: cellFormula,
                  confidence: this.calculateMetricConfidence(metricType, cellValue, metricValue)
                });
              }
            }
          }
        }
      }
    }
    
    return metrics;
  }

  /**
   * Find adjacent cells that contain the actual metric values
   */
  findAdjacentValue(values, row, col) {
    // Check right, below, and diagonal for the actual metric value
    const adjacentCells = [
      [row, col + 1], [row, col + 2], [row, col + 3], // Right
      [row + 1, col], [row + 2, col], // Below
      [row + 1, col + 1], [row + 1, col + 2] // Diagonal
    ];
    
    for (const [r, c] of adjacentCells) {
      if (r < values.length && c < values[r].length) {
        const value = values[r][c];
        if (typeof value === 'number' && !isNaN(value) && Math.abs(value) > 0) {
          return value;
        }
      }
    }
    return null;
  }

  /**
   * Calculate confidence score for metric detection
   */
  calculateMetricConfidence(metricType, label, value) {
    let confidence = 0.5; // Base confidence
    
    // Boost confidence based on label specificity
    if (label.toLowerCase().includes(metricType)) {
      confidence += 0.3;
    }
    
    // Boost confidence based on value ranges
    switch (metricType) {
      case 'moic':
        if (value > 1 && value < 10) confidence += 0.2;
        break;
      case 'irr':
        if ((value > 0 && value < 1) || (value > 5 && value < 50)) confidence += 0.2;
        break;
      case 'revenue':
        if (value > 1000) confidence += 0.1;
        break;
    }
    
    return Math.min(confidence, 1.0);
  }

  /**
   * Enhanced formula analysis with M&A model patterns
   */
  analyzeFormulas(formulas) {
    const analysis = {
      totalFormulas: 0,
      complexFormulas: [],
      functionTypes: {},
      externalReferences: [],
      circularReferences: [],
      maModelPatterns: [],
      keyCalculations: []
    };
    
    // M&A specific function patterns
    const maPatterns = {
      irr: /XIRR|IRR\(/i,
      npv: /XNPV|NPV\(/i,
      wacc: /(SUMPRODUCT.*WEIGHT|WEIGHTED.*AVERAGE)/i,
      dcf: /(PV\(|FV\(|PMT\()/i,
      leverage: /(DEBT.*EQUITY|LTV|DSCR)/i,
      multiple: /(EV.*EBITDA|P.*E\s*RATIO)/i
    };
    
    for (let row = 0; row < formulas.length; row++) {
      for (let col = 0; col < formulas[row].length; col++) {
        const formula = formulas[row][col];
        
        if (formula && formula.startsWith('=')) {
          analysis.totalFormulas++;
          
          // Detect function types
          const functions = formula.match(/[A-Z]+\(/g);
          if (functions) {
            functions.forEach(func => {
              const funcName = func.slice(0, -1);
              analysis.functionTypes[funcName] = (analysis.functionTypes[funcName] || 0) + 1;
            });
          }
          
          // Detect M&A model patterns
          for (const [patternType, pattern] of Object.entries(maPatterns)) {
            if (pattern.test(formula)) {
              analysis.maModelPatterns.push({
                type: patternType,
                location: `${this.getColumnLetter(col)}${row + 1}`,
                formula: formula.substring(0, 100) + (formula.length > 100 ? '...' : '')
              });
            }
          }
          
          // Detect key calculations (XIRR, XNPV, complex SUMPRODUCTs)
          if (formula.includes('XIRR') || formula.includes('XNPV') || 
              (formula.includes('SUMPRODUCT') && formula.length > 30)) {
            analysis.keyCalculations.push({
              location: `${this.getColumnLetter(col)}${row + 1}`,
              formula: formula,
              type: formula.includes('XIRR') ? 'irr' : 
                    formula.includes('XNPV') ? 'npv' : 'complex'
            });
          }
          
          // Detect complex formulas (5+ functions or 50+ characters)
          if ((functions && functions.length >= 5) || formula.length > 50) {
            analysis.complexFormulas.push({
              location: `${this.getColumnLetter(col)}${row + 1}`,
              formula: formula,
              complexity: functions ? functions.length : 0,
              length: formula.length
            });
          }
          
          // Detect external references
          if (formula.includes('[') && formula.includes(']')) {
            analysis.externalReferences.push({
              location: `${this.getColumnLetter(col)}${row + 1}`,
              formula: formula
            });
          }
        }
      }
    }
    
    return analysis;
  }

  /**
   * Convert column index to Excel column letter
   */
  getColumnLetter(colIndex) {
    let letter = '';
    while (colIndex >= 0) {
      letter = String.fromCharCode(65 + (colIndex % 26)) + letter;
      colIndex = Math.floor(colIndex / 26) - 1;
    }
    return letter;
  }

  /**
   * Map formula dependencies for key cells using Excel API
   */
  async mapFormulaDependencies(sheet, usedRange, context) {
    const dependencies = {};
    const formulas = usedRange.formulas;
    
    // Focus on cells with key M&A formulas (IRR, NPV, complex calculations)
    for (let row = 0; row < formulas.length; row++) {
      for (let col = 0; col < formulas[row].length; col++) {
        const formula = formulas[row][col];
        
        if (formula && formula.startsWith('=') && 
            (formula.includes('XIRR') || formula.includes('XNPV') || 
             formula.includes('SUMPRODUCT') || formula.length > 50)) {
          
          const cellAddress = `${this.getColumnLetter(col)}${row + 1}`;
          
          try {
            // Get formula dependencies using Excel API
            const range = sheet.getRange(cellAddress);
            const precedents = range.getDirectPrecedents();
            const dependents = range.getDirectDependents();
            
            precedents.load("address");
            dependents.load("address");
            await context.sync();
            
            dependencies[cellAddress] = {
              formula: formula,
              precedents: precedents.areas.items.map(area => area.address),
              dependents: dependents.areas.items.map(area => area.address),
              isKeyCalculation: true
            };
            
          } catch (error) {
            // If dependency tracking fails, still record the formula
            dependencies[cellAddress] = {
              formula: formula,
              precedents: [],
              dependents: [],
              isKeyCalculation: true,
              error: error.message
            };
          }
        }
      }
    }
    
    return dependencies;
  }

  /**
   * Set up real-time workbook monitoring (call once on initialization)
   */
  setupWorkbookMonitoring() {
    return Excel.run(async (context) => {
      const workbook = context.workbook;
      
      // Set up change event handler for real-time monitoring
      workbook.onChanged.add(async (event) => {
        console.log('📊 Workbook changed, invalidating cache');
        
        // Clear cache when workbook changes to ensure fresh data
        this.cache.clear();
        
        // Optional: Emit event for UI to refresh if needed
        if (window.dispatchEvent) {
          window.dispatchEvent(new CustomEvent('excelWorkbookChanged', {
            detail: { changeType: event.changeType, source: event.source }
          }));
        }
      });
      
      await context.sync();
      console.log('✅ Workbook monitoring setup complete');
    }).catch(error => {
      console.warn('⚠️ Could not setup workbook monitoring:', error.message);
    });
  }

  /**
   * Enhanced data validation with M&A model checks
   */
  validateMandAModel(structure) {
    const validationResults = {
      warnings: [],
      errors: [],
      suggestions: [],
      modelScore: 0
    };
    
    let score = 0;
    const maxScore = 100;
    
    // Check for essential M&A model components
    const hasIRRCalculation = Object.values(structure.keyMetrics).some(metrics => 
      metrics.some(m => m.type === 'irr'));
    const hasMOICCalculation = Object.values(structure.keyMetrics).some(metrics => 
      metrics.some(m => m.type === 'moic'));
    const hasRevenueProjections = Object.values(structure.keyMetrics).some(metrics => 
      metrics.some(m => m.type === 'revenue'));
    
    if (hasIRRCalculation) {
      score += 30;
    } else {
      validationResults.errors.push('No IRR calculation found - essential for M&A models');
    }
    
    if (hasMOICCalculation) {
      score += 25;
    } else {
      validationResults.warnings.push('MOIC calculation not detected - recommended for investment analysis');
    }
    
    if (hasRevenueProjections) {
      score += 20;
    } else {
      validationResults.warnings.push('Revenue projections not clearly identified');
    }
    
    // Check for formula complexity and external references
    let totalFormulas = 0;
    let totalExternalRefs = 0;
    
    Object.values(structure.sheets).forEach(sheet => {
      if (sheet.formulaAnalysis) {
        totalFormulas += sheet.formulaAnalysis.totalFormulas;
        totalExternalRefs += sheet.formulaAnalysis.externalReferences.length;
      }
    });
    
    if (totalFormulas > 50) {
      score += 15;
    } else if (totalFormulas > 10) {
      score += 10;
    }
    
    if (totalExternalRefs > 0) {
      validationResults.warnings.push(`${totalExternalRefs} external references found - may cause calculation issues`);
      score -= 5;
    }
    
    // Final score and recommendations
    validationResults.modelScore = Math.max(0, Math.min(score, maxScore));
    
    if (validationResults.modelScore >= 80) {
      validationResults.suggestions.push('Excellent M&A model structure detected');
    } else if (validationResults.modelScore >= 60) {
      validationResults.suggestions.push('Good model structure - consider adding missing M&A components');
    } else {
      validationResults.suggestions.push('Model may benefit from standard M&A analysis components (IRR, MOIC, cash flows)');
    }
    
    return validationResults;
  }

  /**
   * Clear cache
   */
  clearCache() {
    this.cache.clear();
    console.log('🗑️ Excel structure cache cleared');
  }

  /**
   * Get cache statistics
   */
  getCacheStats() {
    return {
      size: this.cache.size,
      keys: Array.from(this.cache.keys()),
      timeout: this.cacheTimeout
    };
  }
}

// Export for global use
window.ExcelStructureFetcher = ExcelStructureFetcher;
window.excelStructureFetcher = new ExcelStructureFetcher();

// Initialize workbook monitoring when Excel is available
if (typeof Excel !== 'undefined') {
  // Set up monitoring after a brief delay to ensure Excel is fully loaded
  setTimeout(() => {
    window.excelStructureFetcher.setupWorkbookMonitoring();
  }, 1000);
}

console.log('📊 Enhanced Excel Structure Fetcher loaded with advanced API features:');
console.log('  ✅ Formula dependency tracking (getDirectPrecedents/Dependents)');
console.log('  ✅ Named ranges detection');
console.log('  ✅ Table detection');
console.log('  ✅ M&A model pattern recognition');
console.log('  ✅ Real-time workbook monitoring');
console.log('  ✅ Enhanced data validation');
console.log('🎯 Ready for Phase 1 testing with professional M&A analysis capabilities');