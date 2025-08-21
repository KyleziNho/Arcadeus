/**
 * ExcelLiveAnalyzer.js
 * Advanced Excel analysis inspired by document analysis agents
 * Provides comprehensive, real-time Excel context and analysis
 */

class ExcelLiveAnalyzer {
  constructor() {
    this.cache = new Map();
    this.structureCache = null;
    this.lastScan = null;
    this.isScanning = false;
    this.changeListeners = [];
    this.analysisHistory = [];
    
    // Financial patterns for M&A models
    this.financialPatterns = {
      moic: /moic|multiple.*invested|money.*multiple/i,
      irr: /irr|internal.*rate|discount.*rate/i,
      cashFlow: /cash.*flow|fcf|free.*cash/i,
      revenue: /revenue|sales|income(?!.*tax)/i,
      ebitda: /ebitda|operating.*income/i,
      assumptions: /assumption|input|driver/i,
      exit: /exit|terminal|disposal/i
    };
  }

  /**
   * Safely get a property value with fallback
   */
  safeGetProperty(obj, property, fallback = 'Unknown') {
    try {
      return obj[property] || fallback;
    } catch (error) {
      console.warn(`Property '${property}' not available:`, error.message);
      return fallback;
    }
  }

  /**
   * Get comprehensive Excel context - far beyond current 10-row limit
   */
  async getComprehensiveContext() {
    if (typeof Excel === 'undefined') {
      return { available: false, error: 'Excel not available' };
    }

    console.log('🔍 Getting comprehensive Excel context...');
    
    try {
      return await Excel.run(async (context) => {
        const workbook = context.workbook;
        const activeWorksheet = workbook.worksheets.getActiveWorksheet();
        const selectedRange = workbook.getSelectedRange();
        
        // Load basic info
        workbook.load(['name']);
        activeWorksheet.load(['name', 'charts', 'tables']);
        selectedRange.load(['address', 'values', 'formulas']);
        
        // Get all worksheets
        const worksheets = workbook.worksheets;
        worksheets.load(['items/name']);
        
        // Get named ranges
        const namedRanges = workbook.names;
        namedRanges.load(['items/name', 'items/formula', 'items/value']);
        
        await context.sync();
        
        // Get full used range (not limited to 10 rows) - handle null case
        const usedRange = activeWorksheet.getUsedRangeOrNullObject();
        usedRange.load(['address', 'values', 'formulas', 'rowCount', 'columnCount']);
        await context.sync();
        
        // Check if used range actually exists
        const hasUsedRange = !usedRange.isNullObject;
        
        // Comprehensive analysis
        const analysis = await this.analyzeWorkbookStructure(
          workbook,
          activeWorksheet,
          hasUsedRange ? usedRange : null,
          worksheets.items,
          namedRanges.items
        );
        
        const comprehensiveContext = {
          timestamp: new Date(),
          workbook: {
            name: this.safeGetProperty(workbook, 'name', 'Unknown Workbook')
          },
          activeWorksheet: {
            name: this.safeGetProperty(activeWorksheet, 'name', 'Unknown Worksheet'),
            usedRange: hasUsedRange ? {
              address: usedRange.address,
              rowCount: usedRange.rowCount,
              columnCount: usedRange.columnCount,
              values: usedRange.values,
              formulas: usedRange.formulas
            } : null
          },
          selectedRange: {
            address: selectedRange.address,
            values: selectedRange.values,
            formulas: selectedRange.formulas
          },
          allWorksheets: worksheets.items.map(ws => ({
            name: this.safeGetProperty(ws, 'name', 'Unknown')
          })),
          namedRanges: namedRanges.items.map(range => ({
            name: this.safeGetProperty(range, 'name', 'Unknown'),
            formula: this.safeGetProperty(range, 'formula', ''),
            value: this.safeGetProperty(range, 'value', '')
          })),
          analysis: analysis,
          financialMetrics: await this.extractFinancialMetrics(hasUsedRange ? usedRange : null),
          calculationChains: await this.mapCalculationDependencies(hasUsedRange ? usedRange : null)
        };
        
        // Cache the structure for performance
        this.structureCache = comprehensiveContext;
        this.lastScan = new Date();
        
        return comprehensiveContext;
      });
    } catch (error) {
      console.error('❌ Comprehensive context analysis failed:', error);
      return { 
        available: false, 
        error: error.message,
        fallback: await this.getFallbackContext()
      };
    }
  }

  /**
   * Analyze workbook structure to understand data organization
   */
  async analyzeWorkbookStructure(workbook, worksheet, usedRange, worksheets, namedRanges) {
    if (!usedRange || !usedRange.values) {
      return { structure: 'empty', sections: [] };
    }

    const analysis = {
      structure: 'unknown',
      sections: [],
      dataTypes: {},
      financialAreas: {},
      inputAreas: [],
      calculationAreas: [],
      outputAreas: []
    };

    try {
      const values = usedRange.values;
      const formulas = usedRange.formulas;
      
      // Analyze data patterns
      for (let row = 0; row < values.length; row++) {
        for (let col = 0; col < values[row].length; col++) {
          const value = values[row][col];
          const formula = formulas[row][col];
          
          if (typeof value === 'string' && value.length > 0) {
            // Check for financial keywords in headers
            for (const [metric, pattern] of Object.entries(this.financialPatterns)) {
              if (pattern.test(value)) {
                if (!analysis.financialAreas[metric]) {
                  analysis.financialAreas[metric] = [];
                }
                analysis.financialAreas[metric].push({
                  location: `${String.fromCharCode(65 + col)}${row + 1}`,
                  text: value,
                  context: this.getContextAround(values, row, col)
                });
              }
            }
          }
          
          // Identify calculation areas (cells with formulas)
          if (formula && typeof formula === 'string' && formula.startsWith('=')) {
            analysis.calculationAreas.push({
              location: `${String.fromCharCode(65 + col)}${row + 1}`,
              formula: formula,
              result: value,
              complexity: this.assessFormulaComplexity(formula)
            });
          }
          
          // Identify input areas (hardcoded numbers)
          if (typeof value === 'number' && (!formula || typeof formula !== 'string' || !formula.startsWith('='))) {
            analysis.inputAreas.push({
              location: `${String.fromCharCode(65 + col)}${row + 1}`,
              value: value,
              type: this.classifyInputType(value, row, col, values)
            });
          }
        }
      }

      // Determine overall structure
      if (analysis.financialAreas.moic || analysis.financialAreas.irr) {
        analysis.structure = 'financial_model';
      } else if (analysis.calculationAreas.length > analysis.inputAreas.length) {
        analysis.structure = 'calculation_heavy';
      } else {
        analysis.structure = 'data_table';
      }

      return analysis;
    } catch (error) {
      console.error('Structure analysis failed:', error);
      return analysis;
    }
  }

  /**
   * Extract financial metrics with context
   */
  async extractFinancialMetrics(usedRange) {
    if (!usedRange || !usedRange.values) return {};

    const metrics = {
      moic: null,
      irr: null,
      revenue: [],
      costs: [],
      cashFlows: [],
      assumptions: []
    };

    try {
      const values = usedRange.values;
      const formulas = usedRange.formulas;

      for (let row = 0; row < values.length; row++) {
        for (let col = 0; col < values[row].length; col++) {
          const value = values[row][col];
          const formula = formulas[row][col];
          const location = `${String.fromCharCode(65 + col)}${row + 1}`;

          // Look for MOIC calculations
          if (formula && formula.toLowerCase().includes('moic')) {
            metrics.moic = {
              location: location,
              value: value,
              formula: formula,
              interpretation: this.interpretMOIC(value)
            };
          }

          // Look for IRR calculations
          if (formula && (formula.toLowerCase().includes('irr') || formula.toLowerCase().includes('xirr'))) {
            metrics.irr = {
              location: location,
              value: value,
              formula: formula,
              interpretation: this.interpretIRR(value)
            };
          }

          // Identify cash flows (arrays of numbers)
          if (typeof value === 'number' && Math.abs(value) > 1000) {
            const context = this.getContextAround(values, row, col);
            if (context.some(cell => typeof cell === 'string' && /cash.*flow|cf/i.test(cell))) {
              metrics.cashFlows.push({
                location: location,
                value: value,
                period: this.identifyTimePeriod(values, row, col)
              });
            }
          }
        }
      }

      return metrics;
    } catch (error) {
      console.error('Financial metrics extraction failed:', error);
      return metrics;
    }
  }

  /**
   * Map calculation dependencies
   */
  async mapCalculationDependencies(usedRange) {
    if (!usedRange || !usedRange.formulas) return {};

    const dependencies = {};
    
    try {
      const formulas = usedRange.formulas;
      
      for (let row = 0; row < formulas.length; row++) {
        for (let col = 0; col < formulas[row].length; col++) {
          const formula = formulas[row][col];
          const location = `${String.fromCharCode(65 + col)}${row + 1}`;
          
          if (formula && typeof formula === 'string' && formula.startsWith('=')) {
            dependencies[location] = {
              formula: formula,
              references: this.extractCellReferences(formula),
              complexity: this.assessFormulaComplexity(formula)
            };
          }
        }
      }
      
      return dependencies;
    } catch (error) {
      console.error('Dependency mapping failed:', error);
      return {};
    }
  }

  /**
   * Start continuous monitoring of Excel changes
   */
  async startLiveMonitoring() {
    if (this.isScanning || typeof Excel === 'undefined') return;

    console.log('🔄 Starting live Excel monitoring...');
    this.isScanning = true;

    try {
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        
        // Set up change event listeners
        workbook.onSelectionChanged.add(async (event) => {
          console.log('👆 Selection changed:', event.address);
          await this.handleSelectionChange(event);
        });

        workbook.worksheets.onChanged.add(async (event) => {
          console.log('📝 Data changed:', event.address);
          await this.handleDataChange(event);
        });

        await context.sync();
      });

      // Also poll for comprehensive updates every 30 seconds
      this.monitoringInterval = setInterval(async () => {
        if (document.hidden) return; // Don't scan when tab is inactive
        
        try {
          const newContext = await this.getComprehensiveContext();
          this.notifyChangeListeners('comprehensive_update', newContext);
        } catch (error) {
          console.error('Monitoring scan failed:', error);
        }
      }, 30000);

    } catch (error) {
      console.error('❌ Failed to start live monitoring:', error);
      this.isScanning = false;
    }
  }

  /**
   * Handle selection changes
   */
  async handleSelectionChange(event) {
    const selectionContext = {
      type: 'selection_change',
      address: event.address,
      timestamp: new Date()
    };

    this.notifyChangeListeners('selection_change', selectionContext);
  }

  /**
   * Handle data changes
   */
  async handleDataChange(event) {
    // Invalidate relevant cache
    this.cache.delete('comprehensive_context');
    
    const changeContext = {
      type: 'data_change',
      address: event.address,
      changeType: event.changeType,
      timestamp: new Date()
    };

    // Get updated context for changed area
    try {
      const updatedContext = await this.getComprehensiveContext();
      changeContext.updatedData = updatedContext;
      
      this.notifyChangeListeners('data_change', changeContext);
    } catch (error) {
      console.error('Failed to get updated context:', error);
    }
  }

  /**
   * Add change listener
   */
  addChangeListener(callback) {
    this.changeListeners.push(callback);
  }

  /**
   * Notify all change listeners
   */
  notifyChangeListeners(eventType, data) {
    for (const callback of this.changeListeners) {
      try {
        callback(eventType, data);
      } catch (error) {
        console.error('Change listener error:', error);
      }
    }
  }

  /**
   * Stop monitoring
   */
  stopLiveMonitoring() {
    this.isScanning = false;
    if (this.monitoringInterval) {
      clearInterval(this.monitoringInterval);
      this.monitoringInterval = null;
    }
    console.log('⏹️ Live monitoring stopped');
  }

  // Helper methods
  
  getContextAround(values, row, col, radius = 2) {
    const context = [];
    for (let r = Math.max(0, row - radius); r <= Math.min(values.length - 1, row + radius); r++) {
      for (let c = Math.max(0, col - radius); c <= Math.min(values[r].length - 1, col + radius); c++) {
        if (r !== row || c !== col) {
          context.push(values[r][c]);
        }
      }
    }
    return context;
  }

  assessFormulaComplexity(formula) {
    if (!formula) return 'none';
    const functions = (formula.match(/[A-Z]+\(/g) || []).length;
    const references = (formula.match(/[A-Z]+\d+/g) || []).length;
    
    if (functions > 5 || references > 10) return 'high';
    if (functions > 2 || references > 5) return 'medium';
    return 'low';
  }

  extractCellReferences(formula) {
    if (!formula || typeof formula !== 'string') return [];
    return (formula.match(/[A-Z]+\d+/g) || []);
  }

  interpretMOIC(value) {
    if (typeof value !== 'number') return 'invalid';
    if (value > 5) return 'very_high';
    if (value > 3) return 'high';
    if (value > 2) return 'good';
    if (value > 1.5) return 'moderate';
    return 'low';
  }

  interpretIRR(value) {
    if (typeof value !== 'number') return 'invalid';
    const percentage = value > 1 ? value : value * 100;
    if (percentage > 30) return 'very_high';
    if (percentage > 20) return 'high';
    if (percentage > 15) return 'good';
    if (percentage > 10) return 'moderate';
    return 'low';
  }

  classifyInputType(value, row, col, values) {
    // Simple heuristics for input classification
    if (value > 0.5 && value < 1.5) return 'multiplier';
    if (value > 1000000) return 'large_currency';
    if (value > 1000) return 'currency';
    if (value > 0 && value < 1) return 'percentage';
    return 'number';
  }

  identifyTimePeriod(values, row, col) {
    // Look for year/period indicators nearby
    const context = this.getContextAround(values, row, col, 1);
    for (const cell of context) {
      if (typeof cell === 'string' || typeof cell === 'number') {
        const str = String(cell);
        if (/20\d{2}/.test(str)) return `Year ${str}`;
        if (/year|yr|period/i.test(str)) return str;
      }
    }
    return 'unknown';
  }

  async getFallbackContext() {
    // Fallback to basic context if comprehensive analysis fails
    try {
      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load(['name']);
        await context.sync();
        
        return {
          worksheetName: worksheet.name || 'Unknown Worksheet',
          error: 'Limited context due to analysis failure'
        };
      });
    } catch (error) {
      return { error: 'Excel completely unavailable' };
    }
  }

  /**
   * Get optimized context for AI analysis
   */
  async getOptimizedContextForAI() {
    const full = await this.getComprehensiveContext();
    
    if (!full || full.error) return full;

    // Optimize for AI consumption - send only relevant data with safe property access
    return {
      structure: full.analysis?.structure || 'unknown',
      financialMetrics: full.financialMetrics || {},
      selectedArea: full.selectedRange || null,
      keyCalculations: full.calculationChains || {},
      worksheetName: full.activeWorksheet?.name || 'Unknown Worksheet',
      summary: this.generateContextSummary(full)
    };
  }

  generateContextSummary(context) {
    try {
      const summary = [];
      
      if (context?.analysis?.structure) {
        summary.push(`Workbook type: ${context.analysis.structure}`);
      }
      
      if (context?.financialMetrics?.moic) {
        summary.push(`MOIC: ${context.financialMetrics.moic.value} (${context.financialMetrics.moic.interpretation})`);
      }
      
      if (context?.financialMetrics?.irr) {
        summary.push(`IRR: ${context.financialMetrics.irr.value} (${context.financialMetrics.irr.interpretation})`);
      }
      
      if (context?.activeWorksheet?.usedRange) {
        summary.push(`Data range: ${context.activeWorksheet.usedRange.address}`);
      }
      
      return summary.join(' | ') || 'Context analysis complete';
    } catch (error) {
      console.warn('Failed to generate context summary:', error);
      return 'Context analysis complete';
    }
  }
}

// Export for use
window.ExcelLiveAnalyzer = ExcelLiveAnalyzer;