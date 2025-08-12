/**
 * LangChain Excel Tools with Proper Tool Calling
 * Implements structured tools for accurate Excel value retrieval
 */

class LangChainExcelTools {
  constructor() {
    this.tools = [];
    this.valueFinder = null;
    this.initialize();
  }

  async initialize() {
    console.log('🛠️ Initializing LangChain Excel Tools...');
    
    // Initialize AccurateExcelValueFinder
    if (typeof AccurateExcelValueFinder !== 'undefined') {
      this.valueFinder = new AccurateExcelValueFinder();
      console.log('✅ Value finder initialized');
    }
    
    // Create tool definitions
    this.createTools();
    console.log('✅ LangChain Excel Tools ready');
  }

  /**
   * Create structured tool definitions for LangChain
   * Following the exact pattern from documentation
   */
  createTools() {
    // Tool 1: Search for Financial Metrics
    this.searchFinancialMetricsTool = {
      name: 'search_financial_metrics',
      description: 'Search for financial metrics like IRR, MOIC, Revenue, EBITDA in Excel workbook',
      schema: {
        type: 'object',
        properties: {
          metricType: {
            type: 'string',
            enum: ['IRR', 'MOIC', 'Revenue', 'EBITDA', 'Exit Value', 'Deal Value', 'Equity', 'Debt', 'All'],
            description: 'The type of financial metric to search for'
          },
          includeFormulas: {
            type: 'boolean',
            description: 'Whether to include formula information',
            default: true
          }
        },
        required: ['metricType']
      },
      execute: async (args) => {
        console.log('🔍 Executing search_financial_metrics tool:', args);
        
        if (!this.valueFinder) {
          return JSON.stringify({ error: 'Value finder not available' });
        }
        
        try {
          // Get all metrics from Excel
          const allMetrics = await this.valueFinder.findAllFinancialMetrics();
          
          // Filter based on requested metric type
          if (args.metricType === 'All') {
            return JSON.stringify(allMetrics);
          }
          
          const requestedMetric = allMetrics[args.metricType];
          if (requestedMetric) {
            return JSON.stringify({
              [args.metricType]: requestedMetric
            });
          }
          
          return JSON.stringify({ 
            message: `${args.metricType} not found in Excel`,
            searchedWorkbook: true 
          });
          
        } catch (error) {
          return JSON.stringify({ error: error.message });
        }
      }
    };

    // Tool 2: Read Specific Cell Value
    this.readCellValueTool = {
      name: 'read_cell_value',
      description: 'Read the value from a specific Excel cell',
      schema: {
        type: 'object',
        properties: {
          cellAddress: {
            type: 'string',
            description: 'Cell address (e.g., "B12" or "Sheet1!A1")'
          }
        },
        required: ['cellAddress']
      },
      execute: async (args) => {
        console.log('📍 Executing read_cell_value tool:', args);
        
        try {
          let result = null;
          
          await Excel.run(async (context) => {
            let range;
            
            if (args.cellAddress.includes('!')) {
              const [sheetName, addr] = args.cellAddress.split('!');
              const sheet = context.workbook.worksheets.getItem(sheetName);
              range = sheet.getRange(addr);
            } else {
              const sheet = context.workbook.worksheets.getActiveWorksheet();
              range = sheet.getRange(args.cellAddress);
            }
            
            range.load(['values', 'formulas', 'numberFormat', 'address']);
            await context.sync();
            
            result = {
              address: range.address,
              value: range.values[0][0],
              formula: range.formulas[0][0],
              numberFormat: range.numberFormat[0][0]
            };
          });
          
          return JSON.stringify(result);
          
        } catch (error) {
          return JSON.stringify({ 
            error: `Could not read cell ${args.cellAddress}: ${error.message}` 
          });
        }
      }
    };

    // Tool 3: Search for Value in Workbook
    this.searchValueInWorkbookTool = {
      name: 'search_value_in_workbook',
      description: 'Search for a specific value or text in the Excel workbook',
      schema: {
        type: 'object',
        properties: {
          searchTerm: {
            type: 'string',
            description: 'The value or text to search for'
          },
          exactMatch: {
            type: 'boolean',
            description: 'Whether to search for exact match only',
            default: false
          },
          maxResults: {
            type: 'number',
            description: 'Maximum number of results to return',
            default: 10
          }
        },
        required: ['searchTerm']
      },
      execute: async (args) => {
        console.log('🔎 Executing search_value_in_workbook tool:', args);
        
        try {
          const results = [];
          
          await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load('items');
            await context.sync();
            
            for (const worksheet of worksheets.items) {
              worksheet.load('name');
              const usedRange = worksheet.getUsedRangeOrNullObject();
              usedRange.load(['values']);
              
              await context.sync();
              
              if (!usedRange.isNullObject && usedRange.values) {
                for (let row = 0; row < usedRange.values.length; row++) {
                  for (let col = 0; col < usedRange.values[row].length; col++) {
                    const cellValue = String(usedRange.values[row][col]);
                    const searchLower = args.searchTerm.toLowerCase();
                    const cellLower = cellValue.toLowerCase();
                    
                    const matches = args.exactMatch 
                      ? cellLower === searchLower
                      : cellLower.includes(searchLower);
                    
                    if (matches) {
                      results.push({
                        worksheet: worksheet.name,
                        address: `${worksheet.name}!${this.getCellAddress(row, col)}`,
                        value: usedRange.values[row][col]
                      });
                      
                      if (results.length >= args.maxResults) {
                        return;
                      }
                    }
                  }
                }
              }
            }
          });
          
          return JSON.stringify({
            searchTerm: args.searchTerm,
            resultsFound: results.length,
            results: results
          });
          
        } catch (error) {
          return JSON.stringify({ error: error.message });
        }
      }
    };

    // Tool 4: Get Worksheet Summary
    this.getWorksheetSummaryTool = {
      name: 'get_worksheet_summary',
      description: 'Get a summary of the active worksheet including used range and key metrics',
      schema: {
        type: 'object',
        properties: {
          worksheetName: {
            type: 'string',
            description: 'Name of worksheet to summarize (optional, defaults to active)',
            optional: true
          }
        }
      },
      execute: async (args) => {
        console.log('📊 Executing get_worksheet_summary tool:', args);
        
        try {
          let summary = {};
          
          await Excel.run(async (context) => {
            const worksheet = args.worksheetName 
              ? context.workbook.worksheets.getItem(args.worksheetName)
              : context.workbook.worksheets.getActiveWorksheet();
            
            worksheet.load(['name']);
            const usedRange = worksheet.getUsedRangeOrNullObject();
            usedRange.load(['address', 'rowCount', 'columnCount', 'values']);
            
            await context.sync();
            
            summary = {
              worksheetName: worksheet.name,
              usedRange: usedRange.isNullObject ? null : {
                address: usedRange.address,
                rows: usedRange.rowCount,
                columns: usedRange.columnCount
              },
              hasData: !usedRange.isNullObject
            };
            
            // Count non-empty cells
            if (!usedRange.isNullObject && usedRange.values) {
              let nonEmptyCells = 0;
              for (let row = 0; row < usedRange.values.length; row++) {
                for (let col = 0; col < usedRange.values[row].length; col++) {
                  if (usedRange.values[row][col] !== null && 
                      usedRange.values[row][col] !== undefined && 
                      usedRange.values[row][col] !== '') {
                    nonEmptyCells++;
                  }
                }
              }
              summary.nonEmptyCells = nonEmptyCells;
            }
          });
          
          return JSON.stringify(summary);
          
        } catch (error) {
          return JSON.stringify({ error: error.message });
        }
      }
    };

    // Tool 5: Calculate Metric Relationships
    this.calculateMetricRelationshipsTool = {
      name: 'calculate_metric_relationships',
      description: 'Calculate relationships between financial metrics (e.g., how MOIC relates to IRR)',
      schema: {
        type: 'object',
        properties: {
          metric1: {
            type: 'string',
            description: 'First metric name'
          },
          metric2: {
            type: 'string',
            description: 'Second metric name'
          }
        },
        required: ['metric1', 'metric2']
      },
      execute: async (args) => {
        console.log('📈 Executing calculate_metric_relationships tool:', args);
        
        try {
          if (!this.valueFinder) {
            return JSON.stringify({ error: 'Value finder not available' });
          }
          
          const metrics = await this.valueFinder.findAllFinancialMetrics();
          
          const m1 = metrics[args.metric1];
          const m2 = metrics[args.metric2];
          
          if (!m1 || !m2) {
            return JSON.stringify({
              error: `One or both metrics not found: ${args.metric1}, ${args.metric2}`
            });
          }
          
          // Calculate relationship
          const relationship = {
            [args.metric1]: {
              value: m1.value,
              location: m1.location
            },
            [args.metric2]: {
              value: m2.value,
              location: m2.location
            },
            analysis: this.analyzeRelationship(args.metric1, m1, args.metric2, m2)
          };
          
          return JSON.stringify(relationship);
          
        } catch (error) {
          return JSON.stringify({ error: error.message });
        }
      }
    };

    // Collect all tools
    this.tools = [
      this.searchFinancialMetricsTool,
      this.readCellValueTool,
      this.searchValueInWorkbookTool,
      this.getWorksheetSummaryTool,
      this.calculateMetricRelationshipsTool
    ];
  }

  /**
   * Analyze relationship between two metrics
   */
  analyzeRelationship(name1, metric1, name2, metric2) {
    const analysis = {};
    
    if (name1 === 'MOIC' && name2 === 'IRR') {
      const moic = metric1.rawValue || parseFloat(metric1.value);
      const irr = metric2.rawValue || parseFloat(metric2.value);
      
      if (moic > 3 && irr > 0.25) {
        analysis.interpretation = 'Both MOIC and IRR are strong, indicating excellent returns';
        analysis.concern = 'Verify assumptions are realistic';
      } else if (moic > 2 && irr > 0.15) {
        analysis.interpretation = 'Solid returns within typical PE range';
        analysis.concern = 'None';
      } else {
        analysis.interpretation = 'Returns may be below target';
        analysis.concern = 'Consider optimization strategies';
      }
    }
    
    return analysis;
  }

  /**
   * Get all tools for binding to LLM
   */
  getAllTools() {
    return this.tools;
  }

  /**
   * Format tools for LangChain binding
   */
  formatToolsForLangChain() {
    return this.tools.map(tool => ({
      name: tool.name,
      description: tool.description,
      schema: tool.schema,
      func: tool.execute
    }));
  }

  /**
   * Execute a specific tool by name
   */
  async executeTool(toolName, args) {
    const tool = this.tools.find(t => t.name === toolName);
    if (!tool) {
      throw new Error(`Tool ${toolName} not found`);
    }
    
    return await tool.execute(args);
  }

  /**
   * Helper: Get cell address from indices
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
}

// Initialize globally
if (typeof window !== 'undefined') {
  window.LangChainExcelTools = LangChainExcelTools;
  window.langChainExcelTools = new LangChainExcelTools();
  console.log('✅ LangChain Excel Tools initialized globally');
}