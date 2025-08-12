/**
 * Proper LangChain Tools Implementation
 * Following expert implementation plan with Tool base classes
 */

// Import would be: import { Tool } from "@langchain/core/tools";
// For now, we'll create a compatible structure

class LangChainTool {
  constructor() {
    this.name = "";
    this.description = "";
    this.schema = {};
  }

  async call(input) {
    return await this._call(input);
  }

  async _call(input) {
    throw new Error("Must implement _call method");
  }
}

class ReadRangeTool extends LangChainTool {
  constructor() {
    super();
    this.name = "read_range";
    this.description = "Read values or formulas from a range in the active Excel workbook. Useful for fetching financial data like revenue projections.";
    this.schema = {
      type: "object",
      properties: { 
        sheetName: { type: "string" }, 
        range: { type: "string", description: "e.g., 'A1:B10'" } 
      },
      required: ["sheetName", "range"]
    };
  }

  async _call(input) {
    console.log(`🔍 ReadRangeTool: Reading ${input.range} from ${input.sheetName}`);
    
    let result;
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(input.sheetName);
        const range = sheet.getRange(input.range);
        range.load(["values", "formulas", "numberFormat", "address"]);
        await context.sync();
        
        result = { 
          values: range.values, 
          formulas: range.formulas,
          numberFormat: range.numberFormat,
          address: range.address,
          sheetName: input.sheetName
        };
      });
      
      return JSON.stringify({
        success: true,
        data: result,
        summary: `Successfully read range ${input.range} from ${input.sheetName}. Found ${result.values.length} rows.`
      });
      
    } catch (error) {
      return JSON.stringify({
        success: false,
        error: error.message,
        suggestion: `Check that sheet '${input.sheetName}' exists and range '${input.range}' is valid.`
      });
    }
  }
}

class WriteRangeTool extends LangChainTool {
  constructor() {
    super();
    this.name = "write_range";
    this.description = "Write values or formulas to a range. Use for updating financial assumptions, e.g., changing discount rates.";
    this.schema = {
      type: "object",
      properties: { 
        sheetName: { type: "string" }, 
        range: { type: "string" }, 
        values: { type: "array", items: { type: "array" } } 
      },
      required: ["sheetName", "range", "values"]
    };
  }

  async _call(input) {
    console.log(`✍️ WriteRangeTool: Writing to ${input.range} on ${input.sheetName}`);
    
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(input.sheetName);
        const range = sheet.getRange(input.range);
        range.values = input.values;
        await context.sync();
      });
      
      return JSON.stringify({
        success: true,
        message: `Successfully updated range ${input.range} on ${input.sheetName}`,
        rowsUpdated: input.values.length,
        columnsUpdated: input.values[0] ? input.values[0].length : 0
      });
      
    } catch (error) {
      return JSON.stringify({
        success: false,
        error: error.message,
        suggestion: `Ensure the range ${input.range} exists and values array matches the range size.`
      });
    }
  }
}

class EvaluateFinancialFormulaTool extends LangChainTool {
  constructor() {
    super();
    this.name = "evaluate_financial_formula";
    this.description = "Evaluate Excel formulas for financial analysis, e.g., NPV, IRR on data ranges.";
    this.schema = {
      type: "object",
      properties: { 
        formula: { type: "string", description: "e.g., 'NPV(0.1, B2:B10)' or 'IRR(B2:B10)'" }, 
        sheetName: { type: "string", description: "Sheet to evaluate formula on (optional)" },
        tempCell: { type: "string", description: "Temporary cell for evaluation (default: A1)", optional: true }
      },
      required: ["formula"]
    };
  }

  async _call(input) {
    console.log(`📊 EvaluateFinancialFormulaTool: Evaluating ${input.formula}`);
    
    const sheetName = input.sheetName || "Sheet1";
    const tempCell = input.tempCell || "A1";
    let result;
    
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        
        if (sheet.isNullObject) {
          throw new Error(`Sheet '${sheetName}' does not exist`);
        }
        
        const evalRange = sheet.getRange(tempCell);
        
        // Add = sign if not present
        const formulaToEvaluate = input.formula.startsWith('=') ? input.formula : `=${input.formula}`;
        
        evalRange.formulas = [[formulaToEvaluate]];
        evalRange.load(["values", "text"]);
        await context.sync();
        
        result = {
          value: evalRange.values[0][0],
          text: evalRange.text[0][0],
          formula: formulaToEvaluate
        };
        
        // Clean up - clear the temp cell
        evalRange.clear();
        await context.sync();
      });
      
      // Format the result based on formula type
      let formattedResult = result.value;
      const formulaLower = input.formula.toLowerCase();
      
      if (formulaLower.includes('irr')) {
        formattedResult = `${(result.value * 100).toFixed(2)}%`;
      } else if (formulaLower.includes('npv')) {
        formattedResult = `$${result.value.toLocaleString()}`;
      }
      
      return JSON.stringify({
        success: true,
        rawValue: result.value,
        formattedValue: formattedResult,
        formula: result.formula,
        interpretation: this.interpretFinancialResult(input.formula, result.value)
      });
      
    } catch (error) {
      return JSON.stringify({
        success: false,
        error: error.message,
        suggestion: `Check that the formula '${input.formula}' is valid Excel syntax and references exist.`
      });
    }
  }

  interpretFinancialResult(formula, value) {
    const formulaLower = formula.toLowerCase();
    
    if (formulaLower.includes('irr')) {
      if (value > 0.2) return "Excellent return - above 20%";
      if (value > 0.15) return "Good return - above 15%";
      if (value > 0.1) return "Moderate return - above 10%";
      return "Below target return - consider optimization";
    }
    
    if (formulaLower.includes('npv')) {
      if (value > 0) return "Positive NPV - value-creating investment";
      return "Negative NPV - value-destroying investment";
    }
    
    if (formulaLower.includes('pv') || formulaLower.includes('fv')) {
      return `Present/Future value calculation result: ${value}`;
    }
    
    return "Financial calculation completed";
  }
}

class FindFinancialMetricTool extends LangChainTool {
  constructor() {
    super();
    this.name = "find_financial_metric";
    this.description = "Search for and locate financial metrics like IRR, MOIC, Revenue in the workbook with precise cell locations.";
    this.schema = {
      type: "object",
      properties: { 
        metricName: { type: "string", description: "Name of the metric to find (e.g., 'IRR', 'MOIC', 'Revenue')" },
        searchAllSheets: { type: "boolean", description: "Whether to search all sheets (default: true)", default: true }
      },
      required: ["metricName"]
    };
  }

  async _call(input) {
    console.log(`🎯 FindFinancialMetricTool: Searching for ${input.metricName}`);
    
    // Use the existing AccurateExcelValueFinder
    if (!window.langChainExcelTools || !window.langChainExcelTools.valueFinder) {
      return JSON.stringify({
        success: false,
        error: "AccurateExcelValueFinder not available"
      });
    }
    
    try {
      const metrics = await window.langChainExcelTools.valueFinder.findAllFinancialMetrics();
      const foundMetric = metrics[input.metricName];
      
      if (foundMetric) {
        return JSON.stringify({
          success: true,
          metric: input.metricName,
          value: foundMetric.value,
          rawValue: foundMetric.rawValue,
          location: foundMetric.location,
          formula: foundMetric.formula,
          confidence: "high"
        });
      } else {
        return JSON.stringify({
          success: false,
          message: `Metric '${input.metricName}' not found in workbook`,
          availableMetrics: Object.keys(metrics),
          suggestion: "Try searching for one of the available metrics listed above"
        });
      }
      
    } catch (error) {
      return JSON.stringify({
        success: false,
        error: error.message
      });
    }
  }
}

class SmartCellFormattingTool extends LangChainTool {
  constructor() {
    super();
    this.name = "smart_cell_formatting";
    this.description = "Intelligently find and format cells based on their content or labels, not just selected cells.";
    this.schema = {
      type: "object",
      properties: { 
        searchTerm: { type: "string", description: "What to search for (e.g., 'unlevered IRR', 'MOIC', 'revenue')" },
        formatType: { type: "string", enum: ["color", "bold", "italic", "border"], description: "Type of formatting to apply" },
        formatValue: { type: "string", description: "Format value (e.g., 'green', 'red', 'bold')" },
        searchAllSheets: { type: "boolean", description: "Whether to search all sheets", default: true }
      },
      required: ["searchTerm", "formatType", "formatValue"]
    };
  }

  async _call(input) {
    console.log(`🎨 SmartCellFormattingTool: Finding and formatting '${input.searchTerm}'`);
    
    try {
      // Step 1: Find the cell(s) containing the search term
      const foundCells = await this.searchForCells(input.searchTerm, input.searchAllSheets);
      
      if (foundCells.length === 0) {
        return JSON.stringify({
          success: false,
          message: `No cells found containing '${input.searchTerm}'`,
          suggestion: "Try a different search term or check spelling"
        });
      }
      
      // Step 2: Apply formatting to found cells
      const formattingResults = [];
      
      for (const cellInfo of foundCells) {
        try {
          await this.formatCell(cellInfo.address, input.formatType, input.formatValue);
          formattingResults.push({
            address: cellInfo.address,
            sheet: cellInfo.sheet,
            content: cellInfo.content,
            success: true
          });
        } catch (error) {
          formattingResults.push({
            address: cellInfo.address,
            sheet: cellInfo.sheet,
            error: error.message,
            success: false
          });
        }
      }
      
      const successCount = formattingResults.filter(r => r.success).length;
      
      return JSON.stringify({
        success: successCount > 0,
        searchTerm: input.searchTerm,
        cellsFound: foundCells.length,
        cellsFormatted: successCount,
        results: formattingResults,
        message: `Found ${foundCells.length} cells matching '${input.searchTerm}' and successfully formatted ${successCount} of them`
      });
      
    } catch (error) {
      return JSON.stringify({
        success: false,
        error: error.message
      });
    }
  }

  async searchForCells(searchTerm, searchAllSheets = true) {
    const foundCells = [];
    const searchLower = searchTerm.toLowerCase();
    
    await Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load('items');
      await context.sync();
      
      for (const worksheet of worksheets.items) {
        worksheet.load('name');
        const usedRange = worksheet.getUsedRangeOrNullObject();
        usedRange.load(['values', 'address']);
        
        await context.sync();
        
        if (!usedRange.isNullObject && usedRange.values) {
          for (let row = 0; row < usedRange.values.length; row++) {
            for (let col = 0; col < usedRange.values[row].length; col++) {
              const cellValue = String(usedRange.values[row][col]).toLowerCase();
              
              // Check if cell contains the search term
              if (cellValue.includes(searchLower)) {
                const cellAddress = this.getCellAddress(row, col);
                foundCells.push({
                  address: `${worksheet.name}!${cellAddress}`,
                  sheet: worksheet.name,
                  content: usedRange.values[row][col],
                  row: row,
                  col: col
                });
              }
              
              // Also check adjacent cells for values (in case it's a label)
              if (cellValue.includes(searchLower) && col + 1 < usedRange.values[row].length) {
                const valueCell = usedRange.values[row][col + 1];
                if (valueCell !== null && valueCell !== undefined && valueCell !== '') {
                  const valueCellAddress = this.getCellAddress(row, col + 1);
                  foundCells.push({
                    address: `${worksheet.name}!${valueCellAddress}`,
                    sheet: worksheet.name,
                    content: valueCell,
                    row: row,
                    col: col + 1,
                    isValue: true,
                    label: usedRange.values[row][col]
                  });
                }
              }
            }
          }
        }
        
        if (!searchAllSheets) break;
      }
    });
    
    return foundCells;
  }

  async formatCell(cellAddress, formatType, formatValue) {
    await Excel.run(async (context) => {
      let range;
      
      if (cellAddress.includes('!')) {
        const [sheetName, rangeAddr] = cellAddress.split('!');
        const sheet = context.workbook.worksheets.getItem(sheetName);
        range = sheet.getRange(rangeAddr);
      } else {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        range = sheet.getRange(cellAddress);
      }
      
      // Apply formatting based on type
      switch (formatType.toLowerCase()) {
        case 'color':
          range.format.fill.color = this.getColorCode(formatValue);
          break;
        case 'bold':
          range.format.font.bold = formatValue.toLowerCase() === 'true' || formatValue.toLowerCase() === 'bold';
          break;
        case 'italic':
          range.format.font.italic = formatValue.toLowerCase() === 'true' || formatValue.toLowerCase() === 'italic';
          break;
        case 'border':
          range.format.borders.getItem('EdgeTop').style = 'Continuous';
          range.format.borders.getItem('EdgeBottom').style = 'Continuous';
          range.format.borders.getItem('EdgeLeft').style = 'Continuous';
          range.format.borders.getItem('EdgeRight').style = 'Continuous';
          break;
      }
      
      await context.sync();
    });
  }

  getColorCode(colorName) {
    const colorMap = {
      'red': '#FF0000',
      'green': '#00FF00',
      'blue': '#0000FF',
      'yellow': '#FFFF00',
      'orange': '#FFA500',
      'purple': '#800080',
      'pink': '#FFC0CB',
      'light green': '#90EE90',
      'light blue': '#ADD8E6',
      'light gray': '#D3D3D3',
      'dark gray': '#A9A9A9'
    };
    
    return colorMap[colorName.toLowerCase()] || colorName;
  }

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

// Export all tools
const excelTools = [
  new ReadRangeTool(),
  new WriteRangeTool(), 
  new EvaluateFinancialFormulaTool(),
  new FindFinancialMetricTool(),
  new SmartCellFormattingTool()
];

// Initialize globally
if (typeof window !== 'undefined') {
  window.LangChainProperTools = {
    ReadRangeTool,
    WriteRangeTool,
    EvaluateFinancialFormulaTool,
    FindFinancialMetricTool,
    SmartCellFormattingTool,
    excelTools
  };
  
  console.log('✅ LangChain Proper Tools initialized with', excelTools.length, 'tools');
}