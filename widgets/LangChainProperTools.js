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

// Export all tools
const excelTools = [
  new ReadRangeTool(),
  new WriteRangeTool(), 
  new EvaluateFinancialFormulaTool(),
  new FindFinancialMetricTool()
];

// Initialize globally
if (typeof window !== 'undefined') {
  window.LangChainProperTools = {
    ReadRangeTool,
    WriteRangeTool,
    EvaluateFinancialFormulaTool,
    FindFinancialMetricTool,
    excelTools
  };
  
  console.log('✅ LangChain Proper Tools initialized with', excelTools.length, 'tools');
}