/**
 * Excel Tool Library for Investment Banking Add-in
 * Extracted patterns from haris-musa/excel-mcp-server
 * Adapted for Office.js and browser environment
 */

// ==========================================
// 1. TOOL STRUCTURE & FUNCTION SIGNATURES
// ==========================================

/**
 * Base class for all Excel tools following MCP server patterns
 */
class ExcelToolBase {
  constructor(name, description) {
    this.name = name;
    this.description = description;
    this.category = 'general';
  }

  /**
   * Validate parameters before execution (MCP pattern)
   */
  async validateParams(params) {
    throw new Error('validateParams must be implemented by subclass');
  }

  /**
   * Execute the tool operation (MCP pattern)
   */
  async execute(params) {
    throw new Error('execute must be implemented by subclass');
  }
}

// ==========================================
// 2. EXTRACTED VALIDATION LOGIC PATTERNS
// ==========================================

class ExcelValidator {
  /**
   * Validate cell reference (extracted from MCP server validation.py)
   */
  static validateCellReference(cellRef) {
    if (!cellRef || typeof cellRef !== 'string') {
      throw new ValidationError('Cell reference must be a non-empty string');
    }
    
    // Regex pattern from MCP server: r'[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?'
    const cellPattern = /^[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?$/;
    if (!cellPattern.test(cellRef)) {
      throw new ValidationError(`Invalid cell reference format: ${cellRef}`);
    }
    
    return true;
  }

  /**
   * Validate Excel formula (extracted from MCP validation patterns)
   */
  static validateFormula(formula) {
    if (!formula || typeof formula !== 'string') {
      throw new ValidationError('Formula must be a non-empty string');
    }

    // Must start with = (MCP server pattern)
    if (!formula.startsWith('=')) {
      throw new ValidationError('Formula must start with =');
    }

    // Check balanced parentheses (MCP server logic)
    let openCount = 0;
    for (const char of formula) {
      if (char === '(') openCount++;
      if (char === ')') openCount--;
      if (openCount < 0) {
        throw new ValidationError('Unbalanced parentheses in formula');
      }
    }
    if (openCount !== 0) {
      throw new ValidationError('Unbalanced parentheses in formula');
    }

    // Block unsafe functions (MCP server security pattern)
    const unsafeFunctions = ['INDIRECT', 'HYPERLINK', 'CALL'];
    const functionPattern = /([A-Z]+)\(/g;
    let match;
    while ((match = functionPattern.exec(formula)) !== null) {
      if (unsafeFunctions.includes(match[1])) {
        throw new ValidationError(`Unsafe function not allowed: ${match[1]}`);
      }
    }

    return true;
  }

  /**
   * Validate range bounds (extracted from MCP server)
   */
  static validateRangeBounds(startCell, endCell) {
    const maxRows = 1048576; // Excel 2007+ limit
    const maxCols = 16384;   // Excel 2007+ limit (XFD)

    // Parse cell coordinates
    const parseCell = (cell) => {
      const match = cell.match(/^([A-Z]+)([0-9]+)$/);
      if (!match) throw new ValidationError(`Invalid cell format: ${cell}`);
      
      const colStr = match[1];
      const row = parseInt(match[2]);
      
      // Convert column letters to number
      let col = 0;
      for (let i = 0; i < colStr.length; i++) {
        col = col * 26 + (colStr.charCodeAt(i) - 64);
      }
      
      return { row, col };
    };

    const start = parseCell(startCell);
    const end = endCell ? parseCell(endCell) : start;

    // Validate bounds
    if (start.row > maxRows || end.row > maxRows) {
      throw new ValidationError(`Row exceeds maximum: ${maxRows}`);
    }
    if (start.col > maxCols || end.col > maxCols) {
      throw new ValidationError(`Column exceeds maximum: ${maxCols}`);
    }

    // Ensure start is before end
    if (start.row > end.row || start.col > end.col) {
      throw new ValidationError('Start cell must be before end cell');
    }

    return true;
  }

  /**
   * Validate sheet name (MCP server pattern)
   */
  static validateSheetName(sheetName) {
    if (!sheetName || typeof sheetName !== 'string') {
      throw new ValidationError('Sheet name must be a non-empty string');
    }
    
    // Excel sheet name restrictions
    const invalidChars = ['/', '\\', '?', '*', '[', ']'];
    for (const char of invalidChars) {
      if (sheetName.includes(char)) {
        throw new ValidationError(`Sheet name contains invalid character: ${char}`);
      }
    }
    
    if (sheetName.length > 31) {
      throw new ValidationError('Sheet name cannot exceed 31 characters');
    }
    
    return true;
  }
}

// ==========================================
// 3. ERROR HANDLING PATTERNS (from exceptions.py)
// ==========================================

/**
 * Base exception class (extracted from MCP server exceptions.py)
 */
class ExcelMCPError extends Error {
  constructor(message, code = 'EXCEL_ERROR') {
    super(message);
    this.name = 'ExcelMCPError';
    this.code = code;
  }
}

// Specific error types from MCP server
class WorkbookError extends ExcelMCPError {
  constructor(message) {
    super(message, 'WORKBOOK_ERROR');
    this.name = 'WorkbookError';
  }
}

class SheetError extends ExcelMCPError {
  constructor(message) {
    super(message, 'SHEET_ERROR');
    this.name = 'SheetError';
  }
}

class DataError extends ExcelMCPError {
  constructor(message) {
    super(message, 'DATA_ERROR');
    this.name = 'DataError';
  }
}

class ValidationError extends ExcelMCPError {
  constructor(message) {
    super(message, 'VALIDATION_ERROR');
    this.name = 'ValidationError';
  }
}

class FormattingError extends ExcelMCPError {
  constructor(message) {
    super(message, 'FORMATTING_ERROR');
    this.name = 'FormattingError';
  }
}

class CalculationError extends ExcelMCPError {
  constructor(message) {
    super(message, 'CALCULATION_ERROR');
    this.name = 'CalculationError';
  }
}

// ==========================================
// 4. TOOL CATEGORIZATION (from MCP server structure)
// ==========================================

/**
 * Data Operations Tools (extracted from data.py patterns)
 */
class ReadDataTool extends ExcelToolBase {
  constructor() {
    super('read_data', 'Read data from Excel range');
    this.category = 'data';
  }

  async validateParams(params) {
    const { sheetName, startCell, endCell } = params;
    
    ExcelValidator.validateSheetName(sheetName);
    ExcelValidator.validateCellReference(startCell);
    
    if (endCell) {
      ExcelValidator.validateCellReference(endCell);
      ExcelValidator.validateRangeBounds(startCell, endCell);
    }
    
    return true;
  }

  async execute(params) {
    const { sheetName, startCell, endCell, previewOnly = false } = params;
    
    try {
      // Validate parameters first (MCP pattern)
      await this.validateParams(params);
      
      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem(sheetName);
        const range = endCell ? 
          worksheet.getRange(`${startCell}:${endCell}`) : 
          worksheet.getRange(startCell);
        
        range.load(['values', 'formulas', 'address']);
        await context.sync();
        
        // Return format matching MCP server pattern
        const data = range.values.map((row, rowIndex) => 
          row.reduce((obj, cell, colIndex) => {
            const colLetter = String.fromCharCode(65 + colIndex);
            obj[`${colLetter}${rowIndex + 1}`] = cell;
            return obj;
          }, {})
        );
        
        return {
          success: true,
          data: previewOnly ? data.slice(0, 10) : data,
          range: range.address,
          rowCount: range.values.length,
          columnCount: range.values[0]?.length || 0
        };
      });
      
    } catch (error) {
      if (error instanceof ValidationError) {
        throw error;
      }
      throw new DataError(`Failed to read data: ${error.message}`);
    }
  }
}

/**
 * Write Data Tool (extracted from MCP server data.py patterns)
 */
class WriteDataTool extends ExcelToolBase {
  constructor() {
    super('write_data', 'Write data to Excel range');
    this.category = 'data';
  }

  async validateParams(params) {
    const { sheetName, startCell, data } = params;
    
    ExcelValidator.validateSheetName(sheetName);
    ExcelValidator.validateCellReference(startCell);
    
    if (!Array.isArray(data)) {
      throw new ValidationError('Data must be an array');
    }
    
    // Validate data structure
    if (data.length > 0 && !Array.isArray(data[0])) {
      throw new ValidationError('Data must be a 2D array');
    }
    
    return true;
  }

  async execute(params) {
    const { sheetName, startCell, data } = params;
    
    try {
      await this.validateParams(params);
      
      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem(sheetName);
        const range = worksheet.getRange(startCell).getResizedRange(
          data.length - 1, 
          data[0]?.length - 1 || 0
        );
        
        range.values = data;
        await context.sync();
        
        return {
          success: true,
          message: `Data written to ${range.address}`,
          range: range.address,
          rowsWritten: data.length,
          columnsWritten: data[0]?.length || 0
        };
      });
      
    } catch (error) {
      if (error instanceof ValidationError) {
        throw error;
      }
      throw new DataError(`Failed to write data: ${error.message}`);
    }
  }
}

/**
 * Formatting Tools (extracted from formatting.py patterns)
 */
class FormatRangeTool extends ExcelToolBase {
  constructor() {
    super('format_range', 'Format Excel cell range');
    this.category = 'formatting';
  }

  async validateParams(params) {
    const { sheetName, startCell, endCell } = params;
    
    ExcelValidator.validateSheetName(sheetName);
    ExcelValidator.validateCellReference(startCell);
    
    if (endCell) {
      ExcelValidator.validateCellReference(endCell);
      ExcelValidator.validateRangeBounds(startCell, endCell);
    }
    
    return true;
  }

  async execute(params) {
    const { 
      sheetName, 
      startCell, 
      endCell,
      bold = false,
      italic = false,
      fontSize = null,
      fontColor = null,
      backgroundColor = null,
      borderStyle = null
    } = params;
    
    try {
      await this.validateParams(params);
      
      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem(sheetName);
        const range = endCell ? 
          worksheet.getRange(`${startCell}:${endCell}`) : 
          worksheet.getRange(startCell);
        
        // Apply formatting (MCP server pattern)
        if (bold) range.format.font.bold = true;
        if (italic) range.format.font.italic = true;
        if (fontSize) range.format.font.size = fontSize;
        if (fontColor) range.format.font.color = fontColor;
        if (backgroundColor) range.format.fill.color = backgroundColor;
        
        if (borderStyle) {
          range.format.borders.getItem('EdgeTop').style = borderStyle;
          range.format.borders.getItem('EdgeBottom').style = borderStyle;
          range.format.borders.getItem('EdgeLeft').style = borderStyle;
          range.format.borders.getItem('EdgeRight').style = borderStyle;
        }
        
        await context.sync();
        
        return {
          success: true,
          message: `Formatting applied to ${range.address}`,
          range: range.address
        };
      });
      
    } catch (error) {
      if (error instanceof ValidationError) {
        throw error;
      }
      throw new FormattingError(`Failed to format range: ${error.message}`);
    }
  }
}

/**
 * Formula Tools (extracted from MCP server calculation patterns)
 */
class ApplyFormulaTool extends ExcelToolBase {
  constructor() {
    super('apply_formula', 'Apply formula to Excel cell');
    this.category = 'formulas';
  }

  async validateParams(params) {
    const { sheetName, cell, formula } = params;
    
    ExcelValidator.validateSheetName(sheetName);
    ExcelValidator.validateCellReference(cell);
    ExcelValidator.validateFormula(formula);
    
    return true;
  }

  async execute(params) {
    const { sheetName, cell, formula } = params;
    
    try {
      await this.validateParams(params);
      
      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem(sheetName);
        const targetCell = worksheet.getRange(cell);
        
        targetCell.formulas = [[formula]];
        targetCell.load(['values', 'formulas']);
        await context.sync();
        
        return {
          success: true,
          message: `Formula applied to ${cell}`,
          cell: cell,
          formula: formula,
          result: targetCell.values[0][0]
        };
      });
      
    } catch (error) {
      if (error instanceof ValidationError) {
        throw error;
      }
      throw new CalculationError(`Failed to apply formula: ${error.message}`);
    }
  }
}

/**
 * Investment Banking Specific Tools (extensions to MCP patterns)
 */
class CalculateIRRTool extends ExcelToolBase {
  constructor() {
    super('calculate_irr', 'Calculate IRR for investment analysis');
    this.category = 'banking';
  }

  async validateParams(params) {
    const { sheetName, cashFlowRange, datesRange } = params;
    
    ExcelValidator.validateSheetName(sheetName);
    ExcelValidator.validateCellReference(cashFlowRange);
    
    if (datesRange) {
      ExcelValidator.validateCellReference(datesRange);
    }
    
    return true;
  }

  async execute(params) {
    const { sheetName, cashFlowRange, datesRange, outputCell } = params;
    
    try {
      await this.validateParams(params);
      
      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem(sheetName);
        
        // Use XIRR if dates provided, otherwise IRR
        const formula = datesRange ? 
          `=XIRR(${cashFlowRange},${datesRange})` : 
          `=IRR(${cashFlowRange})`;
        
        let targetCell;
        if (outputCell) {
          targetCell = worksheet.getRange(outputCell);
        } else {
          // Find empty cell for output
          targetCell = worksheet.getRange('A1'); // Temporary
        }
        
        targetCell.formulas = [[formula]];
        targetCell.load(['values']);
        await context.sync();
        
        const irrValue = targetCell.values[0][0];
        
        return {
          success: true,
          irr: irrValue,
          formatted: typeof irrValue === 'number' ? `${(irrValue * 100).toFixed(2)}%` : 'Error',
          formula: formula,
          outputCell: targetCell.address
        };
      });
      
    } catch (error) {
      if (error instanceof ValidationError) {
        throw error;
      }
      throw new CalculationError(`Failed to calculate IRR: ${error.message}`);
    }
  }
}

// ==========================================
// 5. TOOL REGISTRY (MCP server organization pattern)
// ==========================================

class ExcelToolRegistry {
  constructor() {
    this.tools = new Map();
    this.categories = {
      'data': [],
      'formatting': [],
      'formulas': [],
      'charts': [],
      'banking': []
    };
    
    this.initializeTools();
  }

  initializeTools() {
    // Register tools by category (MCP server pattern)
    const tools = [
      new ReadDataTool(),
      new WriteDataTool(),
      new FormatRangeTool(),
      new ApplyFormulaTool(),
      new CalculateIRRTool()
    ];
    
    tools.forEach(tool => {
      this.tools.set(tool.name, tool);
      this.categories[tool.category].push(tool);
    });
  }

  getTool(name) {
    return this.tools.get(name);
  }

  getToolsByCategory(category) {
    return this.categories[category] || [];
  }

  getAllTools() {
    return Array.from(this.tools.values());
  }

  /**
   * Execute tool with comprehensive error handling (MCP pattern)
   */
  async executeTool(toolName, params) {
    const tool = this.getTool(toolName);
    if (!tool) {
      throw new ExcelMCPError(`Tool not found: ${toolName}`);
    }
    
    try {
      console.log(`Executing tool: ${toolName}`, params);
      const result = await tool.execute(params);
      console.log(`Tool ${toolName} completed successfully`);
      return result;
    } catch (error) {
      console.error(`Tool ${toolName} failed:`, error);
      
      // Return user-friendly error message (MCP pattern)
      return {
        success: false,
        error: error.message,
        code: error.code || 'UNKNOWN_ERROR',
        tool: toolName
      };
    }
  }
}

// ==========================================
// 6. EXPORT FOR USE IN ADD-IN
// ==========================================

// Initialize global registry
window.ExcelToolRegistry = ExcelToolRegistry;
window.ExcelValidator = ExcelValidator;

// Export error classes
window.ExcelMCPError = ExcelMCPError;
window.WorkbookError = WorkbookError;
window.SheetError = SheetError;
window.DataError = DataError;
window.ValidationError = ValidationError;
window.FormattingError = FormattingError;
window.CalculationError = CalculationError;

// Initialize for immediate use
window.excelTools = new ExcelToolRegistry();

console.log('✅ Excel Tool Library initialized with MCP server patterns');
console.log('Available tools:', window.excelTools.getAllTools().map(t => t.name));
console.log('Available categories:', Object.keys(window.excelTools.categories));