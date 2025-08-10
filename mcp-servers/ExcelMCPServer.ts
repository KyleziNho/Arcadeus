/**
 * ExcelMCPServer.ts
 * MCP Server for Excel operations - provides tools for reading, writing, formatting, and calculating
 */

import {
  MCPServerInfo,
  MCPTool,
  MCPToolCall,
  MCPToolResult,
  MCPContent,
  ExcelRange,
  ExcelFormat,
  ExcelConditionalFormat,
  MCPError,
  MCPErrorCode
} from '../mcp-types/interfaces';

import {
  JSONRPCRequest,
  JSONRPCResponse,
  createResponse,
  createNotification,
  createError,
  ErrorCodes
} from '../mcp-types/schemas';

declare const Excel: any;
declare const Office: any;

export class ExcelMCPServer {
  private serverInfo: MCPServerInfo = {
    name: 'excel-operations-server',
    version: '1.0.0',
    description: 'MCP server for Excel operations including read, write, format, and calculate',
    capabilities: {
      tools: { listChanged: true },
      resources: { listChanged: false },
      notifications: true
    }
  };

  private tools: Map<string, MCPTool> = new Map();
  private isInitialized: boolean = false;
  private operationHistory: any[] = [];

  constructor() {
    this.registerAllTools();
  }

  /**
   * Initialize the server and establish Excel context
   */
  async initialize(): Promise<MCPServerInfo> {
    if (!this.isInitialized) {
      await this.ensureExcelContext();
      this.isInitialized = true;
    }
    return this.serverInfo;
  }

  /**
   * Register all Excel operation tools
   */
  private registerAllTools(): void {
    // Reading tools
    this.registerTool({
      name: 'excel/read-range',
      title: 'Read Range',
      description: 'Read values and formulas from an Excel range',
      inputSchema: {
        type: 'object',
        properties: {
          worksheet: { type: 'string', description: 'Worksheet name (optional, uses active if not provided)' },
          range: { type: 'string', description: 'Range address (e.g., A1:B10)' },
          includeFormulas: { type: 'boolean', description: 'Include formulas in response' },
          includeFormat: { type: 'boolean', description: 'Include formatting in response' }
        },
        required: ['range']
      }
    });

    this.registerTool({
      name: 'excel/find-data',
      title: 'Find Data',
      description: 'Search for data patterns across worksheets',
      inputSchema: {
        type: 'object',
        properties: {
          searchPattern: { type: 'string', description: 'Pattern to search for' },
          searchType: { 
            type: 'string', 
            enum: ['value', 'formula', 'both'],
            description: 'What to search in'
          },
          caseSensitive: { type: 'boolean', description: 'Case sensitive search' },
          wholeMatch: { type: 'boolean', description: 'Match whole cell only' }
        },
        required: ['searchPattern']
      }
    });

    this.registerTool({
      name: 'excel/get-financial-metrics',
      title: 'Get Financial Metrics',
      description: 'Extract financial metrics like MOIC, IRR, NPV from the workbook',
      inputSchema: {
        type: 'object',
        properties: {
          metrics: {
            type: 'array',
            items: { type: 'string' },
            description: 'List of metrics to find (e.g., ["MOIC", "IRR", "NPV"])'
          }
        }
      }
    });

    // Writing tools
    this.registerTool({
      name: 'excel/write-value',
      title: 'Write Value',
      description: 'Write a value to a cell or range',
      inputSchema: {
        type: 'object',
        properties: {
          worksheet: { type: 'string', description: 'Worksheet name (optional)' },
          range: { type: 'string', description: 'Range address' },
          value: { description: 'Value to write (can be string, number, boolean, or array for multiple cells)' }
        },
        required: ['range', 'value']
      }
    });

    this.registerTool({
      name: 'excel/write-formula',
      title: 'Write Formula',
      description: 'Write a formula to a cell or range',
      inputSchema: {
        type: 'object',
        properties: {
          worksheet: { type: 'string', description: 'Worksheet name (optional)' },
          range: { type: 'string', description: 'Range address' },
          formula: { type: 'string', description: 'Excel formula (e.g., =SUM(A1:A10))' }
        },
        required: ['range', 'formula']
      }
    });

    // Formatting tools
    this.registerTool({
      name: 'excel/apply-format',
      title: 'Apply Format',
      description: 'Apply formatting to a range (colors, fonts, borders, etc.)',
      inputSchema: {
        type: 'object',
        properties: {
          worksheet: { type: 'string', description: 'Worksheet name (optional)' },
          range: { type: 'string', description: 'Range address' },
          format: {
            type: 'object',
            properties: {
              backgroundColor: { type: 'string', description: 'Background color (hex)' },
              fontColor: { type: 'string', description: 'Font color (hex)' },
              fontSize: { type: 'number', description: 'Font size in points' },
              bold: { type: 'boolean', description: 'Bold text' },
              italic: { type: 'boolean', description: 'Italic text' },
              underline: { type: 'boolean', description: 'Underline text' },
              numberFormat: { type: 'string', description: 'Number format (e.g., "#,##0.00")' }
            }
          }
        },
        required: ['range', 'format']
      }
    });

    this.registerTool({
      name: 'excel/apply-conditional-format',
      title: 'Apply Conditional Format',
      description: 'Apply conditional formatting rules to a range',
      inputSchema: {
        type: 'object',
        properties: {
          worksheet: { type: 'string', description: 'Worksheet name (optional)' },
          range: { type: 'string', description: 'Range address' },
          rule: {
            type: 'object',
            properties: {
              type: { 
                type: 'string',
                enum: ['cellValue', 'colorScale', 'dataBar', 'iconSet'],
                description: 'Type of conditional format'
              },
              operator: { 
                type: 'string',
                enum: ['greaterThan', 'lessThan', 'equal', 'between'],
                description: 'Comparison operator'
              },
              value1: { description: 'First comparison value' },
              value2: { description: 'Second comparison value (for between)' },
              format: { type: 'object', description: 'Format to apply when condition is met' }
            },
            required: ['type']
          }
        },
        required: ['range', 'rule']
      }
    });

    // Calculation tools
    this.registerTool({
      name: 'excel/calculate-irr',
      title: 'Calculate IRR',
      description: 'Calculate Internal Rate of Return for cash flows',
      inputSchema: {
        type: 'object',
        properties: {
          cashFlowRange: { type: 'string', description: 'Range containing cash flows' },
          guess: { type: 'number', description: 'Initial guess for IRR (default: 0.1)' }
        },
        required: ['cashFlowRange']
      }
    });

    this.registerTool({
      name: 'excel/calculate-npv',
      title: 'Calculate NPV',
      description: 'Calculate Net Present Value',
      inputSchema: {
        type: 'object',
        properties: {
          rate: { type: 'number', description: 'Discount rate' },
          cashFlowRange: { type: 'string', description: 'Range containing cash flows' }
        },
        required: ['rate', 'cashFlowRange']
      }
    });

    this.registerTool({
      name: 'excel/calculate-moic',
      title: 'Calculate MOIC',
      description: 'Calculate Multiple on Invested Capital',
      inputSchema: {
        type: 'object',
        properties: {
          totalReturnRange: { type: 'string', description: 'Range or value for total return' },
          investedCapitalRange: { type: 'string', description: 'Range or value for invested capital' }
        },
        required: ['totalReturnRange', 'investedCapitalRange']
      }
    });

    // Chart tools
    this.registerTool({
      name: 'excel/create-chart',
      title: 'Create Chart',
      description: 'Create a chart from data',
      inputSchema: {
        type: 'object',
        properties: {
          dataRange: { type: 'string', description: 'Data range for chart' },
          chartType: { 
            type: 'string',
            enum: ['column', 'line', 'pie', 'bar', 'area', 'scatter'],
            description: 'Type of chart'
          },
          title: { type: 'string', description: 'Chart title' },
          position: { 
            type: 'object',
            properties: {
              top: { type: 'number' },
              left: { type: 'number' },
              width: { type: 'number' },
              height: { type: 'number' }
            }
          }
        },
        required: ['dataRange', 'chartType']
      }
    });
  }

  /**
   * Register a single tool
   */
  private registerTool(tool: MCPTool): void {
    this.tools.set(tool.name, tool);
  }

  /**
   * List all available tools
   */
  async listTools(): Promise<MCPTool[]> {
    return Array.from(this.tools.values());
  }

  /**
   * Execute a tool
   */
  async executeTool(toolCall: MCPToolCall): Promise<MCPToolResult> {
    const tool = this.tools.get(toolCall.name);
    if (!tool) {
      throw new MCPError(
        `Tool not found: ${toolCall.name}`,
        MCPErrorCode.METHOD_NOT_FOUND
      );
    }

    try {
      // Route to appropriate handler based on tool name
      const handler = this.getToolHandler(toolCall.name);
      const result = await handler(toolCall.arguments);
      
      // Record operation for undo/redo
      this.recordOperation(toolCall, result);
      
      return result;
    } catch (error: any) {
      return {
        content: [{
          type: 'error',
          text: error.message || 'Tool execution failed'
        }],
        isError: true,
        errorMessage: error.message
      };
    }
  }

  /**
   * Get the appropriate handler for a tool
   */
  private getToolHandler(toolName: string): (args: any) => Promise<MCPToolResult> {
    const handlers: Record<string, (args: any) => Promise<MCPToolResult>> = {
      'excel/read-range': this.handleReadRange.bind(this),
      'excel/find-data': this.handleFindData.bind(this),
      'excel/get-financial-metrics': this.handleGetFinancialMetrics.bind(this),
      'excel/write-value': this.handleWriteValue.bind(this),
      'excel/write-formula': this.handleWriteFormula.bind(this),
      'excel/apply-format': this.handleApplyFormat.bind(this),
      'excel/apply-conditional-format': this.handleApplyConditionalFormat.bind(this),
      'excel/calculate-irr': this.handleCalculateIRR.bind(this),
      'excel/calculate-npv': this.handleCalculateNPV.bind(this),
      'excel/calculate-moic': this.handleCalculateMOIC.bind(this),
      'excel/create-chart': this.handleCreateChart.bind(this)
    };

    return handlers[toolName] || (() => Promise.reject(new Error('Handler not implemented')));
  }

  // ===== Tool Handlers =====

  private async handleReadRange(args: any): Promise<MCPToolResult> {
    return Excel.run(async (context: any) => {
      const worksheet = args.worksheet 
        ? context.workbook.worksheets.getItem(args.worksheet)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(args.range);
      range.load(['values', 'formulas', 'address', 'rowCount', 'columnCount']);
      
      if (args.includeFormat) {
        range.load(['format']);
      }
      
      await context.sync();
      
      const content: MCPContent[] = [{
        type: 'text',
        text: JSON.stringify({
          address: range.address,
          values: range.values,
          formulas: args.includeFormulas ? range.formulas : undefined,
          rowCount: range.rowCount,
          columnCount: range.columnCount
        }, null, 2)
      }];
      
      return { content };
    });
  }

  private async handleFindData(args: any): Promise<MCPToolResult> {
    return Excel.run(async (context: any) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load('items');
      await context.sync();
      
      const results: any[] = [];
      
      for (const worksheet of worksheets.items) {
        const usedRange = worksheet.getUsedRange();
        usedRange.load(['values', 'formulas', 'address']);
        await context.sync();
        
        if (!usedRange.values) continue;
        
        for (let row = 0; row < usedRange.values.length; row++) {
          for (let col = 0; col < usedRange.values[row].length; col++) {
            const value = usedRange.values[row][col];
            const formula = usedRange.formulas[row][col];
            
            let match = false;
            const searchIn = args.searchType === 'formula' ? formula : 
                           args.searchType === 'both' ? `${value} ${formula}` : value;
            
            if (searchIn) {
              const searchStr = String(searchIn);
              const pattern = String(args.searchPattern);
              
              if (args.wholeMatch) {
                match = args.caseSensitive 
                  ? searchStr === pattern
                  : searchStr.toLowerCase() === pattern.toLowerCase();
              } else {
                match = args.caseSensitive
                  ? searchStr.includes(pattern)
                  : searchStr.toLowerCase().includes(pattern.toLowerCase());
              }
            }
            
            if (match) {
              results.push({
                worksheet: worksheet.name,
                cell: `${String.fromCharCode(65 + col)}${row + 1}`,
                value: value,
                formula: formula
              });
            }
          }
        }
      }
      
      return {
        content: [{
          type: 'text',
          text: JSON.stringify(results, null, 2)
        }]
      };
    });
  }

  private async handleGetFinancialMetrics(args: any): Promise<MCPToolResult> {
    const metrics = args.metrics || ['MOIC', 'IRR', 'NPV'];
    const results: Record<string, any> = {};
    
    await Excel.run(async (context: any) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load('items');
      await context.sync();
      
      for (const metric of metrics) {
        // Search for metric labels and adjacent values
        for (const worksheet of worksheets.items) {
          const usedRange = worksheet.getUsedRange();
          usedRange.load(['values', 'formulas']);
          await context.sync();
          
          if (!usedRange.values) continue;
          
          for (let row = 0; row < usedRange.values.length; row++) {
            for (let col = 0; col < usedRange.values[row].length; col++) {
              const cellValue = String(usedRange.values[row][col] || '');
              
              if (cellValue.toUpperCase().includes(metric.toUpperCase())) {
                // Check adjacent cells for values
                const adjacentCells = [
                  { r: row, c: col + 1 }, // Right
                  { r: row + 1, c: col }, // Below
                  { r: row, c: col - 1 }, // Left
                  { r: row - 1, c: col }  // Above
                ];
                
                for (const adj of adjacentCells) {
                  if (adj.r >= 0 && adj.r < usedRange.values.length &&
                      adj.c >= 0 && adj.c < usedRange.values[adj.r].length) {
                    const adjValue = usedRange.values[adj.r][adj.c];
                    if (typeof adjValue === 'number') {
                      results[metric] = {
                        value: adjValue,
                        formula: usedRange.formulas[adj.r][adj.c],
                        location: {
                          worksheet: worksheet.name,
                          cell: `${String.fromCharCode(65 + adj.c)}${adj.r + 1}`
                        }
                      };
                      break;
                    }
                  }
                }
              }
            }
          }
        }
      }
    });
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify(results, null, 2)
      }]
    };
  }

  private async handleWriteValue(args: any): Promise<MCPToolResult> {
    return Excel.run(async (context: any) => {
      const worksheet = args.worksheet 
        ? context.workbook.worksheets.getItem(args.worksheet)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(args.range);
      range.values = Array.isArray(args.value) ? args.value : [[args.value]];
      
      await context.sync();
      
      return {
        content: [{
          type: 'text',
          text: `Successfully wrote value to ${args.range}`
        }]
      };
    });
  }

  private async handleWriteFormula(args: any): Promise<MCPToolResult> {
    return Excel.run(async (context: any) => {
      const worksheet = args.worksheet 
        ? context.workbook.worksheets.getItem(args.worksheet)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(args.range);
      range.formulas = [[args.formula]];
      
      await context.sync();
      
      return {
        content: [{
          type: 'text',
          text: `Successfully wrote formula to ${args.range}`
        }]
      };
    });
  }

  private async handleApplyFormat(args: any): Promise<MCPToolResult> {
    return Excel.run(async (context: any) => {
      const worksheet = args.worksheet 
        ? context.workbook.worksheets.getItem(args.worksheet)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(args.range);
      const format = args.format;
      
      if (format.backgroundColor) {
        range.format.fill.color = format.backgroundColor;
      }
      if (format.fontColor) {
        range.format.font.color = format.fontColor;
      }
      if (format.fontSize) {
        range.format.font.size = format.fontSize;
      }
      if (format.bold !== undefined) {
        range.format.font.bold = format.bold;
      }
      if (format.italic !== undefined) {
        range.format.font.italic = format.italic;
      }
      if (format.underline !== undefined) {
        range.format.font.underline = format.underline ? 'Single' : 'None';
      }
      if (format.numberFormat) {
        range.numberFormat = [[format.numberFormat]];
      }
      
      await context.sync();
      
      return {
        content: [{
          type: 'text',
          text: `Successfully applied formatting to ${args.range}`
        }]
      };
    });
  }

  private async handleApplyConditionalFormat(args: any): Promise<MCPToolResult> {
    return Excel.run(async (context: any) => {
      const worksheet = args.worksheet 
        ? context.workbook.worksheets.getItem(args.worksheet)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(args.range);
      const rule = args.rule;
      
      const conditionalFormat = range.conditionalFormats.add(
        Excel.ConditionalFormatType[rule.type as keyof typeof Excel.ConditionalFormatType]
      );
      
      if (rule.type === 'cellValue' && conditionalFormat.cellValue) {
        conditionalFormat.cellValue.format.fill.color = rule.format?.backgroundColor || '#FFFF00';
        conditionalFormat.cellValue.rule = {
          formula1: String(rule.value1),
          formula2: rule.value2 ? String(rule.value2) : undefined,
          operator: Excel.ConditionalCellValueOperator[rule.operator as keyof typeof Excel.ConditionalCellValueOperator]
        };
      }
      
      await context.sync();
      
      return {
        content: [{
          type: 'text',
          text: `Successfully applied conditional formatting to ${args.range}`
        }]
      };
    });
  }

  private async handleCalculateIRR(args: any): Promise<MCPToolResult> {
    return Excel.run(async (context: any) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const cashFlowRange = worksheet.getRange(args.cashFlowRange);
      cashFlowRange.load('values');
      
      await context.sync();
      
      // Create a temporary cell for IRR calculation
      const tempCell = worksheet.getRange('Z9999');
      tempCell.formulas = [[`=IRR(${args.cashFlowRange}${args.guess ? ',' + args.guess : ''})`]];
      tempCell.load('values');
      
      await context.sync();
      
      const irrValue = tempCell.values[0][0];
      
      // Clear temp cell
      tempCell.clear();
      await context.sync();
      
      return {
        content: [{
          type: 'text',
          text: `IRR: ${(irrValue * 100).toFixed(2)}%`
        }]
      };
    });
  }

  private async handleCalculateNPV(args: any): Promise<MCPToolResult> {
    return Excel.run(async (context: any) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Create a temporary cell for NPV calculation
      const tempCell = worksheet.getRange('Z9998');
      tempCell.formulas = [[`=NPV(${args.rate},${args.cashFlowRange})`]];
      tempCell.load('values');
      
      await context.sync();
      
      const npvValue = tempCell.values[0][0];
      
      // Clear temp cell
      tempCell.clear();
      await context.sync();
      
      return {
        content: [{
          type: 'text',
          text: `NPV: $${npvValue.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
        }]
      };
    });
  }

  private async handleCalculateMOIC(args: any): Promise<MCPToolResult> {
    return Excel.run(async (context: any) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Get total return
      const returnRange = worksheet.getRange(args.totalReturnRange);
      returnRange.load('values');
      
      // Get invested capital
      const capitalRange = worksheet.getRange(args.investedCapitalRange);
      capitalRange.load('values');
      
      await context.sync();
      
      const totalReturn = this.sumValues(returnRange.values);
      const investedCapital = this.sumValues(capitalRange.values);
      
      const moic = totalReturn / investedCapital;
      
      return {
        content: [{
          type: 'text',
          text: `MOIC: ${moic.toFixed(2)}x\nTotal Return: $${totalReturn.toLocaleString()}\nInvested Capital: $${investedCapital.toLocaleString()}`
        }]
      };
    });
  }

  private async handleCreateChart(args: any): Promise<MCPToolResult> {
    return Excel.run(async (context: any) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const dataRange = worksheet.getRange(args.dataRange);
      
      const chart = worksheet.charts.add(
        Excel.ChartType[args.chartType as keyof typeof Excel.ChartType],
        dataRange,
        Excel.ChartSeriesBy.auto
      );
      
      if (args.title) {
        chart.title.text = args.title;
      }
      
      if (args.position) {
        chart.top = args.position.top || 100;
        chart.left = args.position.left || 100;
        chart.width = args.position.width || 400;
        chart.height = args.position.height || 300;
      }
      
      await context.sync();
      
      return {
        content: [{
          type: 'text',
          text: `Successfully created ${args.chartType} chart`
        }]
      };
    });
  }

  // ===== Helper Methods =====

  private async ensureExcelContext(): Promise<void> {
    if (typeof Excel === 'undefined') {
      throw new MCPError(
        'Excel context not available',
        MCPErrorCode.EXCEL_NOT_READY
      );
    }
  }

  private sumValues(values: any[][]): number {
    let sum = 0;
    for (const row of values) {
      for (const cell of row) {
        if (typeof cell === 'number') {
          sum += cell;
        }
      }
    }
    return sum;
  }

  private recordOperation(toolCall: MCPToolCall, result: MCPToolResult): void {
    this.operationHistory.push({
      timestamp: new Date(),
      tool: toolCall.name,
      arguments: toolCall.arguments,
      result: result,
      id: `op_${Date.now()}`
    });
    
    // Limit history size
    if (this.operationHistory.length > 100) {
      this.operationHistory.shift();
    }
  }

  /**
   * Send a notification to connected clients
   */
  async sendNotification(method: string, params?: any): Promise<void> {
    const notification = createNotification(method, params);
    // This would be sent through the transport layer
    console.log('Sending notification:', notification);
  }

  /**
   * Handle incoming request
   */
  async handleRequest(request: JSONRPCRequest): Promise<JSONRPCResponse> {
    try {
      switch (request.method) {
        case 'initialize':
          const serverInfo = await this.initialize();
          return createResponse(request.id!, {
            protocolVersion: '2025-06-18',
            capabilities: serverInfo.capabilities,
            serverInfo: {
              name: serverInfo.name,
              version: serverInfo.version
            }
          });
          
        case 'tools/list':
          const tools = await this.listTools();
          return createResponse(request.id!, { tools });
          
        case 'tools/call':
          const result = await this.executeTool(request.params);
          return createResponse(request.id!, result);
          
        default:
          return createResponse(
            request.id!,
            undefined,
            createError(ErrorCodes.METHOD_NOT_FOUND, `Method not found: ${request.method}`)
          );
      }
    } catch (error: any) {
      return createResponse(
        request.id!,
        undefined,
        createError(ErrorCodes.INTERNAL_ERROR, error.message)
      );
    }
  }
}

// Export for use in the application
export default ExcelMCPServer;