/**
 * LangChain Excel Tools for Deep Agents
 * Direct copies of MCP server tools converted to LangChain format
 * Based on haris-musa/excel-mcp-server tool signatures
 */

class LangChainExcelTools {
  constructor() {
    this.tools = [];
    this.initialize();
  }

  async initialize() {
    console.log('🛠️ Initializing LangChain Excel Tools with MCP server patterns...');
    
    // Create tool definitions copied from MCP server
    this.createMCPTools();
    console.log('✅ LangChain Excel Tools ready with MCP patterns');
  }

  /**
   * Create MCP server tools for LangChain Deep Agent
   * Direct copies from haris-musa/excel-mcp-server
   */
  createMCPTools() {
    
    // ==========================================
    // DATA OPERATIONS (from MCP data.py)
    // ==========================================
    
    // Read Excel Range - Direct copy of MCP server signature
    this.readExcelRangeTool = {
      name: 'read_excel_range',
      description: 'Read data from Excel range. Direct copy of MCP server read_excel_range function signature.',
      schema: {
        type: 'object',
        properties: {
          sheetName: { type: 'string', description: 'Name of the worksheet' },
          startCell: { type: 'string', default: 'A1', description: 'Starting cell (e.g., A1)' },
          endCell: { type: 'string', description: 'Ending cell (e.g., B10)' },
          previewOnly: { type: 'boolean', default: false, description: 'Return only first 10 rows' }
        },
        required: ['sheetName']
      },
      execute: async ({ sheetName, startCell = 'A1', endCell, previewOnly = false }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheetName);
            const range = endCell ? 
              worksheet.getRange(`${startCell}:${endCell}`) : 
              worksheet.getRange(startCell);
            
            range.load(['values', 'formulas', 'address']);
            await context.sync();
            
            // Return format matching MCP server
            const data = range.values.map((row, rowIndex) => 
              row.reduce((obj, cell, colIndex) => {
                const colLetter = String.fromCharCode(65 + colIndex);
                obj[`${colLetter}${rowIndex + 1}`] = cell;
                return obj;
              }, {})
            );
            
            const result = {
              success: true,
              data: previewOnly ? data.slice(0, 10) : data,
              range: range.address,
              rowCount: range.values.length,
              columnCount: range.values[0]?.length || 0
            };
            
            return JSON.stringify(result);
          });
          
        } catch (error) {
          return JSON.stringify({
            success: false,
            error: `Failed to read Excel range: ${error.message}`
          });
        }
      }
    };

    // Write Data - Direct copy of MCP server signature
    this.writeDataTool = {
      name: 'write_data',
      description: 'Write data to Excel sheet. Direct copy of MCP server write_data function signature.',
      schema: {
        type: 'object',
        properties: {
          sheetName: { type: 'string', description: 'Name of the worksheet' },
          data: { type: 'array', items: { type: 'array' }, description: '2D array of data to write' },
          startCell: { type: 'string', default: 'A1', description: 'Starting cell to write data' }
        },
        required: ['sheetName', 'data']
      },
      execute: async ({ sheetName, data, startCell = 'A1' }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheetName);
            const range = worksheet.getRange(startCell).getResizedRange(
              data.length - 1, 
              data[0]?.length - 1 || 0
            );
            
            range.values = data;
            await context.sync();
            
            const result = {
              success: true,
              message: `Data written to ${range.address}`,
              range: range.address,
              rowsWritten: data.length,
              columnsWritten: data[0]?.length || 0
            };
            
            return JSON.stringify(result);
          });
          
        } catch (error) {
          return JSON.stringify({
            success: false,
            error: `Failed to write data: ${error.message}`
          });
        }
      }
    };

    // ==========================================
    // FORMATTING OPERATIONS (from MCP formatting.py)
    // ==========================================
    
    // Format Range - Direct copy of MCP server signature
    this.formatRangeTool = {
      name: 'format_range',
      description: 'Format Excel cell range. Direct copy of MCP server format_range function signature.',
      schema: {
        type: 'object',
        properties: {
          sheetName: { type: 'string', description: 'Name of the worksheet' },
          startCell: { type: 'string', description: 'Starting cell to format' },
          endCell: { type: 'string', description: 'Ending cell to format' },
          bold: { type: 'boolean', default: false, description: 'Apply bold formatting' },
          italic: { type: 'boolean', default: false, description: 'Apply italic formatting' },
          fontSize: { type: 'number', description: 'Font size' },
          fontColor: { type: 'string', description: 'Font color (hex code)' },
          backgroundColor: { type: 'string', description: 'Background color (hex code)' },
          borderStyle: { type: 'string', description: 'Border style' }
        },
        required: ['sheetName', 'startCell']
      },
      execute: async ({ 
        sheetName, 
        startCell, 
        endCell,
        bold = false,
        italic = false,
        fontSize = null,
        fontColor = null,
        backgroundColor = null,
        borderStyle = null
      }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheetName);
            const range = endCell ? 
              worksheet.getRange(`${startCell}:${endCell}`) : 
              worksheet.getRange(startCell);
            
            // Apply formatting using MCP server pattern
            if (bold) range.format.font.bold = true;
            if (italic) range.format.font.italic = true;
            if (fontSize) range.format.font.size = fontSize;
            if (fontColor) range.format.font.color = fontColor;
            if (backgroundColor) range.format.fill.color = backgroundColor;
            
            if (borderStyle) {
              const borderItems = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'];
              borderItems.forEach(edge => {
                range.format.borders.getItem(edge).style = borderStyle;
              });
            }
            
            await context.sync();
            
            const result = {
              success: true,
              message: `Formatting applied to ${range.address}`,
              range: range.address
            };
            
            return JSON.stringify(result);
          });
          
        } catch (error) {
          return JSON.stringify({
            success: false,
            error: `Failed to format range: ${error.message}`
          });
        }
      }
    };

    // ==========================================
    // FORMULA OPERATIONS (from MCP calculations.py)
    // ==========================================
    
    // Apply Formula - Direct copy of MCP server signature
    this.applyFormulaTool = {
      name: 'apply_formula',
      description: 'Apply formula to Excel cell. Direct copy of MCP server apply_formula function signature.',
      schema: {
        type: 'object',
        properties: {
          sheetName: { type: 'string', description: 'Name of the worksheet' },
          cell: { type: 'string', description: 'Cell to apply formula (e.g., A1)' },
          formula: { type: 'string', description: 'Excel formula to apply (must start with =)' }
        },
        required: ['sheetName', 'cell', 'formula']
      },
      execute: async ({ sheetName, cell, formula }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheetName);
            const targetCell = worksheet.getRange(cell);
            
            targetCell.formulas = [[formula]];
            targetCell.load(['values', 'formulas']);
            await context.sync();
            
            const result = {
              success: true,
              message: `Formula applied to ${cell}`,
              cell: cell,
              formula: formula,
              result: targetCell.values[0][0]
            };
            
            return JSON.stringify(result);
          });
          
        } catch (error) {
          return JSON.stringify({
            success: false,
            error: `Failed to apply formula: ${error.message}`
          });
        }
      }
    };

    // Validate Formula - Direct copy of MCP server signature
    this.validateFormulaTool = {
      name: 'validate_formula',
      description: 'Validate Excel formula syntax. Direct copy of MCP server validate_formula function signature.',
      schema: {
        type: 'object',
        properties: {
          formula: { type: 'string', description: 'Excel formula to validate' }
        },
        required: ['formula']
      },
      execute: async ({ formula }) => {
        try {
          // MCP server validation logic
          if (!formula || typeof formula !== 'string') {
            return JSON.stringify({
              success: false,
              error: 'Formula must be a non-empty string'
            });
          }

          if (!formula.startsWith('=')) {
            return JSON.stringify({
              success: false,
              error: 'Formula must start with ='
            });
          }

          // Check balanced parentheses (MCP server logic)
          let openCount = 0;
          for (const char of formula) {
            if (char === '(') openCount++;
            if (char === ')') openCount--;
            if (openCount < 0) {
              return JSON.stringify({
                success: false,
                error: 'Unbalanced parentheses in formula'
              });
            }
          }
          if (openCount !== 0) {
            return JSON.stringify({
              success: false,
              error: 'Unbalanced parentheses in formula'
            });
          }

          // Block unsafe functions (MCP server security)
          const unsafeFunctions = ['INDIRECT', 'HYPERLINK', 'CALL'];
          const functionPattern = /([A-Z]+)\(/g;
          let match;
          while ((match = functionPattern.exec(formula)) !== null) {
            if (unsafeFunctions.includes(match[1])) {
              return JSON.stringify({
                success: false,
                error: `Unsafe function not allowed: ${match[1]}`
              });
            }
          }

          return JSON.stringify({
            success: true,
            message: 'Formula validation passed',
            formula: formula
          });
          
        } catch (error) {
          return JSON.stringify({
            success: false,
            error: `Formula validation failed: ${error.message}`
          });
        }
      }
    };

    // ==========================================
    // WORKSHEET OPERATIONS (from MCP sheet.py)
    // ==========================================

    // Create Worksheet - Direct copy of MCP server signature
    this.createWorksheetTool = {
      name: 'create_worksheet',
      description: 'Create new worksheet. Direct copy of MCP server create_worksheet function signature.',
      schema: {
        type: 'object',
        properties: {
          sheetName: { type: 'string', description: 'Name of the new worksheet' }
        },
        required: ['sheetName']
      },
      execute: async ({ sheetName }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            const newSheet = worksheets.add(sheetName);
            newSheet.load(['name']);
            await context.sync();
            
            const result = {
              success: true,
              message: `Worksheet '${newSheet.name}' created successfully`,
              sheetName: newSheet.name
            };
            
            return JSON.stringify(result);
          });
          
        } catch (error) {
          return JSON.stringify({
            success: false,
            error: `Failed to create worksheet: ${error.message}`
          });
        }
      }
    };

    // ==========================================
    // CHART OPERATIONS (from MCP chart.py)
    // ==========================================

    // Create Chart - Direct copy of MCP server signature
    this.createChartTool = {
      name: 'create_chart',
      description: 'Create chart in Excel. Direct copy of MCP server create_chart function signature.',
      schema: {
        type: 'object',
        properties: {
          sheetName: { type: 'string', description: 'Name of the worksheet' },
          dataRange: { type: 'string', description: 'Range of data for chart (e.g., A1:B10)' },
          chartType: { type: 'string', description: 'Type of chart (ColumnClustered, Line, Pie, etc.)' },
          targetCell: { type: 'string', description: 'Cell where chart should be placed' },
          title: { type: 'string', default: '', description: 'Chart title' },
          xAxis: { type: 'string', default: '', description: 'X-axis label' },
          yAxis: { type: 'string', default: '', description: 'Y-axis label' }
        },
        required: ['sheetName', 'dataRange', 'chartType', 'targetCell']
      },
      execute: async ({ 
        sheetName, 
        dataRange, 
        chartType, 
        targetCell, 
        title = '',
        xAxis = '',
        yAxis = ''
      }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheetName);
            const sourceRange = worksheet.getRange(dataRange);
            
            const chart = worksheet.charts.add(chartType, sourceRange);
            chart.setPosition(targetCell);
            
            if (title) chart.title.text = title;
            if (xAxis) chart.axes.categoryAxis.title.text = xAxis;
            if (yAxis) chart.axes.valueAxis.title.text = yAxis;
            
            await context.sync();
            
            const result = {
              success: true,
              message: `Chart created at ${targetCell}`,
              chartType: chartType,
              dataRange: dataRange,
              location: targetCell
            };
            
            return JSON.stringify(result);
          });
          
        } catch (error) {
          return JSON.stringify({
            success: false,
            error: `Failed to create chart: ${error.message}`
          });
        }
      }
    };

    // ==========================================
    // INVESTMENT BANKING TOOLS (Extensions using MCP patterns)
    // ==========================================

    // Calculate IRR Tool - Banking extension using MCP patterns
    this.calculateIRRTool = {
      name: 'calculate_irr',
      description: 'Calculate IRR for investment analysis. Banking-specific extension using MCP patterns.',
      schema: {
        type: 'object',
        properties: {
          sheetName: { type: 'string', description: 'Name of the worksheet' },
          cashFlowRange: { type: 'string', description: 'Range containing cash flows (e.g., B2:B6)' },
          datesRange: { type: 'string', description: 'Range containing dates (for XIRR)' },
          outputCell: { type: 'string', description: 'Cell to place IRR result' }
        },
        required: ['sheetName', 'cashFlowRange']
      },
      execute: async ({ sheetName, cashFlowRange, datesRange, outputCell }) => {
        try {
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
            
            const result = {
              success: true,
              irr: irrValue,
              formatted: typeof irrValue === 'number' ? `${(irrValue * 100).toFixed(2)}%` : 'Error',
              formula: formula,
              outputCell: targetCell.address,
              interpretation: typeof irrValue === 'number' ? 
                (irrValue > 0.20 ? 'Excellent return' : 
                 irrValue > 0.15 ? 'Good return' : 
                 irrValue > 0.10 ? 'Acceptable return' : 'Poor return') : 'Calculation error'
            };
            
            return JSON.stringify(result);
          });
          
        } catch (error) {
          return JSON.stringify({
            success: false,
            error: `Failed to calculate IRR: ${error.message}`
          });
        }
      }
    };

    // Get Workbook Metadata - Direct copy of MCP server signature
    this.getWorkbookMetadataTool = {
      name: 'get_workbook_metadata',
      description: 'Get workbook metadata. Direct copy of MCP server get_workbook_metadata function signature.',
      schema: {
        type: 'object',
        properties: {}
      },
      execute: async () => {
        try {
          return await Excel.run(async (context) => {
            const workbook = context.workbook;
            const worksheets = workbook.worksheets;
            
            workbook.load(['name']);
            worksheets.load(['items/name', 'items/position']);
            await context.sync();
            
            const result = {
              success: true,
              workbookName: workbook.name || 'Unknown',
              worksheetCount: worksheets.items.length,
              worksheets: worksheets.items.map(ws => ({
                name: ws.name,
                position: ws.position
              }))
            };
            
            return JSON.stringify(result);
          });
          
        } catch (error) {
          return JSON.stringify({
            success: false,
            error: `Failed to get workbook metadata: ${error.message}`
          });
        }
      }
    };

    // Collect all MCP server tools
    this.tools = [
      // Data operations (from MCP data.py)
      this.readExcelRangeTool,
      this.writeDataTool,
      
      // Formatting operations (from MCP formatting.py) 
      this.formatRangeTool,
      
      // Formula operations (from MCP calculations.py)
      this.applyFormulaTool,
      this.validateFormulaTool,
      
      // Worksheet operations (from MCP sheet.py)
      this.createWorksheetTool,
      
      // Chart operations (from MCP chart.py)
      this.createChartTool,
      
      // Banking operations (extension using MCP patterns)
      this.calculateIRRTool,
      
      // Metadata operations (from MCP workbook.py)
      this.getWorkbookMetadataTool
    ];
  }

  /**
   * Get all tools for Deep Agent
   */
  getAllTools() {
    return this.tools;
  }

  /**
   * Format tools for LangChain-style usage
   */
  formatToolsForDeepAgent() {
    return this.tools.map(tool => ({
      name: tool.name,
      description: tool.description,
      schema: tool.schema,
      execute: tool.execute
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
}

// ==========================================
// DEEP AGENT IMPLEMENTATION
// ==========================================

/**
 * JavaScript Deep Agent following the Python deepagents package pattern
 * Components: planning tool, sub agents, file system, detailed prompt
 */
class DeepExcelAgent {
  constructor(apiKey, instructions = '') {
    this.apiKey = apiKey;
    this.instructions = instructions;
    
    // Deep Agent core components
    this.fileSystem = new Map();  // Mock file system
    this.todoList = [];
    this.messages = [];
    
    // Tools from MCP server
    this.excelTools = new LangChainExcelTools();
    this.tools = this.createDeepAgentTools();
    
    // Sub agents
    this.subAgents = this.createSubAgents();
    
    console.log('🧠 Deep Agent initialized with MCP server tools');
  }

  /**
   * Create tools for Deep Agent (including MCP tools + Deep Agent built-ins)
   */
  createDeepAgentTools() {
    const mcpTools = this.excelTools.getAllTools();
    
    // Add Deep Agent built-in tools
    const builtInTools = [
      {
        name: 'todo_write',
        description: 'Create or update task plan. ALWAYS use this first for complex tasks.',
        schema: {
          type: 'object',
          properties: {
            todos: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  id: { type: 'number' },
                  task: { type: 'string' },
                  status: { type: 'string', enum: ['pending', 'in_progress', 'completed'] }
                }
              }
            }
          },
          required: ['todos']
        },
        execute: async ({ todos }) => {
          this.todoList = todos;
          console.log('📋 Todo list updated:', todos);
          return JSON.stringify({ success: true, todos: this.todoList });
        }
      },
      {
        name: 'write_file',
        description: 'Write content to a file for persistence across operations',
        schema: {
          type: 'object',
          properties: {
            filename: { type: 'string' },
            content: { type: 'string' }
          },
          required: ['filename', 'content']
        },
        execute: async ({ filename, content }) => {
          this.fileSystem.set(filename, content);
          console.log(`💾 File written: ${filename}`);
          return JSON.stringify({ success: true, message: `File ${filename} written` });
        }
      },
      {
        name: 'read_file',
        description: 'Read content from a previously written file',
        schema: {
          type: 'object',
          properties: {
            filename: { type: 'string' }
          },
          required: ['filename']
        },
        execute: async ({ filename }) => {
          const content = this.fileSystem.get(filename);
          if (!content) {
            return JSON.stringify({ success: false, error: `File ${filename} not found` });
          }
          return JSON.stringify({ success: true, content });
        }
      },
      {
        name: 'ls',
        description: 'List all files in the workspace',
        schema: { type: 'object', properties: {} },
        execute: async () => {
          const files = Array.from(this.fileSystem.keys());
          return JSON.stringify({ success: true, files });
        }
      }
    ];
    
    return [...mcpTools, ...builtInTools];
  }

  /**
   * Create sub agents following deepagents pattern
   */
  createSubAgents() {
    return [
      {
        name: 'excel-analyst',
        description: 'Expert at analyzing Excel financial models and extracting insights',
        prompt: `You are an expert Excel analyst specializing in investment banking models. 
                 You excel at finding financial metrics, analyzing formulas, and identifying patterns in complex spreadsheets.
                 Use Excel tools to thoroughly analyze the workbook and provide detailed insights.`
      },
      {
        name: 'formula-expert', 
        description: 'Expert at Excel formulas, validating calculations, and creating new formulas',
        prompt: `You are an Excel formula expert. You can validate formula syntax, create complex formulas,
                 and troubleshoot calculation errors. Always validate formulas before applying them.`
      },
      {
        name: 'banking-advisor',
        description: 'Investment banking expert for financial analysis and deal structuring',
        prompt: `You are a senior investment banking analyst with deep expertise in M&A, LBOs, and valuations.
                 You provide strategic insights on deal structures, financial metrics, and market standards.`
      }
    ];
  }

  /**
   * Main processing method - implements Deep Agent pattern
   */
  async invoke({ messages }) {
    this.messages = [...messages];
    const userMessage = messages[messages.length - 1].content;
    
    // Deep Agent system prompt (based on deepagents package)
    const systemPrompt = `You are a Deep Agent with comprehensive Excel analysis capabilities.

## Core Architecture
You have access to a planning tool, sub agents, file system, and comprehensive Excel tools.

## Available Tools
${this.tools.map(t => `- ${t.name}: ${t.description}`).join('\n')}

## Sub Agents
${this.subAgents.map(s => `- ${s.name}: ${s.description}`).join('\n')}

## Process
1. **ALWAYS start with todo_write** to create a plan for complex tasks
2. **Use Excel tools** to gather data and analyze the workbook
3. **Store results** in files for later reference
4. **Create sub-agents** if needed for specialized analysis
5. **Synthesize findings** into comprehensive response

## Investment Banking Context
${this.instructions || 'You specialize in investment banking financial analysis, focusing on accuracy and professional standards.'}

Focus on thorough analysis, data-driven insights, and actionable recommendations.`;

    // Add system prompt to messages
    const messagesWithSystem = [
      { role: 'system', content: systemPrompt },
      ...this.messages
    ];

    try {
      // Process with OpenAI using function calling
      const response = await this.callOpenAI(messagesWithSystem);
      
      return {
        messages: [...messagesWithSystem, response],
        todoList: this.todoList,
        files: Array.from(this.fileSystem.entries()),
        success: true
      };
      
    } catch (error) {
      console.error('❌ Deep Agent error:', error);
      return {
        error: error.message,
        success: false
      };
    }
  }

  /**
   * Call OpenAI with function calling capabilities
   */
  async callOpenAI(messages) {
    const functions = this.tools.map(tool => ({
      name: tool.name,
      description: tool.description,
      parameters: tool.schema
    }));

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${this.apiKey}`
      },
      body: JSON.stringify({
        model: 'gpt-4-0125-preview',
        messages: messages,
        functions: functions,
        function_call: 'auto',
        temperature: 0.1
      })
    });

    if (!response.ok) {
      throw new Error(`OpenAI API error: ${response.statusText}`);
    }

    const data = await response.json();
    const message = data.choices[0].message;

    // Handle function calls
    if (message.function_call) {
      const functionName = message.function_call.name;
      const functionArgs = JSON.parse(message.function_call.arguments);
      
      console.log(`🔧 Deep Agent executing: ${functionName}`, functionArgs);
      
      const tool = this.tools.find(t => t.name === functionName);
      if (tool) {
        const result = await tool.execute(functionArgs);
        
        // Continue conversation with function result
        const newMessages = [
          ...messages,
          message,
          { role: 'function', name: functionName, content: result }
        ];
        
        return await this.callOpenAI(newMessages);
      }
    }

    return message;
  }
}

/**
 * Create Deep Agent function (mimics deepagents.create_deep_agent)
 */
function createDeepAgent(tools, instructions, options = {}) {
  const apiKey = options.apiKey || localStorage.getItem('openai_api_key');
  if (!apiKey) {
    throw new Error('OpenAI API key required');
  }
  
  return new DeepExcelAgent(apiKey, instructions);
}

// Initialize globally
if (typeof window !== 'undefined') {
  window.LangChainExcelTools = LangChainExcelTools;
  window.DeepExcelAgent = DeepExcelAgent;
  window.createDeepAgent = createDeepAgent;
  
  window.langChainExcelTools = new LangChainExcelTools();
  
  console.log('✅ Deep Agent with MCP server tools initialized');
  console.log('Available MCP tools:', window.langChainExcelTools.getAllTools().map(t => t.name));
}
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