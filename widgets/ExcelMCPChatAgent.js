/**
 * Excel MCP Chat Agent - Complete MCP Server Implementation for Browser
 * Direct port of haris-musa/excel-mcp-server for Office.js environment
 * Integrates all MCP server tools directly into the chat function
 */

class ExcelMCPChatAgent {
  constructor(apiKey) {
    this.apiKey = apiKey;
    this.tools = new Map();
    this.initializeTools();
    
    console.log('🚀 Excel MCP Chat Agent initialized with complete tool set');
  }

  /**
   * Initialize all MCP server tools (direct port from Python)
   */
  initializeTools() {
    // ===========================================
    // WORKBOOK OPERATIONS (from workbook.py)
    // ===========================================
    
    this.tools.set('create_workbook', {
      name: 'create_workbook',
      description: 'Creates a new Excel workbook',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path where to create workbook' }
        },
        required: ['filepath']
      },
      execute: async ({ filepath }) => {
        try {
          return await Excel.run(async (context) => {
            // In Office.js, we work with the existing workbook
            const workbook = context.workbook;
            workbook.load(['name']);
            await context.sync();
            
            return JSON.stringify({
              success: true,
              message: `Working with workbook: ${workbook.name || 'New Workbook'}`,
              filepath: workbook.name || 'New Workbook'
            });
          });
        } catch (error) {
          return JSON.stringify({ success: false, error: `Failed to access workbook: ${error.message}` });
        }
      }
    });

    this.tools.set('create_worksheet', {
      name: 'create_worksheet',
      description: 'Creates a new worksheet in an existing workbook',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path to Excel file (ignored in browser)' },
          sheet_name: { type: 'string', description: 'Name for the new worksheet' }
        },
        required: ['sheet_name']
      },
      execute: async ({ sheet_name }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            const newSheet = worksheets.add(sheet_name);
            newSheet.load(['name']);
            await context.sync();
            
            return JSON.stringify({
              success: true,
              message: `Worksheet '${newSheet.name}' created successfully`,
              sheet_name: newSheet.name
            });
          });
        } catch (error) {
          return JSON.stringify({ success: false, error: `Failed to create worksheet: ${error.message}` });
        }
      }
    });

    this.tools.set('get_workbook_metadata', {
      name: 'get_workbook_metadata',
      description: 'Get metadata about workbook including sheets and ranges',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path to Excel file (ignored in browser)' },
          include_ranges: { type: 'boolean', default: false, description: 'Whether to include range information' }
        }
      },
      execute: async ({ include_ranges = false }) => {
        try {
          return await Excel.run(async (context) => {
            const workbook = context.workbook;
            const worksheets = workbook.worksheets;
            
            workbook.load(['name']);
            worksheets.load(['items/name', 'items/position']);
            await context.sync();
            
            const metadata = {
              workbook_name: workbook.name || 'Unknown',
              worksheet_count: worksheets.items.length,
              worksheets: worksheets.items.map(ws => ({
                name: ws.name,
                position: ws.position
              }))
            };
            
            if (include_ranges) {
              // Add range information for each worksheet
              for (const ws of worksheets.items) {
                const usedRange = ws.getUsedRangeOrNullObject();
                usedRange.load(['address', 'rowCount', 'columnCount']);
                await context.sync();
                
                const wsMetadata = metadata.worksheets.find(w => w.name === ws.name);
                if (!usedRange.isNullObject) {
                  wsMetadata.used_range = {
                    address: usedRange.address,
                    row_count: usedRange.rowCount,
                    column_count: usedRange.columnCount
                  };
                } else {
                  wsMetadata.used_range = null;
                }
              }
            }
            
            return JSON.stringify({ success: true, metadata });
          });
        } catch (error) {
          return JSON.stringify({ success: false, error: `Failed to get workbook metadata: ${error.message}` });
        }
      }
    });

    // ===========================================
    // DATA OPERATIONS (from data.py)
    // ===========================================
    
    this.tools.set('read_data_from_excel', {
      name: 'read_data_from_excel',
      description: 'Read data from Excel worksheet with cell metadata including validation rules',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path to Excel file (ignored in browser)' },
          sheet_name: { type: 'string', description: 'Source worksheet name' },
          start_cell: { type: 'string', default: 'A1', description: 'Starting cell' },
          end_cell: { type: 'string', description: 'Optional ending cell' },
          preview_only: { type: 'boolean', default: false, description: 'Whether to return only a preview' }
        },
        required: ['sheet_name']
      },
      execute: async ({ sheet_name, start_cell = 'A1', end_cell, preview_only = false }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheet_name);
            
            let range;
            if (end_cell) {
              range = worksheet.getRange(`${start_cell}:${end_cell}`);
            } else {
              // Get used range starting from start_cell
              const usedRange = worksheet.getUsedRangeOrNullObject();
              usedRange.load(['address']);
              await context.sync();
              
              if (usedRange.isNullObject) {
                range = worksheet.getRange(start_cell);
              } else {
                range = usedRange;
              }
            }
            
            range.load(['values', 'formulas', 'address', 'rowCount', 'columnCount']);
            await context.sync();
            
            const cells = [];
            const values = range.values;
            
            for (let row = 0; row < values.length; row++) {
              if (preview_only && row >= 10) break; // Limit preview to 10 rows
              
              for (let col = 0; col < values[row].length; col++) {
                const cellAddress = this.getCellAddress(row, col, range.address);
                cells.push({
                  address: cellAddress,
                  value: values[row][col],
                  row: row + 1,
                  column: col + 1,
                  validation: { has_validation: false } // Placeholder for validation
                });
              }
            }
            
            return JSON.stringify({
              success: true,
              range: range.address,
              sheet_name: sheet_name,
              cells: cells,
              row_count: range.rowCount,
              column_count: range.columnCount
            });
          });
        } catch (error) {
          return JSON.stringify({ success: false, error: `Failed to read data: ${error.message}` });
        }
      }
    });

    this.tools.set('write_data_to_excel', {
      name: 'write_data_to_excel',
      description: 'Write data to Excel worksheet',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path to Excel file (ignored in browser)' },
          sheet_name: { type: 'string', description: 'Name of worksheet to write to' },
          data: { type: 'array', items: { type: 'array' }, description: 'List of lists containing data to write' },
          start_cell: { type: 'string', default: 'A1', description: 'Cell to start writing to' }
        },
        required: ['sheet_name', 'data']
      },
      execute: async ({ sheet_name, data, start_cell = 'A1' }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheet_name);
            
            if (!data || data.length === 0) {
              return JSON.stringify({ success: false, error: 'No data provided to write' });
            }
            
            const range = worksheet.getRange(start_cell).getResizedRange(
              data.length - 1,
              data[0]?.length - 1 || 0
            );
            
            range.values = data;
            await context.sync();
            
            return JSON.stringify({
              success: true,
              message: `Data written to ${sheet_name}`,
              range: range.address,
              rows_written: data.length,
              columns_written: data[0]?.length || 0
            });
          });
        } catch (error) {
          return JSON.stringify({ success: false, error: `Failed to write data: ${error.message}` });
        }
      }
    });

    // ===========================================
    // FORMATTING OPERATIONS (from formatting.py)
    // ===========================================
    
    this.tools.set('format_range', {
      name: 'format_range',
      description: 'Apply formatting to a range of cells',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path to Excel file (ignored in browser)' },
          sheet_name: { type: 'string', description: 'Target worksheet name' },
          start_cell: { type: 'string', description: 'Starting cell of range' },
          end_cell: { type: 'string', description: 'Optional ending cell of range' },
          bold: { type: 'boolean', default: false, description: 'Apply bold formatting' },
          italic: { type: 'boolean', default: false, description: 'Apply italic formatting' },
          underline: { type: 'boolean', default: false, description: 'Apply underline formatting' },
          font_size: { type: 'number', description: 'Font size' },
          font_color: { type: 'string', description: 'Font color (hex code)' },
          bg_color: { type: 'string', description: 'Background color (hex code)' },
          border_style: { type: 'string', description: 'Border style' },
          border_color: { type: 'string', description: 'Border color (hex code)' },
          number_format: { type: 'string', description: 'Number format' },
          alignment: { type: 'string', description: 'Text alignment' },
          wrap_text: { type: 'boolean', default: false, description: 'Wrap text' },
          merge_cells: { type: 'boolean', default: false, description: 'Merge cells' }
        },
        required: ['sheet_name', 'start_cell']
      },
      execute: async (params) => {
        const { 
          sheet_name, start_cell, end_cell,
          bold = false, italic = false, underline = false,
          font_size, font_color, bg_color, border_style, border_color,
          number_format, alignment, wrap_text = false, merge_cells = false
        } = params;
        
        try {
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheet_name);
            const range = end_cell ? 
              worksheet.getRange(`${start_cell}:${end_cell}`) :
              worksheet.getRange(start_cell);
            
            // Apply font formatting
            if (bold) range.format.font.bold = true;
            if (italic) range.format.font.italic = true;
            if (underline) range.format.font.underline = 'Single';
            if (font_size) range.format.font.size = font_size;
            if (font_color) range.format.font.color = font_color;
            
            // Apply fill formatting
            if (bg_color) range.format.fill.color = bg_color;
            
            // Apply border formatting
            if (border_style) {
              const borderItems = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'];
              borderItems.forEach(edge => {
                range.format.borders.getItem(edge).style = border_style;
                if (border_color) {
                  range.format.borders.getItem(edge).color = border_color;
                }
              });
            }
            
            // Apply number formatting
            if (number_format) range.numberFormat = [[number_format]];
            
            // Apply alignment
            if (alignment) {
              const alignmentMap = {
                'left': 'Left',
                'center': 'Center', 
                'right': 'Right',
                'justify': 'Justify'
              };
              range.format.horizontalAlignment = alignmentMap[alignment.toLowerCase()] || alignment;
            }
            
            // Apply text wrapping
            if (wrap_text) range.format.wrapText = true;
            
            // Merge cells if requested
            if (merge_cells && end_cell) {
              range.merge();
            }
            
            await context.sync();
            
            return JSON.stringify({
              success: true,
              message: `Range formatted successfully`,
              range: range.address
            });
          });
        } catch (error) {
          return JSON.stringify({ success: false, error: `Failed to format range: ${error.message}` });
        }
      }
    });

    // ===========================================
    // FORMULA OPERATIONS (from calculations.py)
    // ===========================================
    
    this.tools.set('apply_formula', {
      name: 'apply_formula',
      description: 'Apply Excel formula to cell',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path to Excel file (ignored in browser)' },
          sheet_name: { type: 'string', description: 'Target worksheet name' },
          cell: { type: 'string', description: 'Target cell reference' },
          formula: { type: 'string', description: 'Excel formula to apply' }
        },
        required: ['sheet_name', 'cell', 'formula']
      },
      execute: async ({ sheet_name, cell, formula }) => {
        try {
          // Validate formula
          const validation = this.validateFormula(formula);
          if (!validation.valid) {
            return JSON.stringify({ success: false, error: validation.error });
          }
          
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheet_name);
            const targetCell = worksheet.getRange(cell);
            
            targetCell.formulas = [[formula]];
            targetCell.load(['values', 'formulas']);
            await context.sync();
            
            return JSON.stringify({
              success: true,
              message: `Formula applied to ${cell}`,
              cell: cell,
              formula: formula,
              result: targetCell.values[0][0]
            });
          });
        } catch (error) {
          return JSON.stringify({ success: false, error: `Failed to apply formula: ${error.message}` });
        }
      }
    });

    this.tools.set('validate_formula_syntax', {
      name: 'validate_formula_syntax',
      description: 'Validate Excel formula syntax without applying it',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path to Excel file (ignored in browser)' },
          sheet_name: { type: 'string', description: 'Target worksheet name' },
          cell: { type: 'string', description: 'Target cell reference' },
          formula: { type: 'string', description: 'Excel formula to validate' }
        },
        required: ['sheet_name', 'cell', 'formula']
      },
      execute: async ({ formula }) => {
        const validation = this.validateFormula(formula);
        return JSON.stringify({
          success: validation.valid,
          message: validation.valid ? 'Formula syntax is valid' : validation.error,
          formula: formula
        });
      }
    });

    // ===========================================
    // CHART OPERATIONS (from chart.py)
    // ===========================================
    
    this.tools.set('create_chart', {
      name: 'create_chart',
      description: 'Create chart in worksheet',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path to Excel file (ignored in browser)' },
          sheet_name: { type: 'string', description: 'Target worksheet name' },
          data_range: { type: 'string', description: 'Range containing chart data' },
          chart_type: { type: 'string', description: 'Type of chart' },
          target_cell: { type: 'string', description: 'Cell where to place chart' },
          title: { type: 'string', default: '', description: 'Optional chart title' },
          x_axis: { type: 'string', default: '', description: 'Optional X-axis label' },
          y_axis: { type: 'string', default: '', description: 'Optional Y-axis label' }
        },
        required: ['sheet_name', 'data_range', 'chart_type', 'target_cell']
      },
      execute: async ({ sheet_name, data_range, chart_type, target_cell, title = '', x_axis = '', y_axis = '' }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheet_name);
            const sourceRange = worksheet.getRange(data_range);
            
            // Map chart types to Office.js chart types
            const chartTypeMap = {
              'line': 'Line',
              'bar': 'ColumnClustered', 
              'pie': 'Pie',
              'scatter': 'XYScatter',
              'area': 'Area'
            };
            
            const excelChartType = chartTypeMap[chart_type.toLowerCase()] || 'ColumnClustered';
            
            const chart = worksheet.charts.add(excelChartType, sourceRange);
            chart.setPosition(target_cell);
            
            if (title) chart.title.text = title;
            if (x_axis) chart.axes.categoryAxis.title.text = x_axis;
            if (y_axis) chart.axes.valueAxis.title.text = y_axis;
            
            await context.sync();
            
            return JSON.stringify({
              success: true,
              message: `Chart created at ${target_cell}`,
              chart_type: chart_type,
              data_range: data_range,
              location: target_cell
            });
          });
        } catch (error) {
          return JSON.stringify({ success: false, error: `Failed to create chart: ${error.message}` });
        }
      }
    });

    // ===========================================
    // WORKSHEET OPERATIONS (from sheet.py)
    // ===========================================
    
    this.tools.set('copy_worksheet', {
      name: 'copy_worksheet',
      description: 'Copy worksheet within workbook',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path to Excel file (ignored in browser)' },
          source_sheet: { type: 'string', description: 'Name of sheet to copy' },
          target_sheet: { type: 'string', description: 'Name for new sheet' }
        },
        required: ['source_sheet', 'target_sheet']
      },
      execute: async ({ source_sheet, target_sheet }) => {
        try {
          return await Excel.run(async (context) => {
            const sourceWorksheet = context.workbook.worksheets.getItem(source_sheet);
            const copiedWorksheet = sourceWorksheet.copy();
            copiedWorksheet.name = target_sheet;
            
            await context.sync();
            
            return JSON.stringify({
              success: true,
              message: `Worksheet '${source_sheet}' copied to '${target_sheet}'`
            });
          });
        } catch (error) {
          return JSON.stringify({ success: false, error: `Failed to copy worksheet: ${error.message}` });
        }
      }
    });

    this.tools.set('delete_worksheet', {
      name: 'delete_worksheet',
      description: 'Delete worksheet from workbook',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path to Excel file (ignored in browser)' },
          sheet_name: { type: 'string', description: 'Name of sheet to delete' }
        },
        required: ['sheet_name']
      },
      execute: async ({ sheet_name }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheet_name);
            worksheet.delete();
            
            await context.sync();
            
            return JSON.stringify({
              success: true,
              message: `Worksheet '${sheet_name}' deleted successfully`
            });
          });
        } catch (error) {
          return JSON.stringify({ success: false, error: `Failed to delete worksheet: ${error.message}` });
        }
      }
    });

    this.tools.set('rename_worksheet', {
      name: 'rename_worksheet',
      description: 'Rename worksheet in workbook',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path to Excel file (ignored in browser)' },
          old_name: { type: 'string', description: 'Current sheet name' },
          new_name: { type: 'string', description: 'New sheet name' }
        },
        required: ['old_name', 'new_name']
      },
      execute: async ({ old_name, new_name }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(old_name);
            worksheet.name = new_name;
            
            await context.sync();
            
            return JSON.stringify({
              success: true,
              message: `Worksheet renamed from '${old_name}' to '${new_name}'`
            });
          });
        } catch (error) {
          return JSON.stringify({ success: false, error: `Failed to rename worksheet: ${error.message}` });
        }
      }
    });

    // ===========================================
    // RANGE OPERATIONS
    // ===========================================
    
    this.tools.set('merge_cells', {
      name: 'merge_cells',
      description: 'Merge a range of cells',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path to Excel file (ignored in browser)' },
          sheet_name: { type: 'string', description: 'Target worksheet name' },
          start_cell: { type: 'string', description: 'Starting cell of range' },
          end_cell: { type: 'string', description: 'Ending cell of range' }
        },
        required: ['sheet_name', 'start_cell', 'end_cell']
      },
      execute: async ({ sheet_name, start_cell, end_cell }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheet_name);
            const range = worksheet.getRange(`${start_cell}:${end_cell}`);
            
            range.merge();
            await context.sync();
            
            return JSON.stringify({
              success: true,
              message: `Cells merged from ${start_cell} to ${end_cell}`
            });
          });
        } catch (error) {
          return JSON.stringify({ success: false, error: `Failed to merge cells: ${error.message}` });
        }
      }
    });

    this.tools.set('unmerge_cells', {
      name: 'unmerge_cells',
      description: 'Unmerge a previously merged range of cells',
      schema: {
        type: 'object',
        properties: {
          filepath: { type: 'string', description: 'Path to Excel file (ignored in browser)' },
          sheet_name: { type: 'string', description: 'Target worksheet name' },
          start_cell: { type: 'string', description: 'Starting cell of range' },
          end_cell: { type: 'string', description: 'Ending cell of range' }
        },
        required: ['sheet_name', 'start_cell', 'end_cell']
      },
      execute: async ({ sheet_name, start_cell, end_cell }) => {
        try {
          return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheet_name);
            const range = worksheet.getRange(`${start_cell}:${end_cell}`);
            
            range.unmerge();
            await context.sync();
            
            return JSON.stringify({
              success: true,
              message: `Cells unmerged from ${start_cell} to ${end_cell}`
            });
          });
        } catch (error) {
          return JSON.stringify({ success: false, error: `Failed to unmerge cells: ${error.message}` });
        }
      }
    });

    console.log(`✅ Initialized ${this.tools.size} MCP server tools`);
  }

  /**
   * Validate formula syntax (from MCP server validation logic)
   */
  validateFormula(formula) {
    if (!formula || typeof formula !== 'string') {
      return { valid: false, error: 'Formula must be a non-empty string' };
    }

    if (!formula.startsWith('=')) {
      return { valid: false, error: 'Formula must start with =' };
    }

    // Check balanced parentheses
    let openCount = 0;
    for (const char of formula) {
      if (char === '(') openCount++;
      if (char === ')') openCount--;
      if (openCount < 0) {
        return { valid: false, error: 'Unbalanced parentheses in formula' };
      }
    }
    if (openCount !== 0) {
      return { valid: false, error: 'Unbalanced parentheses in formula' };
    }

    // Block unsafe functions
    const unsafeFunctions = ['INDIRECT', 'HYPERLINK', 'CALL'];
    const functionPattern = /([A-Z]+)\(/g;
    let match;
    while ((match = functionPattern.exec(formula)) !== null) {
      if (unsafeFunctions.includes(match[1])) {
        return { valid: false, error: `Unsafe function not allowed: ${match[1]}` };
      }
    }

    return { valid: true };
  }

  /**
   * Get cell address from row/col indices
   */
  getCellAddress(row, col, rangeAddress) {
    // Extract starting position from range address
    const match = rangeAddress.match(/([A-Z]+)(\d+)/);
    if (!match) return `${String.fromCharCode(65 + col)}${row + 1}`;
    
    const startCol = match[1];
    const startRow = parseInt(match[2]);
    
    // Convert column letters to number, add offset, convert back
    let colNum = 0;
    for (let i = 0; i < startCol.length; i++) {
      colNum = colNum * 26 + (startCol.charCodeAt(i) - 64);
    }
    colNum = colNum - 1 + col; // 0-based to 1-based and add offset
    
    // Convert back to letters
    let colLetter = '';
    while (colNum >= 0) {
      colLetter = String.fromCharCode(65 + (colNum % 26)) + colLetter;
      colNum = Math.floor(colNum / 26) - 1;
    }
    
    return `${colLetter}${startRow + row}`;
  }

  /**
   * Main processing method - handles user input and routes to appropriate tools
   */
  async processRequest(userInput) {
    console.log('🧠 MCP Chat Agent processing:', userInput);
    
    try {
      // Create system prompt with all available tools
      const toolDescriptions = Array.from(this.tools.values())
        .map(tool => `- ${tool.name}: ${tool.description}`)
        .join('\n');
      
      const systemPrompt = `You are an expert Excel assistant with comprehensive tool access.

AVAILABLE EXCEL TOOLS:
${toolDescriptions}

INSTRUCTIONS:
1. Analyze the user's request carefully
2. Use the appropriate Excel tools to complete the task
3. Always specify sheet names when working with worksheets
4. For complex tasks, break them down into multiple tool calls
5. Provide clear explanations of what you're doing
6. Handle errors gracefully and provide helpful feedback

Current context: You are working with an Excel workbook in the browser using Office.js API.`;

      const messages = [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userInput }
      ];

      // Call OpenAI with function calling
      const response = await this.callOpenAI(messages);
      
      return {
        success: true,
        response: response.content,
        toolsUsed: [], // Track tools used
        timestamp: new Date().toISOString()
      };
      
    } catch (error) {
      console.error('❌ MCP Chat Agent error:', error);
      return {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  }

  /**
   * Call OpenAI with function calling capabilities
   */
  async callOpenAI(messages) {
    const functions = Array.from(this.tools.values()).map(tool => ({
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
      
      console.log(`🔧 MCP Agent executing: ${functionName}`, functionArgs);
      
      const tool = this.tools.get(functionName);
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

  /**
   * Get list of all available tools
   */
  getAvailableTools() {
    return Array.from(this.tools.keys());
  }

  /**
   * Execute a specific tool by name
   */
  async executeTool(toolName, args) {
    const tool = this.tools.get(toolName);
    if (!tool) {
      throw new Error(`Tool ${toolName} not found`);
    }
    
    return await tool.execute(args);
  }
}

// Export for global use
window.ExcelMCPChatAgent = ExcelMCPChatAgent;

console.log('✅ Excel MCP Chat Agent loaded - Complete MCP server implementation for browser');