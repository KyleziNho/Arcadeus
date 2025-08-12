/**
 * Comprehensive Excel API Toolkit
 * Advanced tools for robust Excel operations and analysis
 */

class ExcelApiTool {
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

/**
 * Analyze the overall structure of the Excel workbook
 */
class AnalyzeWorkbookStructureTool extends ExcelApiTool {
  constructor() {
    super();
    this.name = "analyze_workbook_structure";
    this.description = "Analyze the overall structure of the Excel workbook including sheets, used ranges, and data organization";
    this.schema = {
      type: "object",
      properties: {
        includeFormatting: { type: "boolean", default: true },
        maxRows: { type: "number", default: 100 }
      }
    };
  }

  async _call(input = {}) {
    console.log('📊 AnalyzeWorkbookStructureTool: Analyzing workbook structure');
    
    let result;
    try {
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();

        const workbookAnalysis = {
          sheets: [],
          totalSheets: worksheets.items.length,
          activeSheet: null
        };

        for (let i = 0; i < worksheets.items.length; i++) {
          const sheet = worksheets.items[i];
          const usedRange = sheet.getUsedRange();
          
          try {
            usedRange.load(["address", "rowCount", "columnCount", "values"]);
            await context.sync();

            const sheetInfo = {
              name: sheet.name,
              usedRange: usedRange.address,
              rowCount: usedRange.rowCount,
              columnCount: usedRange.columnCount,
              isEmpty: false
            };

            // Get sample data from first few rows if requested
            if (input.maxRows && usedRange.rowCount > 0) {
              const sampleRowCount = Math.min(input.maxRows, usedRange.rowCount);
              const sampleRange = sheet.getRange(`A1:${this.getColumnLetter(usedRange.columnCount)}${sampleRowCount}`);
              sampleRange.load("values");
              await context.sync();
              sheetInfo.sampleData = sampleRange.values;
            }

            workbookAnalysis.sheets.push(sheetInfo);
          } catch (error) {
            // Sheet is empty
            workbookAnalysis.sheets.push({
              name: sheet.name,
              usedRange: null,
              rowCount: 0,
              columnCount: 0,
              isEmpty: true
            });
          }
        }

        result = {
          success: true,
          analysis: workbookAnalysis,
          timestamp: new Date().toISOString()
        };
      });
    } catch (error) {
      result = {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }

    return JSON.stringify(result);
  }

  getColumnLetter(columnNumber) {
    let result = '';
    while (columnNumber > 0) {
      columnNumber--;
      result = String.fromCharCode(65 + (columnNumber % 26)) + result;
      columnNumber = Math.floor(columnNumber / 26);
    }
    return result;
  }
}

/**
 * Find cells by background or font color
 */
class FindCellsByColorTool extends ExcelApiTool {
  constructor() {
    super();
    this.name = "find_cells_by_color";
    this.description = "Find all cells that have a specific background or font color";
    this.schema = {
      type: "object",
      properties: {
        color: { type: "string", description: "Color name (red, blue, green) or hex code (#FF0000)" },
        colorType: { type: "string", enum: ["background", "font", "both"], default: "background" },
        sheetName: { type: "string", description: "Specific sheet to search, or all sheets if not provided" }
      },
      required: ["color"]
    };
  }

  async _call(input) {
    console.log(`🎨 FindCellsByColorTool: Finding ${input.colorType} color ${input.color}`);
    
    let result;
    try {
      await Excel.run(async (context) => {
        const sheets = input.sheetName ? 
          [context.workbook.worksheets.getItem(input.sheetName)] :
          context.workbook.worksheets.items;

        const colorMatches = [];
        const targetColor = this.normalizeColor(input.color);

        for (const sheet of sheets) {
          const usedRange = sheet.getUsedRange();
          
          try {
            usedRange.load(["address", "format/fill/color", "format/font/color", "values"]);
            await context.sync();

            // Check each cell in the used range
            for (let row = 0; row < usedRange.rowCount; row++) {
              for (let col = 0; col < usedRange.columnCount; col++) {
                const cell = usedRange.getCell(row, col);
                cell.load(["address", "format/fill/color", "format/font/color", "values"]);
                await context.sync();

                const backgroundColor = this.normalizeColor(cell.format.fill.color);
                const fontColor = this.normalizeColor(cell.format.font.color);
                
                let isMatch = false;
                if (input.colorType === "background" || input.colorType === "both") {
                  isMatch = isMatch || backgroundColor === targetColor;
                }
                if (input.colorType === "font" || input.colorType === "both") {
                  isMatch = isMatch || fontColor === targetColor;
                }

                if (isMatch) {
                  colorMatches.push({
                    address: cell.address,
                    sheetName: sheet.name,
                    value: cell.values[0][0],
                    backgroundColor: backgroundColor,
                    fontColor: fontColor
                  });
                }
              }
            }
          } catch (error) {
            console.log(`Sheet ${sheet.name} is empty, skipping`);
          }
        }

        result = {
          success: true,
          matches: colorMatches,
          count: colorMatches.length,
          searchCriteria: input,
          timestamp: new Date().toISOString()
        };
      });
    } catch (error) {
      result = {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }

    return JSON.stringify(result);
  }

  normalizeColor(color) {
    if (!color) return null;
    
    // Convert common color names to hex
    const colorMap = {
      'red': '#FF0000',
      'blue': '#0000FF', 
      'green': '#00FF00',
      'yellow': '#FFFF00',
      'orange': '#FFA500',
      'purple': '#800080',
      'black': '#000000',
      'white': '#FFFFFF'
    };

    const lowerColor = color.toLowerCase();
    if (colorMap[lowerColor]) {
      return colorMap[lowerColor];
    }

    // Return as-is if already hex format
    if (color.startsWith('#')) {
      return color.toUpperCase();
    }

    return color;
  }
}

/**
 * Analyze headers in the workbook
 */
class AnalyzeHeadersTool extends ExcelApiTool {
  constructor() {
    super();
    this.name = "analyze_headers";
    this.description = "Identify and analyze headers in the workbook based on position, formatting, and content patterns";
    this.schema = {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        cells: { type: "array", description: "Specific cells to analyze as potential headers" }
      }
    };
  }

  async _call(input = {}) {
    console.log('📋 AnalyzeHeadersTool: Analyzing headers');
    
    let result;
    try {
      await Excel.run(async (context) => {
        const sheet = input.sheetName ? 
          context.workbook.worksheets.getItem(input.sheetName) :
          context.workbook.getActiveWorksheet();

        const headers = [];

        if (input.cells && input.cells.length > 0) {
          // Analyze specific cells provided
          for (const cellAddress of input.cells) {
            const cell = sheet.getRange(cellAddress);
            cell.load(["address", "values", "format/fill/color", "format/font/bold", "format/font/size"]);
            await context.sync();

            const isHeader = this.isLikelyHeader(cell, null);
            if (isHeader.likely) {
              headers.push({
                address: cell.address,
                value: cell.values[0][0],
                confidence: isHeader.confidence,
                reasons: isHeader.reasons
              });
            }
          }
        } else {
          // Analyze first few rows to find headers
          const topRows = sheet.getRange("A1:Z10"); // Check first 10 rows, 26 columns
          topRows.load(["values", "format"]);
          await context.sync();

          for (let row = 0; row < Math.min(10, topRows.rowCount); row++) {
            for (let col = 0; col < Math.min(26, topRows.columnCount); col++) {
              const cell = topRows.getCell(row, col);
              const cellValue = topRows.values[row][col];
              
              if (cellValue && typeof cellValue === 'string') {
                cell.load(["address", "format/fill/color", "format/font/bold", "format/font/size"]);
                await context.sync();

                const isHeader = this.isLikelyHeader(cell, cellValue);
                if (isHeader.likely) {
                  headers.push({
                    address: cell.address,
                    value: cellValue,
                    confidence: isHeader.confidence,
                    reasons: isHeader.reasons
                  });
                }
              }
            }
          }
        }

        result = {
          success: true,
          headers: headers,
          count: headers.length,
          timestamp: new Date().toISOString()
        };
      });
    } catch (error) {
      result = {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }

    return JSON.stringify(result);
  }

  isLikelyHeader(cell, value) {
    const reasons = [];
    let confidence = 0;

    // Check if it's in first few rows (strong indicator)
    const address = cell.address;
    const rowNum = parseInt(address.match(/\d+/)[0]);
    if (rowNum <= 3) {
      confidence += 0.3;
      reasons.push("Located in top 3 rows");
    }

    // Check formatting
    try {
      if (cell.format.font.bold) {
        confidence += 0.2;
        reasons.push("Bold formatting");
      }
      
      if (cell.format.fill.color && cell.format.fill.color !== '#FFFFFF') {
        confidence += 0.2;
        reasons.push("Has background color");
      }

      if (cell.format.font.size > 11) {
        confidence += 0.1;
        reasons.push("Larger font size");
      }
    } catch (error) {
      // Formatting not available
    }

    // Check content patterns
    if (value) {
      if (typeof value === 'string') {
        const lowerValue = value.toLowerCase();
        
        // Common header words
        const headerKeywords = ['revenue', 'cost', 'profit', 'total', 'summary', 'year', 'month', 'date', 'name', 'id', 'amount'];
        if (headerKeywords.some(keyword => lowerValue.includes(keyword))) {
          confidence += 0.15;
          reasons.push("Contains header keywords");
        }

        // Short descriptive text
        if (value.length < 50 && value.length > 2) {
          confidence += 0.1;
          reasons.push("Appropriate length for header");
        }

        // All caps might be header
        if (value === value.toUpperCase() && value.length > 2) {
          confidence += 0.1;
          reasons.push("All uppercase text");
        }
      }
    }

    return {
      likely: confidence > 0.5,
      confidence: confidence,
      reasons: reasons
    };
  }
}

/**
 * Enhanced cell formatting tool
 */
class FormatCellsTool extends ExcelApiTool {
  constructor() {
    super();
    this.name = "format_cells";
    this.description = "Apply comprehensive formatting to specific cells including colors, fonts, borders, and alignment";
    this.schema = {
      type: "object",
      properties: {
        cells: { type: "array", description: "Array of cell addresses to format" },
        backgroundColor: { type: "string" },
        fontColor: { type: "string" },
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        fontSize: { type: "number" },
        borderStyle: { type: "string", enum: ["thin", "medium", "thick", "none"] },
        borderColor: { type: "string" },
        alignment: { type: "string", enum: ["left", "center", "right"] }
      },
      required: ["cells"]
    };
  }

  async _call(input) {
    console.log(`🎨 FormatCellsTool: Formatting ${input.cells.length} cells`);
    
    let result;
    try {
      await Excel.run(async (context) => {
        const formattedCells = [];

        for (const cellAddress of input.cells) {
          // Parse sheet name if included
          let sheet, range;
          if (cellAddress.includes('!')) {
            const [sheetName, addr] = cellAddress.split('!');
            sheet = context.workbook.worksheets.getItem(sheetName);
            range = sheet.getRange(addr);
          } else {
            sheet = context.workbook.getActiveWorksheet();
            range = sheet.getRange(cellAddress);
          }

          // Apply formatting
          if (input.backgroundColor) {
            range.format.fill.color = this.normalizeColor(input.backgroundColor);
          }
          
          if (input.fontColor) {
            range.format.font.color = this.normalizeColor(input.fontColor);
          }

          if (input.bold !== undefined) {
            range.format.font.bold = input.bold;
          }

          if (input.italic !== undefined) {
            range.format.font.italic = input.italic;
          }

          if (input.fontSize) {
            range.format.font.size = input.fontSize;
          }

          if (input.alignment) {
            const alignmentMap = {
              'left': Excel.HorizontalAlignment.left,
              'center': Excel.HorizontalAlignment.center,
              'right': Excel.HorizontalAlignment.right
            };
            range.format.horizontalAlignment = alignmentMap[input.alignment];
          }

          if (input.borderStyle && input.borderStyle !== 'none') {
            const borderStyleMap = {
              'thin': Excel.BorderLineStyle.continuous,
              'medium': Excel.BorderLineStyle.continuous,
              'thick': Excel.BorderLineStyle.continuous
            };
            
            range.format.borders.getItem(Excel.BorderIndex.edgeTop).style = borderStyleMap[input.borderStyle];
            range.format.borders.getItem(Excel.BorderIndex.edgeBottom).style = borderStyleMap[input.borderStyle];
            range.format.borders.getItem(Excel.BorderIndex.edgeLeft).style = borderStyleMap[input.borderStyle];
            range.format.borders.getItem(Excel.BorderIndex.edgeRight).style = borderStyleMap[input.borderStyle];
            
            if (input.borderColor) {
              range.format.borders.getItem(Excel.BorderIndex.edgeTop).color = this.normalizeColor(input.borderColor);
              range.format.borders.getItem(Excel.BorderIndex.edgeBottom).color = this.normalizeColor(input.borderColor);
              range.format.borders.getItem(Excel.BorderIndex.edgeLeft).color = this.normalizeColor(input.borderColor);
              range.format.borders.getItem(Excel.BorderIndex.edgeRight).color = this.normalizeColor(input.borderColor);
            }
          }

          formattedCells.push({
            address: cellAddress,
            appliedFormatting: input
          });
        }

        await context.sync();

        result = {
          success: true,
          formattedCells: formattedCells,
          count: formattedCells.length,
          timestamp: new Date().toISOString()
        };
      });
    } catch (error) {
      result = {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }

    return JSON.stringify(result);
  }

  normalizeColor(color) {
    if (!color) return null;
    
    const colorMap = {
      'red': '#FF0000',
      'blue': '#0000FF', 
      'green': '#00FF00',
      'yellow': '#FFFF00',
      'orange': '#FFA500',
      'purple': '#800080',
      'black': '#000000',
      'white': '#FFFFFF'
    };

    const lowerColor = color.toLowerCase();
    if (colorMap[lowerColor]) {
      return colorMap[lowerColor];
    }

    if (color.startsWith('#')) {
      return color.toUpperCase();
    }

    return color;
  }
}

/**
 * Read current formatting of cells
 */
class ReadFormattingTool extends ExcelApiTool {
  constructor() {
    super();
    this.name = "read_formatting";
    this.description = "Read the current formatting properties of specified cells";
    this.schema = {
      type: "object", 
      properties: {
        cells: { type: "array", description: "Array of cell addresses to read formatting from" }
      },
      required: ["cells"]
    };
  }

  async _call(input) {
    console.log(`🔍 ReadFormattingTool: Reading formatting for ${input.cells.length} cells`);
    
    let result;
    try {
      await Excel.run(async (context) => {
        const cellFormatting = [];

        for (const cellAddress of input.cells) {
          let sheet, range;
          if (cellAddress.includes('!')) {
            const [sheetName, addr] = cellAddress.split('!');
            sheet = context.workbook.worksheets.getItem(sheetName);
            range = sheet.getRange(addr);
          } else {
            sheet = context.workbook.getActiveWorksheet();
            range = sheet.getRange(cellAddress);
          }

          range.load([
            "values",
            "format/fill/color",
            "format/font/color", 
            "format/font/bold",
            "format/font/italic",
            "format/font/size",
            "format/horizontalAlignment"
          ]);
          
          await context.sync();

          cellFormatting.push({
            address: cellAddress,
            value: range.values[0][0],
            formatting: {
              backgroundColor: range.format.fill.color,
              fontColor: range.format.font.color,
              bold: range.format.font.bold,
              italic: range.format.font.italic,
              fontSize: range.format.font.size,
              horizontalAlignment: range.format.horizontalAlignment
            }
          });
        }

        result = {
          success: true,
          cellFormatting: cellFormatting,
          count: cellFormatting.length,
          timestamp: new Date().toISOString()
        };
      });
    } catch (error) {
      result = {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }

    return JSON.stringify(result);
  }
}

// Create instances of all tools
const excelApiTools = [
  new AnalyzeWorkbookStructureTool(),
  new FindCellsByColorTool(),
  new AnalyzeHeadersTool(),
  new FormatCellsTool(),
  new ReadFormattingTool()
];

// Initialize globally
if (typeof window !== 'undefined') {
  window.ExcelApiToolkit = {
    AnalyzeWorkbookStructureTool,
    FindCellsByColorTool,
    AnalyzeHeadersTool,
    FormatCellsTool,
    ReadFormattingTool,
    excelApiTools
  };
  
  console.log('✅ Excel API Toolkit initialized with', excelApiTools.length, 'advanced tools');
}