/**
 * ExcelMCPServer-SDK.ts
 * Rewritten using the official MCP TypeScript SDK - much simpler and more powerful
 */

import { McpServer, ResourceTemplate } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

// Create MCP server using the SDK
const excelServer = new McpServer({
  name: "excel-operations-server",
  version: "1.0.0"
});

// ===== RESOURCES (Excel Data Exposure) =====

// Workbook schema resource - like SQLite schema
excelServer.registerResource(
  "workbook-schema",
  "excel://workbook/schema",
  {
    title: "Workbook Schema",
    description: "Structure of the Excel workbook including worksheets, ranges, and named ranges",
    mimeType: "application/json"
  },
  async (uri) => {
    const schema = await getWorkbookSchema();
    return {
      contents: [{
        uri: uri.href,
        text: JSON.stringify(schema, null, 2)
      }]
    };
  }
);

// Dynamic worksheet data - like table contents
excelServer.registerResource(
  "worksheet-data",
  new ResourceTemplate("excel://worksheet/{worksheetName}", { 
    list: async () => {
      const worksheets = await getWorksheetNames();
      return worksheets.map(name => ({
        uri: `excel://worksheet/${name}`,
        name: `Worksheet: ${name}`
      }));
    }
  }),
  {
    title: "Worksheet Data",
    description: "Data from a specific worksheet"
  },
  async (uri, { worksheetName }) => {
    const data = await getWorksheetData(worksheetName);
    return {
      contents: [{
        uri: uri.href,
        text: JSON.stringify(data, null, 2),
        mimeType: "application/json"
      }]
    };
  }
);

// Financial metrics resource - like stored procedure results
excelServer.registerResource(
  "financial-metrics",
  "excel://metrics/financial",
  {
    title: "Financial Metrics",
    description: "Calculated financial metrics (MOIC, IRR, NPV) from the workbook",
    mimeType: "application/json"
  },
  async (uri) => {
    const metrics = await calculateFinancialMetrics();
    return {
      contents: [{
        uri: uri.href,
        text: JSON.stringify(metrics, null, 2)
      }]
    };
  }
);

// Named ranges resource
excelServer.registerResource(
  "named-range",
  new ResourceTemplate("excel://range/{rangeName}", {
    list: async () => {
      const ranges = await getNamedRanges();
      return ranges.map(range => ({
        uri: `excel://range/${range.name}`,
        name: `Range: ${range.name}`
      }));
    }
  }),
  {
    title: "Named Range",
    description: "Data from a named range"
  },
  async (uri, { rangeName }) => {
    const data = await getNamedRangeData(rangeName);
    return {
      contents: [{
        uri: uri.href,
        text: JSON.stringify(data, null, 2)
      }]
    };
  }
);

// ===== TOOLS (Excel Operations) =====

// Excel query tool - like SQL query
excelServer.registerTool(
  "excel-query",
  {
    title: "Excel Query",
    description: "Query Excel data using natural language or range specifications",
    inputSchema: {
      query: z.string().describe("Natural language query or range specification"),
      format: z.enum(["json", "csv", "table"]).optional().describe("Output format")
    }
  },
  async ({ query, format = "json" }) => {
    try {
      const results = await executeExcelQuery(query);
      const formatted = formatResults(results, format);
      
      return {
        content: [{
          type: "text",
          text: formatted
        }]
      };
    } catch (error: any) {
      return {
        content: [{
          type: "text", 
          text: `Query error: ${error.message}`
        }],
        isError: true
      };
    }
  }
);

// Write data tool
excelServer.registerTool(
  "write-data",
  {
    title: "Write Data",
    description: "Write values or formulas to Excel cells",
    inputSchema: {
      range: z.string().describe("Target range (e.g., 'A1:B10')"),
      data: z.any().describe("Data to write (value, formula, or array)"),
      worksheet: z.string().optional().describe("Worksheet name (defaults to active)")
    }
  },
  async ({ range, data, worksheet }) => {
    try {
      await writeToExcel(range, data, worksheet);
      return {
        content: [{
          type: "text",
          text: `Successfully wrote data to ${worksheet ? worksheet + '!' : ''}${range}`
        }]
      };
    } catch (error: any) {
      return {
        content: [{
          type: "text",
          text: `Write error: ${error.message}`
        }],
        isError: true
      };
    }
  }
);

// Format cells tool
excelServer.registerTool(
  "format-cells", 
  {
    title: "Format Cells",
    description: "Apply formatting to Excel cells (colors, fonts, borders, conditional formatting)",
    inputSchema: {
      range: z.string().describe("Target range"),
      format: z.object({
        backgroundColor: z.string().optional(),
        fontColor: z.string().optional(),
        bold: z.boolean().optional(),
        fontSize: z.number().optional(),
        numberFormat: z.string().optional()
      }).describe("Formatting options"),
      worksheet: z.string().optional()
    }
  },
  async ({ range, format, worksheet }) => {
    try {
      await formatExcelCells(range, format, worksheet);
      return {
        content: [{
          type: "text",
          text: `Successfully formatted ${worksheet ? worksheet + '!' : ''}${range}`
        }]
      };
    } catch (error: any) {
      return {
        content: [{
          type: "text",
          text: `Format error: ${error.message}`
        }],
        isError: true
      };
    }
  }
);

// Financial calculation tool with AI sampling
excelServer.registerTool(
  "analyze-financials",
  {
    title: "Analyze Financials",
    description: "Perform complex financial analysis using AI assistance",
    inputSchema: {
      analysis_type: z.enum(["irr", "npv", "moic", "sensitivity", "scenario"]),
      parameters: z.object({
        cash_flows: z.string().optional(),
        discount_rate: z.number().optional(),
        scenarios: z.array(z.string()).optional()
      }).optional()
    }
  },
  async ({ analysis_type, parameters }) => {
    try {
      // Get relevant financial data
      const financialData = await getFinancialData(analysis_type, parameters);
      
      // Use AI sampling for complex analysis
      const aiAnalysis = await excelServer.server.createMessage({
        messages: [{
          role: "user",
          content: {
            type: "text",
            text: `Analyze this ${analysis_type.toUpperCase()} data and provide insights:
            
            Data: ${JSON.stringify(financialData, null, 2)}
            
            Please provide:
            1. Calculated ${analysis_type.toUpperCase()} value
            2. Key insights and trends
            3. Recommendations for optimization
            4. Risk factors to consider`
          }
        }],
        maxTokens: 1000
      });
      
      return {
        content: [{
          type: "text",
          text: aiAnalysis.content.type === "text" ? aiAnalysis.content.text : "Analysis completed"
        }]
      };
    } catch (error: any) {
      return {
        content: [{
          type: "text",
          text: `Analysis error: ${error.message}`
        }],
        isError: true
      };
    }
  }
);

// Smart Excel assistant tool with elicitation
excelServer.registerTool(
  "excel-assistant",
  {
    title: "Excel Assistant",
    description: "Interactive Excel assistant that can ask for clarification",
    inputSchema: {
      request: z.string().describe("What you want to do with Excel")
    }
  },
  async ({ request }) => {
    try {
      // Analyze the request complexity
      const complexity = analyzeRequestComplexity(request);
      
      if (complexity.needsMoreInfo) {
        // Use elicitation to get more details
        const userInput = await excelServer.server.elicitInput({
          message: `I need more information to help with: "${request}"`,
          requestedSchema: {
            type: "object",
            properties: {
              target_range: {
                type: "string",
                title: "Target Range",
                description: "Which Excel range should I work with?"
              },
              operation_type: {
                type: "string",
                title: "Operation",
                enum: ["read", "write", "format", "calculate", "analyze"],
                enumNames: ["Read Data", "Write Data", "Format Cells", "Calculate", "Analyze"]
              },
              specific_requirements: {
                type: "string",
                title: "Additional Requirements",
                description: "Any specific requirements or preferences?"
              }
            },
            required: ["target_range", "operation_type"]
          }
        });
        
        if (userInput.action === "accept") {
          // Process with the elicited information
          const result = await processExcelRequest(request, userInput.content);
          return {
            content: [{
              type: "text",
              text: `Completed your request: ${result}`
            }]
          };
        } else {
          return {
            content: [{
              type: "text",
              text: "Request cancelled by user"
            }]
          };
        }
      } else {
        // Process directly
        const result = await processExcelRequest(request, null);
        return {
          content: [{
            type: "text",
            text: result
          }]
        };
      }
    } catch (error: any) {
      return {
        content: [{
          type: "text",
          text: `Assistant error: ${error.message}`
        }],
        isError: true
      };
    }
  }
);

// ===== PROMPTS (Excel Templates) =====

excelServer.registerPrompt(
  "financial-model-review",
  {
    title: "Financial Model Review",
    description: "Review and analyze a financial model",
    argsSchema: {
      model_type: z.enum(["dcf", "lbo", "merger", "valuation"]),
      focus_areas: z.array(z.string()).optional()
    }
  },
  ({ model_type, focus_areas }) => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: `Please review this ${model_type.toUpperCase()} model and analyze:
        ${focus_areas ? `\nFocus areas: ${focus_areas.join(', ')}` : ''}
        
        Look for:
        - Formula errors and circular references
        - Assumption reasonableness
        - Model structure and best practices
        - Key risk factors
        - Sensitivity to key variables`
      }
    }]
  })
);

excelServer.registerPrompt(
  "excel-formula-help",
  {
    title: "Excel Formula Help",
    description: "Get help with Excel formulas",
    argsSchema: {
      formula_type: z.string(),
      use_case: z.string()
    }
  },
  ({ formula_type, use_case }) => ({
    messages: [{
      role: "user", 
      content: {
        type: "text",
        text: `I need help creating a ${formula_type} formula for: ${use_case}
        
        Please provide:
        1. The exact formula syntax
        2. Explanation of how it works
        3. Example with sample data
        4. Common pitfalls to avoid`
      }
    }]
  })
);

// ===== HELPER FUNCTIONS =====

async function getWorkbookSchema(): Promise<any> {
  // Implementation would use Office.js to get workbook structure
  return {
    worksheets: ["Sheet1", "Financial Model", "Assumptions"],
    namedRanges: ["CashFlows", "Assumptions", "Results"],
    metrics: ["MOIC", "IRR", "NPV"]
  };
}

async function getWorksheetNames(): Promise<string[]> {
  if (typeof Excel === 'undefined') return ["Sheet1"];
  
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load('items/name');
    await context.sync();
    return worksheets.items.map(ws => ws.name);
  });
}

async function getWorksheetData(worksheetName: string): Promise<any> {
  if (typeof Excel === 'undefined') return { error: "Excel not available" };
  
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(worksheetName);
    const usedRange = worksheet.getUsedRange();
    usedRange.load(['values', 'formulas', 'address']);
    await context.sync();
    
    return {
      worksheet: worksheetName,
      range: usedRange.address,
      values: usedRange.values,
      formulas: usedRange.formulas
    };
  });
}

async function calculateFinancialMetrics(): Promise<any> {
  // Implementation would calculate MOIC, IRR, NPV from Excel data
  return {
    MOIC: { value: 2.5, formula: "=SUM(Returns)/SUM(Investments)", confidence: "high" },
    IRR: { value: 0.25, formula: "=IRR(CashFlows)", confidence: "high" },
    NPV: { value: 1500000, formula: "=NPV(DiscountRate,CashFlows)", confidence: "medium" }
  };
}

async function executeExcelQuery(query: string): Promise<any> {
  // Natural language to Excel operations
  // This could use AI to interpret the query
  return { results: "Query results would go here" };
}

async function writeToExcel(range: string, data: any, worksheet?: string): Promise<void> {
  if (typeof Excel === 'undefined') throw new Error("Excel not available");
  
  await Excel.run(async (context) => {
    const ws = worksheet 
      ? context.workbook.worksheets.getItem(worksheet)
      : context.workbook.worksheets.getActiveWorksheet();
    
    const targetRange = ws.getRange(range);
    
    if (typeof data === 'string' && data.startsWith('=')) {
      targetRange.formulas = [[data]];
    } else {
      targetRange.values = Array.isArray(data) ? data : [[data]];
    }
    
    await context.sync();
  });
}

async function formatExcelCells(range: string, format: any, worksheet?: string): Promise<void> {
  if (typeof Excel === 'undefined') throw new Error("Excel not available");
  
  await Excel.run(async (context) => {
    const ws = worksheet 
      ? context.workbook.worksheets.getItem(worksheet)
      : context.workbook.worksheets.getActiveWorksheet();
    
    const targetRange = ws.getRange(range);
    
    if (format.backgroundColor) targetRange.format.fill.color = format.backgroundColor;
    if (format.fontColor) targetRange.format.font.color = format.fontColor;
    if (format.bold !== undefined) targetRange.format.font.bold = format.bold;
    if (format.fontSize) targetRange.format.font.size = format.fontSize;
    if (format.numberFormat) targetRange.numberFormat = [[format.numberFormat]];
    
    await context.sync();
  });
}

function analyzeRequestComplexity(request: string): { needsMoreInfo: boolean; complexity: string } {
  // Simple analysis - in production this could use AI
  const ambiguous = ['color', 'format', 'calculate', 'analyze'].some(word => 
    request.toLowerCase().includes(word) && !request.includes('range') && !request.includes('cell')
  );
  
  return {
    needsMoreInfo: ambiguous,
    complexity: ambiguous ? 'high' : 'low'
  };
}

async function processExcelRequest(request: string, context: any): Promise<string> {
  // Process the Excel request with context
  return `Processed request: "${request}" with context: ${JSON.stringify(context)}`;
}

async function getFinancialData(analysisType: string, parameters?: any): Promise<any> {
  // Get relevant financial data for analysis
  return { analysisType, parameters, data: "Sample financial data" };
}

async function getNamedRanges(): Promise<Array<{name: string}>> {
  if (typeof Excel === 'undefined') return [];
  
  return Excel.run(async (context) => {
    const namedRanges = context.workbook.names;
    namedRanges.load('items/name');
    await context.sync();
    return namedRanges.items.map(range => ({ name: range.name }));
  });
}

async function getNamedRangeData(rangeName: string): Promise<any> {
  if (typeof Excel === 'undefined') return { error: "Excel not available" };
  
  return Excel.run(async (context) => {
    const namedRange = context.workbook.names.getItem(rangeName);
    const range = namedRange.getRange();
    range.load(['values', 'formulas', 'address']);
    await context.sync();
    
    return {
      name: rangeName,
      address: range.address,
      values: range.values,
      formulas: range.formulas
    };
  });
}

function formatResults(results: any, format: string): string {
  switch (format) {
    case 'json':
      return JSON.stringify(results, null, 2);
    case 'csv':
      // Convert to CSV format
      return 'CSV format not implemented yet';
    case 'table':
      // Convert to table format
      return 'Table format not implemented yet';
    default:
      return JSON.stringify(results, null, 2);
  }
}

// ===== SERVER STARTUP =====

async function startExcelMCPServer() {
  console.log("🚀 Starting Excel MCP Server with SDK...");
  
  try {
    const transport = new StdioServerTransport();
    await excelServer.connect(transport);
    console.log("✅ Excel MCP Server connected and ready!");
  } catch (error) {
    console.error("❌ Failed to start Excel MCP Server:", error);
    process.exit(1);
  }
}

// Export server for use
export { excelServer, startExcelMCPServer };
export default excelServer;