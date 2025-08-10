/**
 * MessageProcessor.ts
 * Processes user messages, analyzes intent, and routes to appropriate MCP operations
 */

import {
  MCPTool,
  MCPToolCall,
  ExcelState,
  ConversationMessage
} from '../mcp-types/interfaces';

export interface ProcessedIntent {
  type: string;
  confidence: number;
  entity?: string;
  action?: string;
  target?: string;
  parameters?: Record<string, any>;
  requiresConfirmation?: boolean;
  isMultiStep?: boolean;
}

export interface OperationPlan {
  operations: MCPToolCall[];
  description: string;
  estimatedTime?: number;
  requiresApproval?: boolean;
}

export class MessageProcessor {
  private mcpClient: any; // Reference to MCPChatClient
  private intentPatterns: Map<string, RegExp[]> = new Map();
  private entityExtractors: Map<string, RegExp> = new Map();

  constructor(mcpClient: any) {
    this.mcpClient = mcpClient;
    this.initializePatterns();
  }

  /**
   * Initialize intent recognition patterns
   */
  private initializePatterns(): void {
    // Excel read operations
    this.intentPatterns.set('excel_read', [
      /show me|display|view|see|get|read/i,
      /what is|what's|tell me about/i,
      /value of|values in|content of/i
    ]);

    // Excel write operations  
    this.intentPatterns.set('excel_write', [
      /set|put|write|add|enter|input|change/i,
      /update|modify|edit|alter/i,
      /fill|populate|insert/i
    ]);

    // Excel format operations
    this.intentPatterns.set('excel_format', [
      /format|style|color|colour/i,
      /bold|italic|font|highlight/i,
      /border|alignment|background/i
    ]);

    // Excel calculate operations
    this.intentPatterns.set('excel_calculate', [
      /calculate|compute|sum|total/i,
      /average|mean|count|max|min/i,
      /formula|function|equation/i
    ]);

    // Analysis operations
    this.intentPatterns.set('excel_analyze', [
      /analyze|analyse|review|examine/i,
      /insights|trends|patterns/i,
      /summarize|summary|report/i
    ]);

    // Chart/visualization operations
    this.intentPatterns.set('excel_chart', [
      /chart|graph|plot|visualize/i,
      /pie chart|bar chart|line graph/i,
      /dashboard|visualization/i
    ]);

    // Entity extractors
    this.entityExtractors.set('range', /([A-Z]+\d+(?::[A-Z]+\d+)?)/i);
    this.entityExtractors.set('worksheet', /sheet\s*(\w+)|worksheet\s*(\w+)/i);
    this.entityExtractors.set('color', /(red|blue|green|yellow|orange|purple|pink|cyan|magenta|gray|grey|black|white)/i);
    this.entityExtractors.set('percentage', /(\d+(?:\.\d+)?)\s*%/i);
    this.entityExtractors.set('number', /(\d+(?:\.\d+)?)/g);
  }

  /**
   * Process user intent and generate operation plan
   */
  async processIntent(
    intent: ProcessedIntent,
    originalMessage: string,
    excelContext: ExcelState
  ): Promise<MCPToolCall[]> {
    console.log('🎯 Processing intent:', intent);

    try {
      // Generate operation plan based on intent
      const plan = await this.generateOperationPlan(intent, originalMessage, excelContext);
      
      console.log('📋 Generated operation plan:', plan);
      
      return plan.operations;
      
    } catch (error: any) {
      console.error('❌ Failed to process intent:', error);
      throw error;
    }
  }

  /**
   * Analyze message to extract intent
   */
  analyzeMessage(message: string, context?: ExcelState): ProcessedIntent {
    const message_lower = message.toLowerCase();
    let bestMatch: ProcessedIntent = {
      type: 'unknown',
      confidence: 0
    };

    // Check each intent pattern
    for (const [intentType, patterns] of this.intentPatterns) {
      let matches = 0;
      for (const pattern of patterns) {
        if (pattern.test(message_lower)) {
          matches++;
        }
      }
      
      const confidence = matches / patterns.length;
      if (confidence > bestMatch.confidence) {
        bestMatch = {
          type: intentType,
          confidence,
          requiresConfirmation: this.requiresConfirmation(intentType, message),
          isMultiStep: this.isMultiStepOperation(message)
        };
      }
    }

    // Extract entities
    bestMatch.parameters = this.extractEntities(message, context);
    
    // Enhance with context
    bestMatch = this.enhanceWithContext(bestMatch, context);

    console.log('🔍 Analyzed intent:', bestMatch);
    return bestMatch;
  }

  /**
   * Generate operation plan from intent
   */
  private async generateOperationPlan(
    intent: ProcessedIntent,
    message: string,
    context: ExcelState
  ): Promise<OperationPlan> {
    switch (intent.type) {
      case 'excel_read':
        return this.planReadOperation(intent, message, context);
        
      case 'excel_write':
        return this.planWriteOperation(intent, message, context);
        
      case 'excel_format':
        return this.planFormatOperation(intent, message, context);
        
      case 'excel_calculate':
        return this.planCalculateOperation(intent, message, context);
        
      case 'excel_analyze':
        return this.planAnalyzeOperation(intent, message, context);
        
      case 'excel_chart':
        return this.planChartOperation(intent, message, context);
        
      default:
        return this.planGenericOperation(intent, message, context);
    }
  }

  /**
   * Plan read operation
   */
  private planReadOperation(intent: ProcessedIntent, message: string, context: ExcelState): OperationPlan {
    const range = intent.parameters?.range || context.selectedRange;
    const worksheet = intent.parameters?.worksheet || context.activeWorksheet;

    return {
      operations: [{
        name: 'excel/read-range',
        arguments: {
          range: range,
          worksheet: worksheet,
          includeFormulas: message.includes('formula'),
          includeFormat: message.includes('format')
        }
      }],
      description: `Read data from ${worksheet}!${range}`,
      estimatedTime: 1000
    };
  }

  /**
   * Plan write operation
   */
  private planWriteOperation(intent: ProcessedIntent, message: string, context: ExcelState): OperationPlan {
    const range = intent.parameters?.range || context.selectedRange;
    const worksheet = intent.parameters?.worksheet || context.activeWorksheet;
    
    // Extract value to write
    let value = this.extractWriteValue(message, intent.parameters);

    return {
      operations: [{
        name: 'excel/write-data',
        arguments: {
          range: range,
          data: value,
          worksheet: worksheet
        }
      }],
      description: `Write "${value}" to ${worksheet}!${range}`,
      estimatedTime: 1500,
      requiresApproval: this.isSignificantChange(value)
    };
  }

  /**
   * Plan format operation
   */
  private planFormatOperation(intent: ProcessedIntent, message: string, context: ExcelState): OperationPlan {
    const range = intent.parameters?.range || context.selectedRange;
    const worksheet = intent.parameters?.worksheet || context.activeWorksheet;
    
    const format = this.extractFormatSettings(message, intent.parameters);

    return {
      operations: [{
        name: 'excel/format-cells',
        arguments: {
          range: range,
          format: format,
          worksheet: worksheet
        }
      }],
      description: `Format cells in ${worksheet}!${range}`,
      estimatedTime: 1200
    };
  }

  /**
   * Plan calculate operation
   */
  private planCalculateOperation(intent: ProcessedIntent, message: string, context: ExcelState): OperationPlan {
    const operations: MCPToolCall[] = [];
    
    if (message.includes('irr') || message.includes('internal rate')) {
      operations.push({
        name: 'excel/calculate-irr',
        arguments: {
          cashFlowRange: intent.parameters?.range || 'CashFlows'
        }
      });
    } else if (message.includes('npv') || message.includes('net present')) {
      operations.push({
        name: 'excel/calculate-npv',
        arguments: {
          cashFlowRange: intent.parameters?.range || 'CashFlows',
          discountRate: intent.parameters?.discountRate || 0.1
        }
      });
    } else {
      // Generic formula
      const formula = this.generateFormula(message, intent.parameters);
      operations.push({
        name: 'excel/write-data',
        arguments: {
          range: intent.parameters?.range || context.selectedRange,
          data: formula,
          worksheet: context.activeWorksheet
        }
      });
    }

    return {
      operations,
      description: 'Perform calculation',
      estimatedTime: 2000
    };
  }

  /**
   * Plan analyze operation
   */
  private planAnalyzeOperation(intent: ProcessedIntent, message: string, context: ExcelState): OperationPlan {
    return {
      operations: [{
        name: 'excel/analyze-data',
        arguments: {
          range: intent.parameters?.range || context.selectedRange,
          analysisType: this.determineAnalysisType(message),
          includeCharts: message.includes('chart') || message.includes('visual')
        }
      }],
      description: 'Analyze data and provide insights',
      estimatedTime: 3000
    };
  }

  /**
   * Plan chart operation
   */
  private planChartOperation(intent: ProcessedIntent, message: string, context: ExcelState): OperationPlan {
    const chartType = this.extractChartType(message);
    
    return {
      operations: [{
        name: 'excel/create-chart',
        arguments: {
          range: intent.parameters?.range || context.selectedRange,
          chartType: chartType,
          title: intent.parameters?.title || 'Chart'
        }
      }],
      description: `Create ${chartType} chart`,
      estimatedTime: 2500
    };
  }

  /**
   * Plan generic operation
   */
  private planGenericOperation(intent: ProcessedIntent, message: string, context: ExcelState): OperationPlan {
    return {
      operations: [{
        name: 'ai/analyze-request',
        arguments: {
          message: message,
          context: context,
          availableTools: Array.from(this.mcpClient.availableTools?.keys() || [])
        }
      }],
      description: 'Analyze request and determine appropriate action',
      estimatedTime: 4000
    };
  }

  /**
   * Extract entities from message
   */
  private extractEntities(message: string, context?: ExcelState): Record<string, any> {
    const entities: Record<string, any> = {};

    for (const [entityType, pattern] of this.entityExtractors) {
      const matches = message.match(pattern);
      if (matches) {
        entities[entityType] = matches[1] || matches[0];
      }
    }

    return entities;
  }

  /**
   * Enhance intent with context
   */
  private enhanceWithContext(intent: ProcessedIntent, context?: ExcelState): ProcessedIntent {
    if (!context) return intent;

    // If no range specified, use current selection
    if (!intent.parameters?.range && context.selectedRange) {
      intent.parameters = {
        ...intent.parameters,
        range: context.selectedRange
      };
    }

    // If no worksheet specified, use active worksheet
    if (!intent.parameters?.worksheet && context.activeWorksheet) {
      intent.parameters = {
        ...intent.parameters,
        worksheet: context.activeWorksheet
      };
    }

    return intent;
  }

  /**
   * Check if operation requires confirmation
   */
  private requiresConfirmation(intentType: string, message: string): boolean {
    // Write operations that affect multiple cells
    if (intentType === 'excel_write' && message.includes('all')) {
      return true;
    }

    // Format operations that change many cells
    if (intentType === 'excel_format' && (message.includes('entire') || message.includes('all'))) {
      return true;
    }

    // Destructive operations
    if (message.includes('delete') || message.includes('clear') || message.includes('remove')) {
      return true;
    }

    return false;
  }

  /**
   * Check if operation is multi-step
   */
  private isMultiStepOperation(message: string): boolean {
    // Look for multiple actions
    const actionWords = ['then', 'after', 'next', 'and', 'also', 'plus'];
    return actionWords.some(word => message.toLowerCase().includes(word));
  }

  /**
   * Extract value to write from message
   */
  private extractWriteValue(message: string, params?: Record<string, any>): any {
    // Look for quoted values
    const quotedMatch = message.match(/["']([^"']+)["']/);
    if (quotedMatch) return quotedMatch[1];

    // Look for numbers
    const numberMatch = message.match(/(\d+(?:\.\d+)?)/);
    if (numberMatch) return Number(numberMatch[1]);

    // Look for formulas
    if (message.includes('=')) {
      const formulaMatch = message.match(/(=.+?)(?:\s|$)/);
      if (formulaMatch) return formulaMatch[1];
    }

    // Extract from parameters
    if (params?.number) return Number(params.number);
    if (params?.percentage) return Number(params.percentage) / 100;

    return 'Value'; // Default
  }

  /**
   * Extract format settings from message
   */
  private extractFormatSettings(message: string, params?: Record<string, any>): Record<string, any> {
    const format: Record<string, any> = {};

    if (params?.color) {
      format.backgroundColor = params.color;
    }

    if (message.includes('bold')) format.bold = true;
    if (message.includes('red')) format.backgroundColor = 'red';
    if (message.includes('green')) format.backgroundColor = 'green';
    if (message.includes('highlight')) format.backgroundColor = 'yellow';

    return format;
  }

  /**
   * Generate formula from message
   */
  private generateFormula(message: string, params?: Record<string, any>): string {
    if (message.includes('sum')) return '=SUM()';
    if (message.includes('average')) return '=AVERAGE()';
    if (message.includes('count')) return '=COUNT()';
    if (message.includes('max')) return '=MAX()';
    if (message.includes('min')) return '=MIN()';

    return '='; // Default formula start
  }

  /**
   * Determine analysis type from message
   */
  private determineAnalysisType(message: string): string {
    if (message.includes('trend')) return 'trend';
    if (message.includes('summary')) return 'summary';
    if (message.includes('correlation')) return 'correlation';
    if (message.includes('regression')) return 'regression';
    
    return 'descriptive';
  }

  /**
   * Extract chart type from message
   */
  private extractChartType(message: string): string {
    if (message.includes('pie')) return 'pie';
    if (message.includes('bar')) return 'bar';
    if (message.includes('line')) return 'line';
    if (message.includes('scatter')) return 'scatter';
    
    return 'column'; // Default
  }

  /**
   * Check if change is significant (requires approval)
   */
  private isSignificantChange(value: any): boolean {
    // Large numbers might be significant
    if (typeof value === 'number' && value > 1000000) return true;
    
    // Formulas are significant
    if (typeof value === 'string' && value.startsWith('=')) return true;
    
    return false;
  }

  /**
   * Get available operations for a message type
   */
  getAvailableOperations(messageType: string): string[] {
    const operations: string[] = [];
    
    switch (messageType) {
      case 'excel_read':
        operations.push('excel/read-range', 'excel/find-data', 'excel/get-worksheet-info');
        break;
      case 'excel_write':
        operations.push('excel/write-data', 'excel/write-formula', 'excel/append-data');
        break;
      case 'excel_format':
        operations.push('excel/format-cells', 'excel/apply-style', 'excel/conditional-format');
        break;
      case 'excel_calculate':
        operations.push('excel/calculate-irr', 'excel/calculate-npv', 'excel/write-data');
        break;
    }
    
    return operations;
  }
}

export default MessageProcessor;