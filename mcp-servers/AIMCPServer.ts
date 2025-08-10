/**
 * AIMCPServer.ts
 * MCP Server for AI operations - handles intent analysis, response generation, and NLP
 */

import {
  MCPServerInfo,
  MCPTool,
  MCPToolCall,
  MCPToolResult,
  MCPContent,
  AIIntent,
  IntentType,
  Entity,
  EntityType,
  ConversationMessage,
  MCPError,
  MCPErrorCode
} from '../mcp-types/interfaces';

import {
  JSONRPCRequest,
  JSONRPCResponse,
  createResponse,
  createError,
  ErrorCodes
} from '../mcp-types/schemas';

export class AIMCPServer {
  private serverInfo: MCPServerInfo = {
    name: 'ai-operations-server',
    version: '1.0.0',
    description: 'MCP server for AI operations including intent analysis, NLP, and response generation',
    capabilities: {
      tools: { listChanged: false },
      resources: {},
      notifications: true
    }
  };

  private tools: Map<string, MCPTool> = new Map();
  private conversationHistory: ConversationMessage[] = [];
  private apiEndpoint: string;

  constructor() {
    this.registerAllTools();
    this.setupAPIEndpoint();
  }

  /**
   * Initialize the server
   */
  async initialize(): Promise<MCPServerInfo> {
    return this.serverInfo;
  }

  /**
   * Set up API endpoint based on environment
   */
  private setupAPIEndpoint(): void {
    const isLocal = typeof window !== 'undefined' && 
                   (window.location?.hostname === 'localhost' || 
                    window.location?.hostname === '127.0.0.1');
    this.apiEndpoint = isLocal 
      ? 'http://localhost:8888/.netlify/functions/chat' 
      : '/.netlify/functions/chat';
  }

  /**
   * Register all AI operation tools
   */
  private registerAllTools(): void {
    // Intent Analysis
    this.registerTool({
      name: 'ai/analyze-intent',
      title: 'Analyze Intent',
      description: 'Analyze user message to determine intent and extract entities',
      inputSchema: {
        type: 'object',
        properties: {
          message: { type: 'string', description: 'User message to analyze' },
          context: { 
            type: 'object', 
            description: 'Conversation context including history and Excel state' 
          }
        },
        required: ['message']
      }
    });

    // Response Generation
    this.registerTool({
      name: 'ai/generate-response',
      title: 'Generate Response',
      description: 'Generate natural language response based on operation results',
      inputSchema: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Original user query' },
          results: { type: 'array', description: 'Results from executed operations' },
          context: { type: 'object', description: 'Conversation context' }
        },
        required: ['query', 'results']
      }
    });

    // Excel Command Translation
    this.registerTool({
      name: 'ai/translate-to-excel',
      title: 'Translate to Excel',
      description: 'Translate natural language request to Excel operations',
      inputSchema: {
        type: 'object',
        properties: {
          request: { type: 'string', description: 'Natural language request' },
          availableTools: { 
            type: 'array', 
            items: { type: 'string' },
            description: 'List of available Excel tools' 
          }
        },
        required: ['request']
      }
    });

    // Data Explanation
    this.registerTool({
      name: 'ai/explain-data',
      title: 'Explain Data',
      description: 'Provide explanation and insights about Excel data',
      inputSchema: {
        type: 'object',
        properties: {
          data: { type: 'object', description: 'Excel data to explain' },
          question: { type: 'string', description: 'Specific question about the data' }
        },
        required: ['data']
      }
    });

    // Formula Generation
    this.registerTool({
      name: 'ai/generate-formula',
      title: 'Generate Formula',
      description: 'Generate Excel formula based on natural language description',
      inputSchema: {
        type: 'object',
        properties: {
          description: { type: 'string', description: 'What the formula should do' },
          context: { 
            type: 'object',
            description: 'Context including available ranges and data types' 
          }
        },
        required: ['description']
      }
    });

    // Model Validation
    this.registerTool({
      name: 'ai/validate-model',
      title: 'Validate Model',
      description: 'Validate financial model for errors and best practices',
      inputSchema: {
        type: 'object',
        properties: {
          modelData: { type: 'object', description: 'Financial model data' },
          validationType: { 
            type: 'string',
            enum: ['formulas', 'logic', 'best-practices', 'all'],
            description: 'Type of validation to perform'
          }
        },
        required: ['modelData']
      }
    });

    // Suggestion Generation
    this.registerTool({
      name: 'ai/suggest-next-action',
      title: 'Suggest Next Action',
      description: 'Suggest next actions based on current context',
      inputSchema: {
        type: 'object',
        properties: {
          currentState: { type: 'object', description: 'Current Excel state' },
          recentActions: { type: 'array', description: 'Recent user actions' },
          userGoal: { type: 'string', description: 'Inferred or stated user goal' }
        },
        required: ['currentState']
      }
    });

    // Clarification
    this.registerTool({
      name: 'ai/request-clarification',
      title: 'Request Clarification',
      description: 'Generate clarification questions when intent is unclear',
      inputSchema: {
        type: 'object',
        properties: {
          ambiguousRequest: { type: 'string', description: 'The unclear request' },
          possibleIntents: { type: 'array', description: 'Possible interpretations' }
        },
        required: ['ambiguousRequest']
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
      const handler = this.getToolHandler(toolCall.name);
      return await handler(toolCall.arguments);
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
      'ai/analyze-intent': this.handleAnalyzeIntent.bind(this),
      'ai/generate-response': this.handleGenerateResponse.bind(this),
      'ai/translate-to-excel': this.handleTranslateToExcel.bind(this),
      'ai/explain-data': this.handleExplainData.bind(this),
      'ai/generate-formula': this.handleGenerateFormula.bind(this),
      'ai/validate-model': this.handleValidateModel.bind(this),
      'ai/suggest-next-action': this.handleSuggestNextAction.bind(this),
      'ai/request-clarification': this.handleRequestClarification.bind(this)
    };

    return handlers[toolName] || (() => Promise.reject(new Error('Handler not implemented')));
  }

  // ===== Tool Handlers =====

  private async handleAnalyzeIntent(args: any): Promise<MCPToolResult> {
    const { message, context } = args;
    
    // Analyze the message for intent
    const intent = await this.analyzeUserIntent(message, context);
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify(intent, null, 2)
      }]
    };
  }

  private async handleGenerateResponse(args: any): Promise<MCPToolResult> {
    const { query, results, context } = args;
    
    // Build prompt for response generation
    const prompt = this.buildResponsePrompt(query, results, context);
    
    // Call AI service
    const response = await this.callAIService(prompt, 'chat');
    
    return {
      content: [{
        type: 'text',
        text: response
      }]
    };
  }

  private async handleTranslateToExcel(args: any): Promise<MCPToolResult> {
    const { request, availableTools } = args;
    
    const prompt = `
      Translate this natural language request into Excel operations:
      "${request}"
      
      Available tools: ${JSON.stringify(availableTools || [])}
      
      Return a JSON array of operations to perform, with tool names and arguments.
      Example: [{"tool": "excel/write-value", "args": {"range": "A1", "value": 100}}]
    `;
    
    const response = await this.callAIService(prompt, 'translation');
    
    try {
      const operations = JSON.parse(response);
      return {
        content: [{
          type: 'text',
          text: JSON.stringify(operations, null, 2)
        }]
      };
    } catch (error) {
      return {
        content: [{
          type: 'text',
          text: response
        }]
      };
    }
  }

  private async handleExplainData(args: any): Promise<MCPToolResult> {
    const { data, question } = args;
    
    const prompt = `
      Explain this Excel data:
      ${JSON.stringify(data, null, 2)}
      
      ${question ? `Specific question: ${question}` : 'Provide insights and observations.'}
      
      Be specific and reference actual values from the data.
    `;
    
    const response = await this.callAIService(prompt, 'explanation');
    
    return {
      content: [{
        type: 'text',
        text: response
      }]
    };
  }

  private async handleGenerateFormula(args: any): Promise<MCPToolResult> {
    const { description, context } = args;
    
    const prompt = `
      Generate an Excel formula for: "${description}"
      
      Context: ${JSON.stringify(context || {}, null, 2)}
      
      Return only the formula, starting with =
      Include a brief explanation of how it works.
    `;
    
    const response = await this.callAIService(prompt, 'formula');
    
    return {
      content: [{
        type: 'text',
        text: response
      }]
    };
  }

  private async handleValidateModel(args: any): Promise<MCPToolResult> {
    const { modelData, validationType = 'all' } = args;
    
    const prompt = `
      Validate this financial model for ${validationType}:
      ${JSON.stringify(modelData, null, 2)}
      
      Check for:
      - Formula errors
      - Circular references
      - Best practices violations
      - Logical inconsistencies
      
      Return specific issues found with recommendations.
    `;
    
    const response = await this.callAIService(prompt, 'validation');
    
    return {
      content: [{
        type: 'text',
        text: response
      }]
    };
  }

  private async handleSuggestNextAction(args: any): Promise<MCPToolResult> {
    const { currentState, recentActions, userGoal } = args;
    
    const prompt = `
      Based on the current Excel state and recent actions, suggest next steps:
      
      Current State: ${JSON.stringify(currentState, null, 2)}
      Recent Actions: ${JSON.stringify(recentActions || [], null, 2)}
      ${userGoal ? `User Goal: ${userGoal}` : ''}
      
      Provide 3-5 actionable suggestions.
    `;
    
    const response = await this.callAIService(prompt, 'suggestion');
    
    return {
      content: [{
        type: 'text',
        text: response
      }]
    };
  }

  private async handleRequestClarification(args: any): Promise<MCPToolResult> {
    const { ambiguousRequest, possibleIntents } = args;
    
    const prompt = `
      The user's request is unclear: "${ambiguousRequest}"
      
      Possible interpretations: ${JSON.stringify(possibleIntents || [], null, 2)}
      
      Generate a clarifying question to understand what the user wants.
      Be specific and provide options if helpful.
    `;
    
    const response = await this.callAIService(prompt, 'clarification');
    
    return {
      content: [{
        type: 'text',
        text: response
      }]
    };
  }

  // ===== Helper Methods =====

  /**
   * Analyze user intent from message
   */
  private async analyzeUserIntent(message: string, context?: any): Promise<AIIntent> {
    // Pattern matching for common intents
    const patterns: Record<IntentType, RegExp[]> = {
      'read-data': [/show|display|what|get|read|find/i],
      'write-data': [/set|write|change|update|modify|enter/i],
      'format-cells': [/format|color|highlight|style|bold|italic/i],
      'calculate': [/calculate|sum|average|count|total/i],
      'create-chart': [/chart|graph|plot|visualize/i],
      'analyze-data': [/analyze|explain|insight|trend|pattern/i],
      'validate-model': [/validate|check|verify|audit|review/i],
      'explain-formula': [/explain|how|what.*formula|understand/i],
      'find-errors': [/error|mistake|wrong|issue|problem/i],
      'optimize-model': [/optimize|improve|enhance|better|faster/i]
    };

    let detectedIntent: IntentType = 'read-data';
    let confidence = 0;

    for (const [intent, regexes] of Object.entries(patterns)) {
      for (const regex of regexes) {
        if (regex.test(message)) {
          detectedIntent = intent as IntentType;
          confidence = 0.8;
          break;
        }
      }
    }

    // Extract entities
    const entities = this.extractEntities(message);

    // Determine suggested tools based on intent
    const suggestedTools = this.getSuggestedTools(detectedIntent);

    return {
      type: detectedIntent,
      confidence,
      entities,
      suggestedTools
    };
  }

  /**
   * Extract entities from message
   */
  private extractEntities(message: string): Entity[] {
    const entities: Entity[] = [];

    // Range detection (e.g., A1:B10, A1, Sheet1!A1:B10)
    const rangeRegex = /([A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?)/gi;
    const rangeMatches = message.match(rangeRegex);
    if (rangeMatches) {
      rangeMatches.forEach(match => {
        entities.push({
          type: 'range' as EntityType,
          value: match,
          confidence: 0.9
        });
      });
    }

    // Color detection
    const colorRegex = /\b(red|blue|green|yellow|orange|purple|black|white|gray|grey)\b/gi;
    const colorMatches = message.match(colorRegex);
    if (colorMatches) {
      colorMatches.forEach(match => {
        entities.push({
          type: 'color' as EntityType,
          value: match.toLowerCase(),
          confidence: 0.8,
          normalizedValue: this.colorToHex(match.toLowerCase())
        });
      });
    }

    // Number detection
    const numberRegex = /\b\d+(?:\.\d+)?%?\b/g;
    const numberMatches = message.match(numberRegex);
    if (numberMatches) {
      numberMatches.forEach(match => {
        entities.push({
          type: 'value' as EntityType,
          value: match,
          confidence: 0.9,
          normalizedValue: parseFloat(match.replace('%', '')) / (match.includes('%') ? 100 : 1)
        });
      });
    }

    // Financial metric detection
    const metricRegex = /\b(MOIC|IRR|NPV|ROI|EBITDA|revenue|profit|cost)\b/gi;
    const metricMatches = message.match(metricRegex);
    if (metricMatches) {
      metricMatches.forEach(match => {
        entities.push({
          type: 'metric-name' as EntityType,
          value: match.toUpperCase(),
          confidence: 0.9
        });
      });
    }

    return entities;
  }

  /**
   * Get suggested tools for an intent
   */
  private getSuggestedTools(intent: IntentType): string[] {
    const toolMap: Record<IntentType, string[]> = {
      'read-data': ['excel/read-range', 'excel/find-data'],
      'write-data': ['excel/write-value', 'excel/write-formula'],
      'format-cells': ['excel/apply-format', 'excel/apply-conditional-format'],
      'calculate': ['excel/calculate-irr', 'excel/calculate-npv', 'excel/calculate-moic'],
      'create-chart': ['excel/create-chart'],
      'analyze-data': ['ai/explain-data', 'excel/get-financial-metrics'],
      'validate-model': ['ai/validate-model'],
      'explain-formula': ['ai/explain-data'],
      'find-errors': ['ai/validate-model'],
      'optimize-model': ['ai/suggest-next-action']
    };

    return toolMap[intent] || [];
  }

  /**
   * Build prompt for response generation
   */
  private buildResponsePrompt(query: string, results: any[], context?: any): string {
    return `
      User Query: "${query}"
      
      Operation Results:
      ${JSON.stringify(results, null, 2)}
      
      Context:
      ${context ? JSON.stringify(context, null, 2) : 'No additional context'}
      
      Generate a clear, conversational response that:
      1. Confirms what was done
      2. Highlights key results
      3. Suggests next steps if appropriate
      
      Be specific and reference actual values from the results.
    `;
  }

  /**
   * Call AI service (OpenAI via Netlify function)
   */
  private async callAIService(prompt: string, mode: string): Promise<string> {
    try {
      const response = await fetch(this.apiEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        body: JSON.stringify({
          message: prompt,
          batchType: 'chat',
          systemPrompt: this.getSystemPrompt(mode),
          temperature: 0.7,
          maxTokens: 2000
        })
      });

      if (!response.ok) {
        throw new Error(`API error: ${response.status}`);
      }

      const result = await response.json();
      
      if (result.error) {
        throw new Error(result.error);
      }

      return result.content || result.response || 'No response generated';
    } catch (error: any) {
      console.error('AI service error:', error);
      throw new MCPError(
        `AI service failed: ${error.message}`,
        MCPErrorCode.AI_SERVICE_ERROR
      );
    }
  }

  /**
   * Get system prompt based on mode
   */
  private getSystemPrompt(mode: string): string {
    const prompts: Record<string, string> = {
      chat: 'You are an expert Excel and M&A financial modeling assistant. Provide clear, conversational responses.',
      translation: 'You are an Excel operation translator. Convert natural language to specific Excel operations.',
      explanation: 'You are a data analyst. Explain Excel data clearly with specific insights.',
      formula: 'You are an Excel formula expert. Generate accurate, efficient formulas.',
      validation: 'You are a financial model auditor. Identify issues and provide specific recommendations.',
      suggestion: 'You are a productivity assistant. Suggest helpful next actions based on context.',
      clarification: 'You are a helpful assistant. Ask clear questions to understand user intent.'
    };

    return prompts[mode] || prompts.chat;
  }

  /**
   * Convert color name to hex
   */
  private colorToHex(color: string): string {
    const colors: Record<string, string> = {
      red: '#FF0000',
      blue: '#0000FF',
      green: '#00FF00',
      yellow: '#FFFF00',
      orange: '#FFA500',
      purple: '#800080',
      black: '#000000',
      white: '#FFFFFF',
      gray: '#808080',
      grey: '#808080'
    };

    return colors[color] || '#000000';
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
export default AIMCPServer;