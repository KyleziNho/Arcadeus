/**
 * MCPChatClient.ts
 * Main MCP client that manages connections to MCP servers and handles chat interactions
 */

import {
  MCPClientInfo,
  MCPServerInfo,
  MCPSession,
  MCPTool,
  MCPToolCall,
  MCPToolResult,
  ConversationContext,
  ConversationMessage,
  ExcelState,
  MCPError,
  MCPErrorCode
} from '../mcp-types/interfaces';

import {
  JSONRPCRequest,
  JSONRPCResponse,
  createRequest,
  generateId
} from '../mcp-types/schemas';

import { ConversationManager } from './ConversationManager';
import { MessageProcessor } from './MessageProcessor';
import { OperationHistory } from './OperationHistory';
import SamplingHandler from './SamplingHandler';
import ElicitationHandler from './ElicitationHandler';
import ExcelMCPServer from '../mcp-servers/ExcelMCPServer';
import AIMCPServer from '../mcp-servers/AIMCPServer';

export class MCPChatClient {
  private clientInfo: MCPClientInfo = {
    name: 'excel-addin-chat',
    version: '1.0.0',
    capabilities: {
      tools: { listChanged: true },
      resources: { listChanged: true },
      notifications: true
    }
  };

  private sessions: Map<string, MCPSession> = new Map();
  private excelServer: ExcelMCPServer | null = null;
  private aiServer: AIMCPServer | null = null;
  private conversationManager: ConversationManager;
  private messageProcessor: MessageProcessor;
  private operationHistory: OperationHistory;
  private samplingHandler: SamplingHandler;
  private elicitationHandler: ElicitationHandler;
  private isInitialized: boolean = false;
  private availableTools: Map<string, MCPTool> = new Map();
  private roots: string[] = []; // Filesystem roots for servers

  constructor() {
    this.conversationManager = new ConversationManager();
    this.messageProcessor = new MessageProcessor(this);
    this.operationHistory = new OperationHistory();
    this.samplingHandler = new SamplingHandler();
    this.elicitationHandler = new ElicitationHandler();
    
    // Set up sampling approval callback
    this.samplingHandler.setApprovalCallback(this.handleSamplingApproval.bind(this));
  }

  /**
   * Initialize the MCP client and connect to servers
   */
  async initialize(): Promise<void> {
    if (this.isInitialized) return;

    console.log('🚀 Initializing MCP Chat Client...');

    try {
      // Initialize local MCP servers
      await this.initializeLocalServers();
      
      // Discover available tools
      await this.discoverTools();
      
      // Set up Excel event listeners
      await this.setupExcelEventListeners();
      
      this.isInitialized = true;
      console.log('✅ MCP Chat Client initialized successfully');
      
    } catch (error) {
      console.error('❌ Failed to initialize MCP Chat Client:', error);
      throw new MCPError(
        'Failed to initialize MCP client',
        MCPErrorCode.CONNECTION_FAILED,
        error
      );
    }
  }

  /**
   * Initialize local MCP servers (Excel and AI)
   */
  private async initializeLocalServers(): Promise<void> {
    // Initialize Excel MCP Server
    this.excelServer = new ExcelMCPServer();
    const excelServerInfo = await this.excelServer.initialize();
    
    // Create session for Excel server
    const excelSession: MCPSession = {
      id: generateId(),
      client: this.clientInfo,
      server: excelServerInfo,
      transport: this.createLocalTransport('excel'),
      context: this.conversationManager.getContext(),
      isConnected: true,
      createdAt: new Date(),
      lastActivity: new Date()
    };
    
    this.sessions.set('excel', excelSession);
    console.log('✅ Excel MCP Server connected');

    // Initialize AI MCP Server
    this.aiServer = new AIMCPServer();
    const aiServerInfo = await this.aiServer.initialize();
    
    // Create session for AI server
    const aiSession: MCPSession = {
      id: generateId(),
      client: this.clientInfo,
      server: aiServerInfo,
      transport: this.createLocalTransport('ai'),
      context: this.conversationManager.getContext(),
      isConnected: true,
      createdAt: new Date(),
      lastActivity: new Date()
    };
    
    this.sessions.set('ai', aiSession);
    console.log('✅ AI MCP Server connected');
  }

  /**
   * Create a local transport for in-process communication
   */
  private createLocalTransport(serverType: string): any {
    return {
      type: 'local',
      send: async (message: any) => {
        // Direct method call to local server
        if (serverType === 'excel' && this.excelServer) {
          return await this.excelServer.handleRequest(message);
        } else if (serverType === 'ai' && this.aiServer) {
          return await this.aiServer.handleRequest(message);
        }
      },
      receive: async function* () {
        // For local transport, we handle notifications directly
      },
      close: async () => {
        console.log(`Closing ${serverType} transport`);
      }
    };
  }

  /**
   * Discover all available tools from connected servers
   */
  private async discoverTools(): Promise<void> {
    console.log('🔍 Discovering available tools...');
    
    // Get tools from Excel server
    if (this.excelServer) {
      const excelTools = await this.excelServer.listTools();
      excelTools.forEach(tool => {
        this.availableTools.set(tool.name, tool);
      });
      console.log(`📊 Found ${excelTools.length} Excel tools`);
    }
    
    // Get tools from AI server
    if (this.aiServer) {
      const aiTools = await this.aiServer.listTools();
      aiTools.forEach(tool => {
        this.availableTools.set(tool.name, tool);
      });
      console.log(`🤖 Found ${aiTools.length} AI tools`);
    }
    
    console.log(`✅ Total tools available: ${this.availableTools.size}`);
  }

  /**
   * Process a user message through MCP
   */
  async processUserMessage(message: string): Promise<string> {
    console.log('💬 Processing user message:', message);
    
    try {
      // Add message to conversation history
      this.conversationManager.addMessage('user', message);
      
      // Get current Excel context
      const excelContext = await this.getCurrentExcelContext();
      
      // Analyze intent using AI server
      const intentResult = await this.callTool('ai/analyze-intent', {
        message: message,
        context: {
          conversation: this.conversationManager.getHistory(),
          excel: excelContext
        }
      });
      
      const intent = JSON.parse(intentResult.content[0].text!);
      console.log('🎯 Detected intent:', intent);
      
      // Process the request based on intent
      const operations = await this.messageProcessor.processIntent(
        intent,
        message,
        excelContext
      );
      
      // Execute operations
      const results = await this.executeOperations(operations);
      
      // Generate response using AI
      const responseResult = await this.callTool('ai/generate-response', {
        query: message,
        results: results,
        context: this.conversationManager.getContext()
      });
      
      const response = responseResult.content[0].text || 'Operation completed successfully';
      
      // Add response to conversation history
      this.conversationManager.addMessage('assistant', response, {
        toolCalls: operations,
        toolResults: results
      });
      
      return response;
      
    } catch (error: any) {
      console.error('❌ Error processing message:', error);
      const errorMessage = `I encountered an error: ${error.message}. Please try again.`;
      this.conversationManager.addMessage('assistant', errorMessage);
      return errorMessage;
    }
  }

  /**
   * Call a specific MCP tool
   */
  async callTool(toolName: string, args: any): Promise<MCPToolResult> {
    console.log(`🔧 Calling tool: ${toolName}`);
    
    // Determine which server handles this tool
    const server = toolName.startsWith('excel/') ? this.excelServer : 
                  toolName.startsWith('ai/') ? this.aiServer : null;
    
    if (!server) {
      throw new MCPError(
        `No server found for tool: ${toolName}`,
        MCPErrorCode.METHOD_NOT_FOUND
      );
    }
    
    const toolCall: MCPToolCall = {
      name: toolName,
      arguments: args
    };
    
    const result = await server.executeTool(toolCall);
    
    // Record operation for undo/redo if it's an Excel operation
    if (toolName.startsWith('excel/') && !toolName.includes('read') && !toolName.includes('find')) {
      this.operationHistory.recordOperation({
        id: generateId(),
        type: 'write-value',
        timestamp: new Date(),
        description: `Execute ${toolName}`,
        params: args,
        status: result.isError ? 'failed' : 'completed',
        error: result.errorMessage
      });
    }
    
    return result;
  }

  /**
   * Execute a list of operations
   */
  private async executeOperations(operations: MCPToolCall[]): Promise<MCPToolResult[]> {
    const results: MCPToolResult[] = [];
    
    for (const operation of operations) {
      try {
        const result = await this.callTool(operation.name, operation.arguments);
        results.push(result);
      } catch (error: any) {
        results.push({
          content: [{
            type: 'error',
            text: error.message
          }],
          isError: true,
          errorMessage: error.message
        });
      }
    }
    
    return results;
  }

  /**
   * Get current Excel context
   */
  private async getCurrentExcelContext(): Promise<ExcelState> {
    try {
      // Get selected range
      const selectedRangeResult = await this.callTool('excel/read-range', {
        range: 'SELECTION',
        includeFormulas: true,
        includeFormat: true
      });
      
      // Parse the result
      const selectedData = JSON.parse(selectedRangeResult.content[0].text!);
      
      return {
        activeWorksheet: selectedData.worksheet || 'Sheet1',
        selectedRange: selectedData.address || 'A1',
        worksheets: [], // Would need to get worksheet list
        namedRanges: [], // Would need to get named ranges
        recentChanges: this.operationHistory.getRecentChanges()
      };
    } catch (error) {
      console.error('Error getting Excel context:', error);
      return {
        activeWorksheet: 'Sheet1',
        selectedRange: 'A1',
        worksheets: [],
        namedRanges: [],
        recentChanges: []
      };
    }
  }

  /**
   * Set up Excel event listeners
   */
  private async setupExcelEventListeners(): Promise<void> {
    if (typeof Excel === 'undefined') {
      console.warn('Excel API not available, skipping event listeners');
      return;
    }
    
    try {
      await Excel.run(async (context) => {
        // Listen for selection changes
        context.workbook.onSelectionChanged.add(async (event) => {
          console.log('📍 Selection changed:', event.address);
          this.conversationManager.updateExcelContext({
            selectedRange: event.address
          });
        });
        
        // Listen for data changes
        context.workbook.worksheets.onChanged.add(async (event) => {
          console.log('📝 Worksheet changed:', event);
          // Could trigger notifications here
        });
        
        await context.sync();
      });
      
      console.log('✅ Excel event listeners set up');
    } catch (error) {
      console.error('Failed to set up Excel event listeners:', error);
    }
  }

  /**
   * Undo last operation
   */
  async undo(): Promise<boolean> {
    const operation = this.operationHistory.getLastOperation();
    if (!operation || !operation.inverse) {
      return false;
    }
    
    try {
      // Execute the inverse operation
      await this.callTool(operation.inverse.type, operation.inverse.params);
      this.operationHistory.markAsUndone(operation.id);
      return true;
    } catch (error) {
      console.error('Undo failed:', error);
      return false;
    }
  }

  /**
   * Redo last undone operation
   */
  async redo(): Promise<boolean> {
    const operation = this.operationHistory.getLastUndoneOperation();
    if (!operation) {
      return false;
    }
    
    try {
      // Re-execute the operation
      await this.callTool(operation.type, operation.params);
      this.operationHistory.markAsRedone(operation.id);
      return true;
    } catch (error) {
      console.error('Redo failed:', error);
      return false;
    }
  }

  /**
   * Get available tools
   */
  getAvailableTools(): MCPTool[] {
    return Array.from(this.availableTools.values());
  }

  /**
   * Get conversation history
   */
  getConversationHistory(): ConversationMessage[] {
    return this.conversationManager.getHistory();
  }

  /**
   * Clear conversation
   */
  clearConversation(): void {
    this.conversationManager.clear();
    this.operationHistory.clear();
  }

  /**
   * Export conversation
   */
  exportConversation(): string {
    return JSON.stringify({
      conversation: this.conversationManager.getHistory(),
      operations: this.operationHistory.getAllOperations(),
      timestamp: new Date().toISOString()
    }, null, 2);
  }

  /**
   * Handle sampling approval requests (human-in-the-loop)
   */
  private async handleSamplingApproval(request: any): Promise<any> {
    return new Promise((resolve) => {
      // Create approval UI
      const approvalUI = this.samplingHandler.createApprovalUI(request);
      document.body.appendChild(approvalUI);
      
      // Set up global handlers for approval
      (window as any).approveSampling = () => {
        approvalUI.remove();
        resolve({ approved: true });
      };
      
      (window as any).denySampling = () => {
        approvalUI.remove();
        resolve({ approved: false, reason: 'User denied request' });
      };
      
      (window as any).modifySampling = () => {
        // In a real implementation, this would open a modification UI
        approvalUI.remove();
        resolve({ approved: true, modifiedRequest: request });
      };
    });
  }

  /**
   * Handle sampling requests from servers
   */
  async handleSamplingRequest(serverId: string, request: any): Promise<any> {
    const context = this.conversationManager.getHistory();
    return await this.samplingHandler.handleSamplingRequest(serverId, request, context);
  }

  /**
   * Handle elicitation requests from servers
   */
  async handleElicitationRequest(serverId: string, request: any): Promise<any> {
    return await this.elicitationHandler.handleElicitationRequest(serverId, request);
  }

  /**
   * Update filesystem roots
   */
  updateRoots(newRoots: string[]): void {
    this.roots = [...newRoots];
    console.log('📁 Updated filesystem roots:', this.roots);
    
    // Notify servers about root changes
    this.notifyServersOfRootChanges();
  }

  /**
   * Add a root directory
   */
  addRoot(rootPath: string): void {
    if (!this.roots.includes(rootPath)) {
      this.roots.push(rootPath);
      this.notifyServersOfRootChanges();
    }
  }

  /**
   * Remove a root directory
   */
  removeRoot(rootPath: string): void {
    const index = this.roots.indexOf(rootPath);
    if (index > -1) {
      this.roots.splice(index, 1);
      this.notifyServersOfRootChanges();
    }
  }

  /**
   * Get current filesystem roots
   */
  getRoots(): string[] {
    return [...this.roots];
  }

  /**
   * Notify servers of root changes
   */
  private notifyServersOfRootChanges(): void {
    // In a full implementation, this would send notifications to servers
    console.log('🔄 Notifying servers of root changes');
  }

  /**
   * Enhanced client capabilities declaration
   */
  getClientCapabilities(): any {
    return {
      ...this.clientInfo.capabilities,
      sampling: {
        supportedModels: ['gpt-3.5-turbo', 'gpt-4'],
        humanInTheLoop: true
      },
      elicitation: {
        supportedTypes: ['text', 'number', 'boolean', 'select', 'multi-select']
      },
      roots: {
        supported: true,
        listChanged: true
      }
    };
  }

  /**
   * Process message with enhanced MCP features
   */
  async processMessageWithMCPFeatures(message: string): Promise<string> {
    console.log('💬 Processing message with full MCP features:', message);
    
    try {
      // Check if this is a request that requires sampling or elicitation
      const requiresAdvancedFeatures = await this.analyzeMessageComplexity(message);
      
      if (requiresAdvancedFeatures.needsSampling) {
        // Handle complex analysis that requires AI sampling
        return await this.processWithSampling(message, requiresAdvancedFeatures);
      }
      
      if (requiresAdvancedFeatures.needsElicitation) {
        // Handle requests that need user input
        return await this.processWithElicitation(message, requiresAdvancedFeatures);
      }
      
      // Standard processing
      return await this.processUserMessage(message);
      
    } catch (error: any) {
      console.error('❌ Enhanced processing failed:', error);
      return `I encountered an error with enhanced processing: ${error.message}`;
    }
  }

  /**
   * Analyze message complexity to determine if advanced features are needed
   */
  private async analyzeMessageComplexity(message: string): Promise<any> {
    // Simple heuristics - in practice, this could use AI
    return {
      needsSampling: message.includes('analyze') || message.includes('complex') || message.includes('optimize'),
      needsElicitation: message.includes('missing') || message.includes('need to know') || message.includes('configure'),
      complexity: 'standard'
    };
  }

  /**
   * Process with sampling (for complex AI tasks)
   */
  private async processWithSampling(message: string, analysis: any): Promise<string> {
    console.log('🤖 Processing with sampling...');
    
    // Create sampling request
    const samplingRequest = {
      messages: [{ role: 'user', content: message }],
      systemPrompt: 'You are an expert Excel and M&A analyst. Provide detailed analysis.',
      maxTokens: 2000,
      modelPreferences: {
        intelligencePriority: 0.9,
        speedPriority: 0.3,
        costPriority: 0.5
      }
    };
    
    const result = await this.handleSamplingRequest('ai', samplingRequest);
    return result[0]?.text || 'Analysis completed via sampling';
  }

  /**
   * Process with elicitation (when user input is needed)
   */
  private async processWithElicitation(message: string, analysis: any): Promise<string> {
    console.log('📝 Processing with elicitation...');
    
    // Create elicitation request
    const elicitationRequest = {
      method: 'elicitation/requestInput',
      params: {
        message: 'I need some additional information to complete this request:',
        schema: {
          type: 'object',
          properties: {
            specificRange: {
              type: 'string',
              description: 'Which Excel range should I focus on?'
            },
            analysisType: {
              type: 'string',
              enum: ['basic', 'detailed', 'comprehensive'],
              description: 'Level of analysis required'
            }
          },
          required: ['specificRange']
        }
      }
    };
    
    const response = await this.handleElicitationRequest('ai', elicitationRequest);
    
    if (response.cancelled) {
      return 'Request cancelled by user';
    }
    
    // Continue processing with elicited information
    const enhancedMessage = `${message} (focusing on ${response.data?.specificRange} with ${response.data?.analysisType} analysis)`;
    return await this.processUserMessage(enhancedMessage);
  }

  /**
   * Cleanup and disconnect
   */
  async disconnect(): Promise<void> {
    for (const [name, session] of this.sessions) {
      await session.transport.close();
      console.log(`Disconnected from ${name} server`);
    }
    this.sessions.clear();
    this.isInitialized = false;
  }
}

// Export for use in the application
export default MCPChatClient;