/**
 * ConversationManager.ts
 * Manages conversation state, history, and context for MCP chat system
 */

import {
  ConversationMessage,
  ConversationContext,
  ExcelState,
  MCPToolCall,
  MCPToolResult
} from '../mcp-types/interfaces';

export interface ConversationMetadata {
  toolCalls?: MCPToolCall[];
  toolResults?: MCPToolResult[];
  excelOperations?: string[];
  userId?: string;
  sessionId?: string;
}

export class ConversationManager {
  private messages: ConversationMessage[] = [];
  private maxMessages: number = 100; // Keep last 100 messages
  private context: ConversationContext;
  private sessionId: string;
  private startTime: Date;

  constructor() {
    this.sessionId = this.generateSessionId();
    this.startTime = new Date();
    this.context = this.initializeContext();
  }

  /**
   * Initialize conversation context
   */
  private initializeContext(): ConversationContext {
    return {
      sessionId: this.sessionId,
      startTime: this.startTime,
      totalMessages: 0,
      excel: {
        activeWorksheet: 'Sheet1',
        selectedRange: 'A1',
        worksheets: [],
        namedRanges: [],
        recentChanges: []
      },
      preferences: {
        verbosity: 'normal',
        autoSave: true,
        confirmActions: true
      },
      capabilities: {
        canReadExcel: true,
        canWriteExcel: true,
        canFormatExcel: true,
        canAnalyzeData: true
      }
    };
  }

  /**
   * Add a message to the conversation history
   */
  addMessage(
    role: 'user' | 'assistant' | 'system',
    content: string,
    metadata?: ConversationMetadata
  ): void {
    const message: ConversationMessage = {
      id: this.generateMessageId(),
      role,
      content,
      timestamp: new Date(),
      metadata: metadata || {}
    };

    this.messages.push(message);
    this.context.totalMessages++;

    // Maintain message limit
    if (this.messages.length > this.maxMessages) {
      this.messages = this.messages.slice(-this.maxMessages);
    }

    // Update context based on message
    this.updateContextFromMessage(message);

    console.log(`📝 Added ${role} message to conversation`, message);
  }

  /**
   * Get conversation history
   */
  getHistory(): ConversationMessage[] {
    return [...this.messages]; // Return copy to prevent external modification
  }

  /**
   * Get recent messages (last N)
   */
  getRecentMessages(count: number = 10): ConversationMessage[] {
    return this.messages.slice(-count);
  }

  /**
   * Get conversation context
   */
  getContext(): ConversationContext {
    return {
      ...this.context,
      // Update dynamic values
      totalMessages: this.messages.length,
      lastActivity: this.messages.length > 0 ? this.messages[this.messages.length - 1].timestamp : this.startTime
    };
  }

  /**
   * Update Excel context
   */
  updateExcelContext(excelUpdate: Partial<ExcelState>): void {
    this.context.excel = {
      ...this.context.excel,
      ...excelUpdate
    };

    console.log('📊 Updated Excel context:', excelUpdate);
  }

  /**
   * Update context from message content and metadata
   */
  private updateContextFromMessage(message: ConversationMessage): void {
    // Track Excel operations
    if (message.metadata?.toolCalls) {
      const excelOps = message.metadata.toolCalls
        .filter(call => call.name.startsWith('excel/'))
        .map(call => call.name);
      
      if (excelOps.length > 0) {
        this.context.excel.recentChanges = [
          ...this.context.excel.recentChanges.slice(-9), // Keep last 9
          ...excelOps.map(op => ({
            operation: op,
            timestamp: message.timestamp,
            messageId: message.id
          }))
        ];
      }
    }

    // Update preferences based on user patterns
    if (message.role === 'user') {
      this.analyzeUserPreferences(message);
    }
  }

  /**
   * Analyze user message patterns to infer preferences
   */
  private analyzeUserPreferences(message: ConversationMessage): void {
    const content = message.content.toLowerCase();
    
    // Verbosity preference
    if (content.includes('explain') || content.includes('detail') || content.includes('why')) {
      this.context.preferences.verbosity = 'detailed';
    } else if (content.includes('just') || content.includes('quick') || content.includes('briefly')) {
      this.context.preferences.verbosity = 'concise';
    }

    // Confirmation preference
    if (content.includes('just do it') || content.includes('go ahead') || content.includes('no need to ask')) {
      this.context.preferences.confirmActions = false;
    } else if (content.includes('check with me') || content.includes('ask first') || content.includes('confirm')) {
      this.context.preferences.confirmActions = true;
    }
  }

  /**
   * Search conversation history
   */
  searchHistory(query: string): ConversationMessage[] {
    const lowerQuery = query.toLowerCase();
    return this.messages.filter(message => 
      message.content.toLowerCase().includes(lowerQuery) ||
      message.metadata?.toolCalls?.some(call => 
        call.name.toLowerCase().includes(lowerQuery)
      )
    );
  }

  /**
   * Get messages by role
   */
  getMessagesByRole(role: 'user' | 'assistant' | 'system'): ConversationMessage[] {
    return this.messages.filter(message => message.role === role);
  }

  /**
   * Get conversation summary
   */
  getConversationSummary(): {
    totalMessages: number;
    userMessages: number;
    assistantMessages: number;
    systemMessages: number;
    duration: number;
    excelOperations: number;
    topicsDiscussed: string[];
  } {
    const userMessages = this.getMessagesByRole('user').length;
    const assistantMessages = this.getMessagesByRole('assistant').length;
    const systemMessages = this.getMessagesByRole('system').length;
    
    const duration = Date.now() - this.startTime.getTime();
    
    const excelOperations = this.context.excel.recentChanges.length;
    
    // Extract key topics from user messages
    const topicsDiscussed = this.extractTopics(this.getMessagesByRole('user'));

    return {
      totalMessages: this.messages.length,
      userMessages,
      assistantMessages,
      systemMessages,
      duration,
      excelOperations,
      topicsDiscussed
    };
  }

  /**
   * Extract topics from user messages
   */
  private extractTopics(userMessages: ConversationMessage[]): string[] {
    const topics = new Set<string>();
    const keywords = [
      'format', 'calculate', 'analyze', 'chart', 'graph', 'pivot', 'formula',
      'sum', 'average', 'count', 'vlookup', 'filter', 'sort', 'color',
      'revenue', 'cost', 'profit', 'irr', 'npv', 'valuation', 'forecast',
      'debt', 'equity', 'cash flow', 'ebitda', 'assumptions'
    ];

    userMessages.forEach(message => {
      const content = message.content.toLowerCase();
      keywords.forEach(keyword => {
        if (content.includes(keyword)) {
          topics.add(keyword);
        }
      });
    });

    return Array.from(topics).slice(0, 10); // Return top 10 topics
  }

  /**
   * Clear conversation history
   */
  clear(): void {
    this.messages = [];
    this.context = this.initializeContext();
    console.log('🗑️ Conversation cleared');
  }

  /**
   * Export conversation for backup/analysis
   */
  export(): string {
    return JSON.stringify({
      sessionId: this.sessionId,
      startTime: this.startTime,
      messages: this.messages,
      context: this.context,
      summary: this.getConversationSummary(),
      exportTime: new Date().toISOString()
    }, null, 2);
  }

  /**
   * Import conversation from backup
   */
  import(conversationData: string): void {
    try {
      const data = JSON.parse(conversationData);
      this.sessionId = data.sessionId || this.generateSessionId();
      this.startTime = new Date(data.startTime) || new Date();
      this.messages = data.messages || [];
      this.context = data.context || this.initializeContext();
      
      console.log('📥 Conversation imported successfully');
    } catch (error) {
      console.error('❌ Failed to import conversation:', error);
      throw new Error('Invalid conversation data format');
    }
  }

  /**
   * Get context for AI reasoning
   */
  getContextForAI(): {
    recentMessages: ConversationMessage[];
    excelState: ExcelState;
    userPreferences: any;
    conversationSummary: any;
  } {
    return {
      recentMessages: this.getRecentMessages(5),
      excelState: this.context.excel,
      userPreferences: this.context.preferences,
      conversationSummary: this.getConversationSummary()
    };
  }

  /**
   * Update user preferences
   */
  updatePreferences(newPreferences: Partial<ConversationContext['preferences']>): void {
    this.context.preferences = {
      ...this.context.preferences,
      ...newPreferences
    };
    console.log('⚙️ Updated user preferences:', newPreferences);
  }

  /**
   * Check if conversation needs attention (long gaps, errors, etc.)
   */
  needsAttention(): {
    needsAttention: boolean;
    reasons: string[];
    suggestions: string[];
  } {
    const reasons: string[] = [];
    const suggestions: string[] = [];
    
    const lastMessage = this.messages[this.messages.length - 1];
    
    // Check for long gaps
    if (lastMessage && Date.now() - lastMessage.timestamp.getTime() > 5 * 60 * 1000) {
      reasons.push('Long gap since last message');
      suggestions.push('Check if user needs help or context refresh');
    }

    // Check for repeated errors
    const recentErrors = this.messages
      .slice(-10)
      .filter(msg => msg.content.includes('error') || msg.content.includes('failed'))
      .length;
    
    if (recentErrors > 2) {
      reasons.push('Multiple recent errors');
      suggestions.push('Offer to help troubleshoot or clarify requirements');
    }

    // Check for confusion patterns
    const confusionWords = ['confused', 'don\'t understand', 'what do you mean', 'unclear'];
    const recentConfusion = this.messages
      .slice(-5)
      .filter(msg => confusionWords.some(word => msg.content.toLowerCase().includes(word)))
      .length;
      
    if (recentConfusion > 0) {
      reasons.push('User expressing confusion');
      suggestions.push('Provide clearer explanations or examples');
    }

    return {
      needsAttention: reasons.length > 0,
      reasons,
      suggestions
    };
  }

  /**
   * Generate unique session ID
   */
  private generateSessionId(): string {
    return `session_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * Generate unique message ID
   */
  private generateMessageId(): string {
    return `msg_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * Get conversation statistics
   */
  getStatistics(): {
    messagesPerHour: number;
    averageMessageLength: number;
    mostActiveHour: number;
    excelOperationsPerHour: number;
  } {
    const duration = Date.now() - this.startTime.getTime();
    const hours = duration / (1000 * 60 * 60);
    
    const messagesPerHour = this.messages.length / Math.max(hours, 0.01);
    
    const totalLength = this.messages.reduce((sum, msg) => sum + msg.content.length, 0);
    const averageMessageLength = totalLength / Math.max(this.messages.length, 1);
    
    // Group messages by hour to find most active
    const hourCounts: { [hour: number]: number } = {};
    this.messages.forEach(msg => {
      const hour = msg.timestamp.getHours();
      hourCounts[hour] = (hourCounts[hour] || 0) + 1;
    });
    
    const mostActiveHour = Object.entries(hourCounts)
      .sort(([,a], [,b]) => b - a)[0]?.[0] || 0;
    
    const excelOperationsPerHour = this.context.excel.recentChanges.length / Math.max(hours, 0.01);

    return {
      messagesPerHour,
      averageMessageLength,
      mostActiveHour: Number(mostActiveHour),
      excelOperationsPerHour
    };
  }
}

export default ConversationManager;