/**
 * SamplingHandler.ts
 * Handles sampling requests from MCP servers, allowing them to request AI completions
 * with human-in-the-loop approval
 */

import {
  MCPContent,
  ConversationMessage
} from '../mcp-types/interfaces';

export interface SamplingRequest {
  messages: Array<{
    role: 'user' | 'assistant' | 'system';
    content: string;
  }>;
  modelPreferences?: {
    hints?: Array<{ name: string }>;
    costPriority?: number;
    speedPriority?: number;
    intelligencePriority?: number;
  };
  systemPrompt?: string;
  maxTokens?: number;
  includeContext?: 'none' | 'thisServer' | 'allServers';
  metadata?: Record<string, any>;
}

export interface SamplingApproval {
  approved: boolean;
  modifiedRequest?: SamplingRequest;
  reason?: string;
}

export class SamplingHandler {
  private pendingRequests: Map<string, SamplingRequest> = new Map();
  private approvalCallback?: (request: SamplingRequest) => Promise<SamplingApproval>;
  private apiEndpoint: string;

  constructor() {
    this.setupAPIEndpoint();
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
   * Set the approval callback for human-in-the-loop control
   */
  setApprovalCallback(callback: (request: SamplingRequest) => Promise<SamplingApproval>): void {
    this.approvalCallback = callback;
  }

  /**
   * Handle a sampling request from a server
   */
  async handleSamplingRequest(
    serverId: string,
    request: SamplingRequest,
    context?: ConversationMessage[]
  ): Promise<MCPContent[]> {
    console.log(`🤖 Sampling request from server ${serverId}:`, request);

    try {
      // Step 1: Get user approval (human-in-the-loop)
      const approval = await this.requestUserApproval(serverId, request);
      
      if (!approval.approved) {
        return [{
          type: 'error',
          text: `Sampling request denied: ${approval.reason || 'User declined'}`
        }];
      }

      // Use modified request if user made changes
      const finalRequest = approval.modifiedRequest || request;

      // Step 2: Include context if requested
      const messagesWithContext = this.buildMessagesWithContext(
        finalRequest,
        context,
        finalRequest.includeContext
      );

      // Step 3: Call AI service
      const response = await this.callAIService(messagesWithContext, finalRequest);

      // Step 4: Get user approval for response (second human-in-the-loop)
      const responseApproval = await this.requestResponseApproval(serverId, response);
      
      if (!responseApproval.approved) {
        return [{
          type: 'error',
          text: `Response rejected: ${responseApproval.reason || 'User declined response'}`
        }];
      }

      // Return approved response
      return [{
        type: 'text',
        text: responseApproval.modifiedResponse || response
      }];

    } catch (error: any) {
      console.error('❌ Sampling request failed:', error);
      return [{
        type: 'error',
        text: `Sampling failed: ${error.message}`
      }];
    }
  }

  /**
   * Request user approval for sampling
   */
  private async requestUserApproval(
    serverId: string,
    request: SamplingRequest
  ): Promise<SamplingApproval> {
    if (!this.approvalCallback) {
      // Auto-approve if no callback set (for testing only)
      console.warn('⚠️ No approval callback set, auto-approving sampling request');
      return { approved: true };
    }

    // Show approval UI through callback
    const approval = await this.approvalCallback(request);
    
    // Log the decision
    console.log(`✅ Sampling request ${approval.approved ? 'approved' : 'denied'} by user`);
    
    return approval;
  }

  /**
   * Request user approval for AI response
   */
  private async requestResponseApproval(
    serverId: string,
    response: string
  ): Promise<{ approved: boolean; modifiedResponse?: string; reason?: string }> {
    // In production, this would show UI for response approval
    // For now, auto-approve
    return { approved: true, modifiedResponse: response };
  }

  /**
   * Build messages array with context if requested
   */
  private buildMessagesWithContext(
    request: SamplingRequest,
    context?: ConversationMessage[],
    includeContext?: string
  ): Array<{ role: string; content: string }> {
    const messages = [...request.messages];

    if (includeContext && includeContext !== 'none' && context) {
      // Add relevant context messages
      const contextMessages = context
        .slice(-5) // Last 5 messages for context
        .map(msg => ({
          role: msg.role,
          content: msg.content
        }));
      
      // Insert context before the current request
      messages.unshift(...contextMessages);
    }

    return messages;
  }

  /**
   * Call AI service with sampling request
   */
  private async callAIService(
    messages: Array<{ role: string; content: string }>,
    request: SamplingRequest
  ): Promise<string> {
    try {
      const response = await fetch(this.apiEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        body: JSON.stringify({
          message: messages[messages.length - 1].content, // Last message as main prompt
          batchType: 'chat',
          systemPrompt: request.systemPrompt || 'You are a helpful assistant.',
          temperature: this.calculateTemperature(request.modelPreferences),
          maxTokens: request.maxTokens || 2000
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
      throw error;
    }
  }

  /**
   * Calculate temperature based on model preferences
   */
  private calculateTemperature(preferences?: any): number {
    if (!preferences) return 0.7;
    
    // Higher intelligence priority = lower temperature (more focused)
    const intelligenceFactor = preferences.intelligencePriority || 0.5;
    return 1.0 - (intelligenceFactor * 0.5); // Range: 0.5 to 1.0
  }

  /**
   * Create sampling approval UI component
   */
  createApprovalUI(request: SamplingRequest): HTMLElement {
    const container = document.createElement('div');
    container.className = 'sampling-approval-modal';
    container.innerHTML = `
      <div class="sampling-approval-content">
        <h3>🤖 AI Sampling Request</h3>
        <p>A server is requesting AI assistance:</p>
        
        <div class="sampling-details">
          <div class="sampling-field">
            <label>Request:</label>
            <div class="sampling-messages">
              ${request.messages.map(m => `
                <div class="sampling-message ${m.role}">
                  <strong>${m.role}:</strong> ${this.escapeHtml(m.content)}
                </div>
              `).join('')}
            </div>
          </div>
          
          ${request.systemPrompt ? `
            <div class="sampling-field">
              <label>System Prompt:</label>
              <div>${this.escapeHtml(request.systemPrompt)}</div>
            </div>
          ` : ''}
          
          <div class="sampling-field">
            <label>Max Tokens:</label>
            <div>${request.maxTokens || 2000}</div>
          </div>
          
          ${request.modelPreferences ? `
            <div class="sampling-field">
              <label>Priorities:</label>
              <div>
                Intelligence: ${(request.modelPreferences.intelligencePriority || 0.5) * 100}%<br>
                Speed: ${(request.modelPreferences.speedPriority || 0.5) * 100}%<br>
                Cost: ${(request.modelPreferences.costPriority || 0.5) * 100}%
              </div>
            </div>
          ` : ''}
        </div>
        
        <div class="sampling-actions">
          <button class="approve-btn" onclick="window.approveSampling()">✅ Approve</button>
          <button class="modify-btn" onclick="window.modifySampling()">✏️ Modify</button>
          <button class="deny-btn" onclick="window.denySampling()">❌ Deny</button>
        </div>
      </div>
    `;
    
    return container;
  }

  /**
   * Escape HTML for safe display
   */
  private escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }
}

export default SamplingHandler;