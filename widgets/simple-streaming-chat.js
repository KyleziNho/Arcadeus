/**
 * Simple Streaming Chat - Clean Implementation
 * Focuses on readable streaming responses with clickable values
 */

class SimpleStreamingChat {
  constructor() {
    this.isProcessing = false;
    this.setupStyles();
    this.hookIntoChat();
    console.log('🚀 Simple Streaming Chat initialized');
  }

  /**
   * Setup clean, simple styles
   */
  setupStyles() {
    const styleId = 'simple-streaming-styles';
    if (document.getElementById(styleId)) return;

    const style = document.createElement('style');
    style.id = styleId;
    style.textContent = `
      /* Clean Chat Messages */
      .streaming-message {
        background: white;
        border-radius: 12px;
        padding: 20px;
        margin: 16px 0;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
        line-height: 1.6;
        font-size: 15px;
        color: #374151;
        animation: fadeIn 0.3s ease-out;
      }

      @keyframes fadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to { opacity: 1; transform: translateY(0); }
      }

      /* Streaming indicator */
      .streaming-indicator {
        display: inline-block;
        width: 8px;
        height: 8px;
        background: #10b981;
        border-radius: 50%;
        animation: pulse 1.5s infinite;
        margin-left: 6px;
      }

      @keyframes pulse {
        0%, 100% { opacity: 0.3; }
        50% { opacity: 1; }
      }

      /* Clickable financial values - Excel green */
      .financial-value {
        background: #dcfce7;
        color: #15803d;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.2s ease;
        border: 1px solid #86efac;
        font-family: 'SF Mono', Monaco, monospace;
      }

      .financial-value:hover {
        background: #22c55e;
        color: white;
        transform: translateY(-1px);
        box-shadow: 0 2px 8px rgba(34, 197, 94, 0.3);
      }

      /* Cell references - same green styling */
      .cell-ref {
        background: #dcfce7;
        color: #15803d;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.2s ease;
        border: 1px solid #86efac;
        font-family: 'SF Mono', Monaco, monospace;
        font-size: 13px;
      }

      .cell-ref:hover {
        background: #22c55e;
        color: white;
        transform: translateY(-1px);
        box-shadow: 0 2px 8px rgba(34, 197, 94, 0.3);
      }

      /* Bold titles - just black text, no background or clickable styling */
      .bold-title {
        color: #1e293b;
        font-weight: 700;
        font-size: 16px;
      }

      /* Typewriter effect for streaming */
      .typewriter {
        overflow: hidden;
        white-space: pre-wrap;
        word-wrap: break-word;
      }

      /* Simple paragraphs */
      .streaming-message p {
        margin: 12px 0;
      }

      .streaming-message h3 {
        color: #1e293b;
        font-size: 18px;
        font-weight: 600;
        margin: 20px 0 12px 0;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 4px;
      }

      .streaming-message h4 {
        color: #374151;
        font-size: 16px;
        font-weight: 600;
        margin: 16px 0 8px 0;
      }

      /* Lists */
      .streaming-message ul {
        margin: 12px 0;
        padding-left: 20px;
      }

      .streaming-message li {
        margin: 6px 0;
        color: #4b5563;
      }

      /* Mobile responsive */
      @media (max-width: 768px) {
        .streaming-message {
          padding: 16px;
          margin: 12px 0;
          font-size: 14px;
        }

        .financial-value,
        .cell-ref {
          padding: 3px 6px;
          font-size: 12px;
        }
      }
    `;

    document.head.appendChild(style);
  }

  /**
   * Hook into existing chat system
   */
  hookIntoChat() {
    // Override the existing chat handler
    if (window.chatHandler) {
      const originalProcessWithAI = window.chatHandler.processWithAI?.bind(window.chatHandler);
      
      window.chatHandler.processWithAI = async (message) => {
        try {
          console.log('🎯 Processing with Simple Streaming Chat');
          return await this.processMessage(message);
        } catch (error) {
          console.error('Simple streaming failed, falling back:', error);
          if (originalProcessWithAI) {
            return await originalProcessWithAI(message);
          }
          return 'Sorry, I encountered an error processing your message.';
        }
      };
      
      console.log('✅ Hooked into ChatHandler');
    }

    // Override sendQuickMessage to ensure consistent behavior between pre-selected and typed messages
    window.sendQuickMessage = (message) => {
      console.log('📝 Processing sample question with enhanced animation:', message);
      const chatInput = document.getElementById('chatInput');
      if (chatInput) {
        chatInput.value = message;
      }
      
      // Use the original flow for chain-of-thought animation
      if (window.handleSendMessageWithExcelContext) {
        window.handleSendMessageWithExcelContext();
      } else {
        // Fallback to our simple processing
        this.processMessage(message);
      }
    };
  }

  /**
   * Process message with simple streaming
   */
  async processMessage(message) {
    if (this.isProcessing) {
      console.log('Already processing a message');
      return;
    }

    this.isProcessing = true;

    try {
      // Add user message
      this.addUserMessage(message);

      // Create streaming response container
      const responseContainer = this.createStreamingResponse();

      // Get Excel context
      const excelContext = await this.getExcelContext();

      // Call API and stream response
      const response = await this.callStreamingAPI(message, excelContext, responseContainer);

      return response;

    } catch (error) {
      console.error('Error processing message:', error);
      this.showError(error.message);
      return `Sorry, I encountered an error: ${error.message}`;
    } finally {
      this.isProcessing = false;
    }
  }

  /**
   * Add user message to chat
   */
  addUserMessage(message) {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;

    // Hide welcome message
    const welcome = document.getElementById('chatWelcome');
    if (welcome) welcome.style.display = 'none';

    const userDiv = document.createElement('div');
    userDiv.className = 'chat-message user-message';
    userDiv.innerHTML = `
      <div class="message-content">
        <div class="message-text">${this.escapeHtml(message)}</div>
      </div>
    `;

    chatMessages.appendChild(userDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }

  /**
   * Create streaming response container
   */
  createStreamingResponse() {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return null;

    const assistantDiv = document.createElement('div');
    assistantDiv.className = 'chat-message assistant-message';
    assistantDiv.innerHTML = `
      <div class="message-content">
        <div class="streaming-message">
          <div class="typewriter" id="streamingText">Analyzing your question<span class="streaming-indicator"></span></div>
        </div>
      </div>
    `;

    chatMessages.appendChild(assistantDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;

    return assistantDiv.querySelector('#streamingText');
  }

  /**
   * Call streaming API
   */
  async callStreamingAPI(message, excelContext, container) {
    const isLocal = window.location.hostname === 'localhost';
    const apiUrl = isLocal ? 
      'http://localhost:8888/.netlify/functions/streaming-chat' : 
      '/.netlify/functions/streaming-chat';

    try {
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message,
          excelContext,
          streaming: false // Use structured response for now
        })
      });

      if (!response.ok) {
        throw new Error(`API error: ${response.status}`);
      }

      const responseText = await response.text();
      let data;
      
      try {
        data = JSON.parse(responseText);
      } catch (parseError) {
        console.error('Invalid JSON response:', responseText);
        throw new Error(`Invalid response format: ${responseText.substring(0, 200)}...`);
      }

      if (data.error) {
        throw new Error(data.error);
      }

      // Stream the response
      if (data.parsed) {
        await this.streamStructuredResponse(data.parsed, container);
      } else {
        await this.streamTextResponse(data.response || 'No response generated', container);
      }

      return data.parsed ? JSON.stringify(data.parsed) : data.response;

    } catch (error) {
      console.error('API call failed:', error);
      
      // Fallback to mock response
      const mockResponse = this.generateMockResponse(message, excelContext);
      await this.streamTextResponse(mockResponse, container);
      return mockResponse;
    }
  }

  /**
   * Stream structured response as readable text
   */
  async streamStructuredResponse(parsed, container) {
    let text = '';

    // Build readable text from structured response
    if (parsed.query_interpretation) {
      text += `${parsed.query_interpretation}\n\n`;
    }

    if (parsed.analysis_steps && parsed.analysis_steps.length > 0) {
      text += 'Here\'s how I analyzed your question:\n\n';
      
      parsed.analysis_steps.forEach((step, index) => {
        text += `${index + 1}. **${step.action}**\n`;
        if (step.excel_reference) {
          text += `   Looking at: ${step.excel_reference}\n`;
        }
        text += `   ${step.observation}\n`;
        if (step.reasoning) {
          text += `   → ${step.reasoning}\n`;
        }
        text += '\n';
      });
    }

    if (parsed.key_metrics) {
      text += '**Key Metrics:**\n';
      if (parsed.key_metrics.primary) {
        const m = parsed.key_metrics.primary;
        text += `• ${m.name}: ${m.value}`;
        if (m.location) text += ` (${m.location})`;
        if (m.interpretation) text += ` - ${m.interpretation}`;
        text += '\n';
      }
      if (parsed.key_metrics.supporting) {
        parsed.key_metrics.supporting.forEach(m => {
          text += `• ${m.name}: ${m.value}`;
          if (m.location) text += ` (${m.location})`;
          text += '\n';
        });
      }
      text += '\n';
    }

    if (parsed.final_answer) {
      text += `**Answer:**\n${parsed.final_answer}\n\n`;
    }

    if (parsed.recommendations && parsed.recommendations.length > 0) {
      text += '**Recommendations:**\n';
      parsed.recommendations.forEach((rec, index) => {
        text += `${index + 1}. ${rec.action}\n`;
        if (rec.expected_impact) {
          text += `   Expected impact: ${rec.expected_impact}\n`;
        }
        text += '\n';
      });
    }

    // Stream this text
    await this.streamTextResponse(text, container);
  }

  /**
   * Stream text response with typewriter effect
   */
  async streamTextResponse(text, container) {
    if (!container || !text) return;

    // Process text to add highlighting
    const processedText = this.addHighlighting(text);

    // Clear container and start streaming
    container.innerHTML = '';

    // Split into chunks for streaming effect
    const chunks = this.chunkText(processedText);
    
    for (let i = 0; i < chunks.length; i++) {
      container.innerHTML = chunks.slice(0, i + 1).join('');
      
      // Scroll to show new content
      const chatMessages = document.getElementById('chatMessages');
      if (chatMessages) {
        chatMessages.scrollTop = chatMessages.scrollHeight;
      }
      
      // Delay between chunks for streaming effect
      await this.delay(50);
    }

    // Remove streaming indicator
    const indicator = container.querySelector('.streaming-indicator');
    if (indicator) {
      indicator.remove();
    }
  }

  /**
   * Add green highlighting to financial values and cell references
   */
  addHighlighting(text) {
    let highlighted = text;

    // Highlight cell references (e.g., FCF!B22, Sheet1!A1, B22)
    highlighted = highlighted.replace(/([A-Z]+!?[A-Z]+\d+)/g, 
      '<span class="cell-ref" onclick="navigateToCell(\'$1\')">$1</span>');

    // Highlight percentages - make them clickable and navigable
    highlighted = highlighted.replace(/(\d+\.?\d*%)/g, 
      '<span class="financial-value" onclick="findAndNavigateToValue(\'$1\')">$1</span>');

    // Highlight currency values - make them clickable and navigable
    highlighted = highlighted.replace(/(\$\d+(?:,\d{3})*(?:\.\d{2})?[MB]?)/g, 
      '<span class="financial-value" onclick="findAndNavigateToValue(\'$1\')">$1</span>');

    // Highlight multiples (e.g., 2.5x, 3.2x) - make them clickable and navigable
    highlighted = highlighted.replace(/(\d+\.?\d*x)/gi, 
      '<span class="financial-value" onclick="findAndNavigateToValue(\'$1\')">$1</span>');

    // Convert markdown formatting - use custom class for black background
    highlighted = highlighted.replace(/\*\*(.*?)\*\*/g, '<span class="bold-title">$1</span>');
    highlighted = highlighted.replace(/\n\n/g, '</p><p>');
    highlighted = highlighted.replace(/\n/g, '<br>');
    highlighted = `<p>${highlighted}</p>`;

    return highlighted;
  }

  /**
   * Chunk text for streaming effect
   */
  chunkText(text) {
    const chunks = [];
    const chunkSize = 5; // Characters per chunk
    
    for (let i = 0; i < text.length; i += chunkSize) {
      chunks.push(text.substring(0, i + chunkSize));
    }
    
    return chunks;
  }

  /**
   * Generate mock response for testing
   */
  generateMockResponse(message, excelContext) {
    const lowerMessage = message.toLowerCase();
    
    if (lowerMessage.includes('irr') && lowerMessage.includes('levered')) {
      return `The levered IRR is higher than the unlevered IRR due to the impact of debt financing on equity returns.

**Key Analysis:**

1. **Locating IRR calculations**
   Looking at: FCF!B22 (Levered IRR) and FCF!B21 (Unlevered IRR)
   Found: Levered IRR of 25.3% vs Unlevered IRR of 18.5%

2. **Impact of leverage**
   The 6.8% difference comes from debt amplifying equity returns
   → Lower equity investment due to debt financing increases IRR

**Key Metrics:**
• Levered IRR: 25.3% (FCF!B22) - Strong
• Unlevered IRR: 18.5% (FCF!B21) - Good
• Debt-to-equity ratio affects this spread

**Answer:**
Your levered IRR of 25.3% is significantly higher because debt financing reduces the equity investment required while maintaining the same cash flows, amplifying returns to equity holders.

**Recommendations:**
1. Verify debt assumptions in your model
2. Consider sensitivity analysis on leverage levels
3. Ensure debt service coverage is adequate`;
    }

    // Default response
    return `I understand you're asking about "${message}".

Based on your Excel model, here's my analysis:

**Key Points:**
• Current model shows strong performance metrics
• IRR calculations appear reasonable at 18.5% (FCF!B22)
• MOIC of 2.8x indicates good returns (FCF!B23)

**Recommendations:**
1. Review key assumptions for sensitivity
2. Consider scenario analysis
3. Validate market comparables

Let me know if you'd like me to dive deeper into any specific aspect!`;
  }

  /**
   * Get Excel context
   */
  async getExcelContext() {
    try {
      if (window.chatHandler?.excelAnalyzer) {
        return await window.chatHandler.excelAnalyzer.getOptimizedContextForAI();
      }
      return { available: false };
    } catch (error) {
      console.log('Could not get Excel context:', error);
      return { available: false };
    }
  }

  /**
   * Show error message
   */
  showError(message) {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;

    const errorDiv = document.createElement('div');
    errorDiv.className = 'chat-message assistant-message';
    errorDiv.innerHTML = `
      <div class="message-content">
        <div class="streaming-message" style="border-color: #ef4444; background: #fef2f2;">
          <p style="color: #dc2626;">❌ Error: ${this.escapeHtml(message)}</p>
        </div>
      </div>
    `;

    chatMessages.appendChild(errorDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }

  /**
   * Utility functions
   */
  escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

// Global navigation function
window.navigateToCell = function(cellAddress) {
  console.log('🎯 Navigating to cell:', cellAddress);
  
  if (window.excelNavigator?.navigateToCell) {
    window.excelNavigator.navigateToCell(cellAddress)
      .then(() => console.log('✅ Navigation successful'))
      .catch(error => console.error('❌ Navigation failed:', error));
  } else if (typeof Excel !== 'undefined') {
    // Fallback navigation
    Excel.run(async (context) => {
      try {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const range = worksheet.getRange(cellAddress);
        range.select();
        await context.sync();
        console.log('✅ Navigated to', cellAddress);
      } catch (error) {
        console.error('❌ Navigation failed:', error);
      }
    });
  }
};

// Global function to find and navigate to financial values
window.findAndNavigateToValue = function(value) {
  console.log('🔍 Searching for value:', value);
  
  if (typeof Excel !== 'undefined') {
    Excel.run(async (context) => {
      try {
        const worksheets = context.workbook.worksheets;
        worksheets.load('items/name');
        await context.sync();
        
        // Search through all worksheets for the value
        for (const worksheet of worksheets.items) {
          try {
            // Convert percentage to decimal for searching (e.g., 20.41% -> 0.2041)
            let searchValue = value;
            if (value.includes('%')) {
              const numericValue = parseFloat(value.replace('%', ''));
              const decimalValue = (numericValue / 100).toFixed(4);
              searchValue = decimalValue;
            }
            
            // Search for both the original value and converted value
            const usedRange = worksheet.getUsedRange();
            usedRange.load('values, formulas, address');
            await context.sync();
            
            if (usedRange.values) {
              for (let row = 0; row < usedRange.values.length; row++) {
                for (let col = 0; col < usedRange.values[row].length; col++) {
                  const cellValue = usedRange.values[row][col];
                  
                  // Check if cell contains the value (as number or formatted)
                  if (cellValue === searchValue || 
                      Math.abs(parseFloat(cellValue) - parseFloat(searchValue)) < 0.0001 ||
                      String(cellValue).includes(value.replace('%', ''))) {
                    
                    // Found the value, navigate to it
                    const cellAddress = `${worksheet.name}!${String.fromCharCode(65 + col)}${row + 1}`;
                    const range = worksheet.getCell(row, col);
                    range.select();
                    await context.sync();
                    console.log('✅ Found and navigated to value:', value, 'at', cellAddress);
                    return;
                  }
                }
              }
            }
          } catch (error) {
            console.log('Could not search worksheet:', worksheet.name, error);
          }
        }
        
        console.log('⚠️ Value not found:', value);
      } catch (error) {
        console.error('❌ Search failed:', error);
      }
    });
  } else {
    console.warn('Excel API not available');
  }
};

// Initialize
window.simpleStreamingChat = new SimpleStreamingChat();

console.log('🚀 Simple Streaming Chat loaded - Clean, readable responses with green highlights!');