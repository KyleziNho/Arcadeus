/**
 * Response Formatter Fix - Final Solution
 * This will definitely fix the markdown formatting issue
 */

class ResponseFormatterFix {
  constructor() {
    this.initializeStyles();
    this.hookIntoExistingChat();
    console.log('🎨 Response Formatter Fix initialized');
  }

  /**
   * Initialize the CSS styles for formatting
   */
  initializeStyles() {
    const styleId = 'response-formatter-styles';
    if (document.getElementById(styleId)) return;

    const style = document.createElement('style');
    style.id = styleId;
    style.textContent = `
      /* Professional Response Formatting Styles */
      .formatted-response {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        line-height: 1.6;
        color: #2d3748;
      }

      .section-header {
        font-size: 16px;
        font-weight: 600;
        color: #1a202c;
        margin: 16px 0 8px 0;
        padding-bottom: 4px;
        border-bottom: 2px solid #e2e8f0;
      }

      .value-highlight {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 14px;
      }

      .cell-highlight {
        background: #f0f9ff;
        color: #0369a1;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 600;
        font-family: 'Courier New', monospace;
        border: 1px solid #bae6fd;
        cursor: pointer;
        transition: all 0.2s ease;
      }

      .cell-highlight:hover {
        background: #0369a1;
        color: white;
        transform: translateY(-1px);
        box-shadow: 0 2px 8px rgba(3, 105, 161, 0.3);
      }

      .money-highlight {
        background: #f0fdf4;
        color: #166534;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 600;
        border-left: 3px solid #22c55e;
      }

      .percentage-highlight {
        background: #fef3c7;
        color: #92400e;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 600;
        border-left: 3px solid #f59e0b;
      }

      .insight-item {
        margin: 12px 0;
        padding: 12px;
        background: #f8fafc;
        border-left: 4px solid #667eea;
        border-radius: 0 8px 8px 0;
      }

      .insight-label {
        font-weight: 600;
        color: #1e293b;
        display: block;
        margin-bottom: 4px;
      }

      .bullet-item {
        margin: 6px 0;
        padding-left: 16px;
        position: relative;
        color: #4b5563;
      }

      .bullet-item:before {
        content: '•';
        position: absolute;
        left: 0;
        color: #667eea;
        font-weight: bold;
      }

      .recommendation {
        background: #f0f9ff;
        border: 1px solid #bae6fd;
        border-radius: 8px;
        padding: 16px;
        margin: 16px 0;
      }

      .recommendation-title {
        font-weight: 600;
        color: #0369a1;
        margin-bottom: 8px;
      }
    `;
    
    document.head.appendChild(style);
  }

  /**
   * Hook into the existing chat system to format responses
   */
  hookIntoExistingChat() {
    // Override the ChatHandler's addFormattedChatMessage method
    if (window.chatHandler && window.chatHandler.addFormattedChatMessage) {
      const originalMethod = window.chatHandler.addFormattedChatMessage.bind(window.chatHandler);
      
      window.chatHandler.addFormattedChatMessage = (role, content) => {
        if (role === 'assistant') {
          // Apply our professional formatting
          const professionalContent = this.formatResponse(content);
          originalMethod(role, professionalContent);
        } else {
          originalMethod(role, content);
        }
      };
      
      console.log('✅ Hooked into ChatHandler.addFormattedChatMessage');
    }

    // Also hook into addChatMessage as a fallback
    if (window.chatHandler && window.chatHandler.addChatMessage) {
      const originalAddMessage = window.chatHandler.addChatMessage.bind(window.chatHandler);
      
      window.chatHandler.addChatMessage = (role, content) => {
        if (role === 'assistant') {
          const professionalContent = this.formatResponse(content);
          originalAddMessage(role, professionalContent);
        } else {
          originalAddMessage(role, content);
        }
      };
      
      console.log('✅ Hooked into ChatHandler.addChatMessage');
    }
  }

  /**
   * Format response with professional styling
   */
  formatResponse(content) {
    if (!content || typeof content !== 'string') return content;

    console.log('🎨 Formatting response:', content.substring(0, 100) + '...');

    let formatted = content
      // Remove LaTeX completely
      .replace(/\\text\{([^}]+)\}/g, '$1')
      .replace(/\\times/g, '×')
      .replace(/\\frac\{([^}]+)\}\{([^}]+)\}/g, '$1/$2')
      .replace(/\\\[[\s\S]*?\\\]/g, '')
      .replace(/\\\([\s\S]*?\\\)/g, '')
      
      // Convert headers to clean format
      .replace(/###\s*([^#\n]+)/g, '<div class="section-header">$1</div>')
      .replace(/##\s*([^#\n]+)/g, '<div class="section-header">$1</div>')
      
      // Convert **bold** to professional highlights
      .replace(/\*\*([^*\n]+)\*\*/g, '<span class="value-highlight">$1</span>')
      
      // Make cell references clickable and highlighted
      .replace(/\b([A-Z]+![A-Z]+\d+(?::[A-Z]+\d+)?)\b/g, '<span class="cell-highlight" onclick="navigateToExcelCell(\'$1\')">$1</span>')
      
      // Highlight financial values
      .replace(/\$[\d,]+(?:\.\d{2})?(?:\s*(?:million|Million|M|K|thousand))?/g, '<span class="money-highlight">$&</span>')
      .replace(/\b\d+(?:\.\d+)?%/g, '<span class="percentage-highlight">$&</span>')
      
      // Convert numbered insights
      .replace(/(\d+)\.\s*\*\*([^*]+)\*\*:?\s*([^\n]+)/g, '<div class="insight-item"><span class="insight-label">$2</span>$3</div>')
      
      // Convert bullet points
      .replace(/^[-•*]\s*\*\*([^*]+)\*\*:?\s*([^\n]+)/gm, '<div class="insight-item"><span class="insight-label">$1</span>$2</div>')
      .replace(/^[-•*]\s*([^\n]+)/gm, '<div class="bullet-item">$1</div>')
      
      // Convert recommendations sections
      .replace(/Recommendations?:?\s*\n((?:[-•*].*\n?)*)/gi, '<div class="recommendation"><div class="recommendation-title">💡 Recommendations</div>$1</div>')
      
      // Clean up excessive whitespace
      .replace(/\n{3,}/g, '\n\n')
      .replace(/\n/g, '<br>')
      .trim();

    console.log('✅ Response formatted successfully');
    return formatted;
  }

  /**
   * Apply formatting to existing messages (for testing)
   */
  reformatExistingMessages() {
    const messages = document.querySelectorAll('.assistant-message .message-text');
    console.log(`🔄 Reformatting ${messages.length} existing messages`);
    
    messages.forEach(messageElement => {
      if (!messageElement.classList.contains('formatted-response')) {
        const originalText = messageElement.textContent || messageElement.innerText;
        const formattedHTML = this.formatResponse(originalText);
        messageElement.innerHTML = formattedHTML;
        messageElement.classList.add('formatted-response');
      }
    });
  }

  /**
   * Force format the last message (for immediate testing)
   */
  formatLastMessage() {
    const messages = document.querySelectorAll('.assistant-message .message-text');
    const lastMessage = messages[messages.length - 1];
    
    if (lastMessage) {
      console.log('🎨 Formatting last message...');
      const originalText = lastMessage.textContent || lastMessage.innerText;
      const formattedHTML = this.formatResponse(originalText);
      lastMessage.innerHTML = formattedHTML;
      lastMessage.classList.add('formatted-response');
      console.log('✅ Last message formatted');
    } else {
      console.log('❌ No messages found to format');
    }
  }
}

// Create global instance
window.responseFormatterFix = new ResponseFormatterFix();

// Global function for cell navigation
window.navigateToExcelCell = function(cellReference) {
  console.log('🎯 Navigating to cell:', cellReference);
  
  if (window.excelNavigator && window.excelNavigator.navigateToCell) {
    window.excelNavigator.navigateToCell(cellReference)
      .then(() => console.log('✅ Navigation successful'))
      .catch(error => console.error('❌ Navigation failed:', error));
  } else {
    console.warn('⚠️ Excel navigator not available');
  }
};

console.log('🎨 Response Formatter Fix loaded - formatting issues should be resolved!');
console.log('💡 To test: responseFormatterFix.formatLastMessage()');
console.log('💡 To reformat all: responseFormatterFix.reformatExistingMessages()');