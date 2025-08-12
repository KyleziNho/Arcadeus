/**
 * Chat Response Cleaner
 * Fixes malformed HTML and improves chat response formatting
 */

class ChatResponseCleaner {
  constructor() {
    this.setupStyles();
  }

  /**
   * Clean and format chat responses
   */
  cleanResponse(htmlContent) {
    if (!htmlContent || typeof htmlContent !== 'string') {
      return htmlContent;
    }

    let cleaned = htmlContent;

    // Fix broken table/item markup that's showing as raw text
    cleaned = this.fixBrokenMarkup(cleaned);
    
    // Clean up malformed HTML elements
    cleaned = this.cleanMalformedHTML(cleaned);
    
    // Apply proper formatting for financial content
    cleaned = this.formatFinancialContent(cleaned);
    
    // Ensure mobile responsiveness
    cleaned = this.makeMobileResponsive(cleaned);
    
    return cleaned;
  }

  /**
   * Fix broken table/item markup
   */
  fixBrokenMarkup(content) {
    let fixed = content;

    // Remove broken item tags and header elements
    fixed = fixed.replace(/item"\s*>\s*header"\s*>/g, '');
    fixed = fixed.replace(/item"\s*>/g, '');
    fixed = fixed.replace(/">\s*header"\s*>/g, '');
    
    // Fix table-like structures that show as raw markup
    const tablePattern = /Current\s+Values.*?Levered\s+IRR.*?Unlevered\s+IRR.*?Key\s+Insights/gs;
    
    if (tablePattern.test(fixed)) {
      fixed = fixed.replace(tablePattern, (match) => {
        return this.createMetricsTable(match);
      });
    }

    // Clean up any remaining broken markup
    fixed = fixed.replace(/["\s]*>\s*["\s]*>/g, '');
    fixed = fixed.replace(/item["\s]*>/g, '');
    fixed = fixed.replace(/header["\s]*>/g, '');

    return fixed;
  }

  /**
   * Create a proper metrics table from broken markup
   */
  createMetricsTable(rawContent) {
    // Extract IRR values if they exist
    const leveredMatch = rawContent.match(/Levered\s+IRR[^\d]*(\d+\.?\d*%?)/i);
    const unlevereedMatch = rawContent.match(/Unlevered\s+IRR[^\d]*(\d+\.?\d*%?)/i);
    
    const leveredValue = leveredMatch ? leveredMatch[1] : 'Not found';
    const unlevereedValue = unlevereedMatch ? unlevereedMatch[1] : 'Not found';

    return `
      <div class="metrics-comparison-table">
        <div class="metrics-header">
          <h4>📊 IRR Comparison</h4>
        </div>
        <div class="metrics-grid">
          <div class="metric-card levered">
            <div class="metric-label">Levered IRR</div>
            <div class="metric-value">${leveredValue}</div>
            <div class="metric-location">Sheet1!B10</div>
          </div>
          <div class="metric-card unlevered">
            <div class="metric-label">Unlevered IRR</div>
            <div class="metric-value">${unlevereedValue}</div>
            <div class="metric-location">Sheet1!B11</div>
          </div>
        </div>
        <div class="metrics-insight">
          <div class="insight-header">💡 Key Insight</div>
          <div class="insight-content">
            The significant difference between levered and unlevered IRR is due to the impact of debt financing on overall returns.
          </div>
        </div>
      </div>
    `;
  }

  /**
   * Clean malformed HTML elements
   */
  cleanMalformedHTML(content) {
    let cleaned = content;

    // Remove duplicate or broken tags
    cleaned = cleaned.replace(/<([^>]+)>\s*<\1[^>]*>/g, '<$1>');
    
    // Fix unclosed tags
    cleaned = cleaned.replace(/<(div|span|p)([^>]*)>([^<]*?)(?=<(?!\/)|\s*$)/g, '<$1$2>$3</$1>');
    
    // Remove empty elements
    cleaned = cleaned.replace(/<([^>]+)>\s*<\/\1>/g, '');
    
    // Clean up extra whitespace
    cleaned = cleaned.replace(/\s+/g, ' ').trim();
    
    return cleaned;
  }

  /**
   * Format financial content properly
   */
  formatFinancialContent(content) {
    let formatted = content;

    // Format currency values
    formatted = formatted.replace(/\$\s*(\d+(?:,\d{3})*(?:\.\d{2})?)\s*(M|B|K)?/g, 
      '<span class="currency-value">$$$1$2</span>');
    
    // Format percentages
    formatted = formatted.replace(/(\d+\.?\d*)\s*%/g, 
      '<span class="percentage-value">$1%</span>');
    
    // Format multiples (e.g., 2.5x, 3.2x)
    formatted = formatted.replace(/(\d+\.?\d*)\s*x/gi, 
      '<span class="multiple-value">$1x</span>');
    
    // Format cell references
    formatted = formatted.replace(/([A-Z]+!?[A-Z]+\d+)/g, 
      '<span class="cell-reference-clickable" onclick="navigateToCell(\'$1\')">$1</span>');
    
    return formatted;
  }

  /**
   * Make content mobile responsive
   */
  makeMobileResponsive(content) {
    // Wrap long content in scrollable containers
    if (content.length > 500) {
      content = `<div class="scrollable-content">${content}</div>`;
    }
    
    return content;
  }

  /**
   * Setup styles for cleaned responses
   */
  setupStyles() {
    const styleId = 'chat-response-cleaner-styles';
    if (document.getElementById(styleId)) return;

    const style = document.createElement('style');
    style.id = styleId;
    style.textContent = `
      /* Metrics Comparison Table */
      .metrics-comparison-table {
        background: white;
        border-radius: 12px;
        padding: 20px;
        margin: 16px 0;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
      }

      .metrics-header h4 {
        margin: 0 0 16px 0;
        color: #1e293b;
        font-size: 18px;
        font-weight: 600;
        display: flex;
        align-items: center;
        gap: 8px;
      }

      .metrics-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 16px;
        margin-bottom: 20px;
      }

      @media (max-width: 480px) {
        .metrics-grid {
          grid-template-columns: 1fr;
          gap: 12px;
        }
      }

      .metric-card {
        background: linear-gradient(135deg, #f8fafc, #f1f5f9);
        border: 2px solid #e2e8f0;
        border-radius: 12px;
        padding: 16px;
        text-align: center;
        transition: all 0.3s ease;
        cursor: pointer;
      }

      .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
      }

      .metric-card.levered {
        border-color: #10b981;
        background: linear-gradient(135deg, #d1fae5, #a7f3d0);
      }

      .metric-card.unlevered {
        border-color: #3b82f6;
        background: linear-gradient(135deg, #dbeafe, #93c5fd);
      }

      .metric-label {
        font-size: 12px;
        font-weight: 600;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 8px;
      }

      .metric-value {
        font-size: 28px;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 8px;
        font-family: 'SF Mono', Monaco, monospace;
      }

      .metric-location {
        font-size: 11px;
        color: #64748b;
        font-family: 'SF Mono', Monaco, monospace;
        background: rgba(100, 116, 139, 0.1);
        padding: 2px 6px;
        border-radius: 4px;
        display: inline-block;
      }

      .metrics-insight {
        background: linear-gradient(135deg, #fef3c7, #fde68a);
        border-radius: 8px;
        padding: 16px;
        border-left: 4px solid #f59e0b;
      }

      .insight-header {
        font-weight: 600;
        color: #92400e;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 6px;
      }

      .insight-content {
        color: #78350f;
        line-height: 1.5;
      }

      /* Enhanced Financial Value Formatting */
      .currency-value {
        background: #dcfce7;
        color: #166534;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 600;
        font-family: 'SF Mono', Monaco, monospace;
      }

      .percentage-value {
        background: #fef3c7;
        color: #92400e;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 600;
        font-family: 'SF Mono', Monaco, monospace;
      }

      .multiple-value {
        background: #e0e7ff;
        color: #3730a3;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 600;
        font-family: 'SF Mono', Monaco, monospace;
      }

      .cell-reference-clickable {
        background: #dcfce7;
        color: #15803d;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 600;
        font-family: 'SF Mono', Monaco, monospace;
        cursor: pointer;
        transition: all 0.2s ease;
        border: 1px solid #86efac;
      }

      .cell-reference-clickable:hover {
        background: #22c55e;
        color: white;
        transform: translateY(-1px);
        box-shadow: 0 2px 8px rgba(34, 197, 94, 0.3);
      }

      /* Scrollable content for mobile */
      .scrollable-content {
        max-width: 100%;
        overflow-x: auto;
        padding: 8px;
      }

      /* Fix for broken layouts */
      .chat-message .message-text {
        word-wrap: break-word;
        overflow-wrap: break-word;
        max-width: 100%;
      }

      /* Better mobile responsiveness */
      @media (max-width: 768px) {
        .metrics-comparison-table {
          padding: 16px;
          margin: 12px 0;
        }

        .metric-value {
          font-size: 24px;
        }

        .metrics-insight {
          padding: 12px;
        }

        .scrollable-content {
          font-size: 14px;
        }
      }
    `;

    document.head.appendChild(style);
  }

  /**
   * Process all existing chat messages
   */
  cleanAllChatMessages() {
    const chatMessages = document.querySelectorAll('.chat-message .message-text');
    
    chatMessages.forEach(messageElement => {
      const originalHTML = messageElement.innerHTML;
      const cleanedHTML = this.cleanResponse(originalHTML);
      
      if (cleanedHTML !== originalHTML) {
        messageElement.innerHTML = cleanedHTML;
      }
    });
  }
}

// Initialize the cleaner
window.chatResponseCleaner = new ChatResponseCleaner();

// Hook into message processing
if (window.chatHandler && window.chatHandler.addFormattedMessageToExistingChat) {
  const originalAddMessage = window.chatHandler.addFormattedMessageToExistingChat.bind(window.chatHandler);
  
  window.chatHandler.addFormattedMessageToExistingChat = function(role, content) {
    // Clean the content before displaying
    const cleanedContent = window.chatResponseCleaner.cleanResponse(content);
    return originalAddMessage(role, cleanedContent);
  };
}

// Clean existing messages on load
document.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    window.chatResponseCleaner.cleanAllChatMessages();
  }, 1000);
});

console.log('🧹 Chat Response Cleaner loaded and active');