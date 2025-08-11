/**
 * Enhanced Formatting Injector
 * Injects CSS and modifies existing chat system for better formatting
 */

class EnhancedFormattingInjector {
  constructor() {
    this.injected = false;
  }

  /**
   * Inject enhanced formatting into existing chat system
   */
  inject() {
    if (this.injected) return;

    console.log('💅 Injecting enhanced chat formatting...');

    // Inject CSS styles
    this.injectCSS();

    // Override existing addMessageToChat function
    this.enhanceExistingChatSystem();

    // Add live search indicators support
    this.addLiveSearchSupport();

    this.injected = true;
    console.log('✅ Enhanced formatting injected successfully');
  }

  /**
   * Inject CSS styles for enhanced formatting
   */
  injectCSS() {
    const style = document.createElement('style');
    style.id = 'enhanced-chat-formatting';
    
    style.textContent = `
      /* Enhanced message formatting */
      .formatted-response {
        line-height: 1.6 !important;
        font-size: 15px !important;
      }

      /* Value highlighting */
      .formatted-response .value-highlight {
        background: #dbeafe;
        color: #1e40af;
        padding: 2px 6px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 14px;
      }

      .formatted-response .cell-highlight {
        background: #d1fae5;
        color: #065f46;
        padding: 2px 6px;
        border-radius: 6px;
        font-family: 'Monaco', 'Consolas', monospace;
        font-size: 13px;
        font-weight: 600;
      }

      .formatted-response .money-highlight {
        background: #fef3c7;
        color: #92400e;
        padding: 2px 6px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 14px;
      }

      .formatted-response .percentage-highlight {
        background: #fce7f3;
        color: #be185d;
        padding: 2px 6px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 14px;
      }

      /* Section headers */
      .formatted-response .section-header {
        font-weight: 700;
        font-size: 16px;
        color: #111827;
        margin: 16px 0 8px 0;
        padding-bottom: 4px;
        border-bottom: 2px solid #e5e7eb;
      }

      /* Insight items */
      .formatted-response .insight-item {
        margin: 8px 0;
        padding: 10px;
        background: #f9fafb;
        border-radius: 8px;
        border-left: 3px solid #10b981;
      }

      .formatted-response .insight-label {
        font-weight: 600;
        color: #059669;
        margin-right: 8px;
      }

      .formatted-response .bullet-item {
        margin: 6px 0;
        padding-left: 16px;
        position: relative;
        color: #4b5563;
      }

      .formatted-response .bullet-item:before {
        content: "•";
        color: #10b981;
        font-weight: bold;
        position: absolute;
        left: 0;
      }

      /* Live search indicators */
      .live-search-indicators {
        margin: 15px 0;
        padding: 15px;
        background: #f8f9fa;
        border-radius: 12px;
        border-left: 4px solid #10b981;
        animation: slideInUp 0.3s ease-out;
      }

      .search-step {
        display: flex;
        align-items: center;
        padding: 8px 0;
        opacity: 0.3;
        transition: all 0.4s ease;
      }

      .search-step.active {
        opacity: 1;
        animation: highlightStep 0.6s ease-out;
      }

      .search-icon {
        font-size: 16px;
        margin-right: 10px;
        min-width: 20px;
      }

      .search-text {
        font-weight: 500;
        color: #374151;
        margin-right: 10px;
      }

      .search-highlight {
        background: #d1fae5;
        color: #065f46;
        padding: 4px 8px;
        border-radius: 6px;
        font-family: 'Monaco', 'Consolas', monospace;
        font-size: 13px;
        font-weight: 600;
      }

      /* Animations */
      @keyframes slideInUp {
        from {
          transform: translateY(20px);
          opacity: 0;
        }
        to {
          transform: translateY(0);
          opacity: 1;
        }
      }

      @keyframes highlightStep {
        0% {
          background: transparent;
        }
        50% {
          background: #d1fae5;
        }
        100% {
          background: transparent;
        }
      }
    `;

    document.head.appendChild(style);
  }

  /**
   * Enhance existing addMessageToChat function
   */
  enhanceExistingChatSystem() {
    // Store original function
    if (window.originalAddMessageToChat) return; // Already enhanced
    
    window.originalAddMessageToChat = window.addMessageToChat;

    // Create enhanced version
    window.addMessageToChat = (type, content) => {
      if (type === 'assistant') {
        // Apply enhanced formatting to assistant messages
        const formattedContent = this.enhanceResponseFormatting(content);
        window.originalAddMessageToChat(type, formattedContent);
      } else {
        // Use original function for other message types
        window.originalAddMessageToChat(type, content);
      }
    };
  }

  /**
   * Enhanced response formatting - same as ChatHandler
   */
  enhanceResponseFormatting(content) {
    if (!content || typeof content !== 'string') return content;

    // Remove excessive markdown formatting
    let formatted = content
      // Remove markdown headers and make them simple bold text
      .replace(/### ([^#\n]+)/g, '<div class="section-header">$1</div>')
      .replace(/## ([^#\n]+)/g, '<div class="section-header">$1</div>')
      
      // Convert **bold** to highlighted values
      .replace(/\*\*([^*]+)\*\*/g, '<span class="value-highlight">$1</span>')
      
      // Convert cell references to highlighted ranges
      .replace(/([A-Z]+![A-Z]+\d+(?::[A-Z]+\d+)?)/g, '<span class="cell-highlight">$1</span>')
      
      // Convert percentages and financial figures to highlighted values
      .replace(/(\$[\d,]+(?:\.\d{2})?(?:\s?million|\s?M)?)/g, '<span class="money-highlight">$1</span>')
      .replace(/(\d+(?:\.\d+)?%)/g, '<span class="percentage-highlight">$1</span>')
      .replace(/(\d+(?:\.\d+)?x)/g, '<span class="percentage-highlight">$1</span>')
      
      // Convert numbered lists to cleaner format
      .replace(/\d+\.\s*\*\*([^*]+)\*\*:?\s*([^\n]+)/g, '<div class="insight-item"><span class="insight-label">$1</span>$2</div>')
      
      // Convert bullet points to cleaner format
      .replace(/[-•]\s*\*\*([^*]+)\*\*:?\s*([^\n]+)/g, '<div class="insight-item"><span class="insight-label">$1</span>$2</div>')
      .replace(/[-•]\s*([^\n]+)/g, '<div class="bullet-item">$1</div>')
      
      // Remove LaTeX formatting
      .replace(/\\\[[\s\S]*?\\\]/g, '')
      .replace(/\\\([\s\S]*?\\\)/g, '')
      
      // Clean up excessive line breaks
      .replace(/\n{3,}/g, '\n\n')
      .trim();

    return formatted;
  }

  /**
   * Add live search indicators support
   */
  addLiveSearchSupport() {
    // Override sendMessageToOpenAI to add live search indicators
    if (window.originalSendMessageToOpenAI) return; // Already enhanced
    
    window.originalSendMessageToOpenAI = window.sendMessageToOpenAI;

    window.sendMessageToOpenAI = async (userMessage) => {
      // Show live search indicators
      this.showLiveSearchIndicators(userMessage);
      
      try {
        // Call original function
        await window.originalSendMessageToOpenAI(userMessage);
      } finally {
        // Hide indicators when done
        this.hideLiveSearchIndicators();
      }
    };
  }

  /**
   * Show live search indicators
   */
  showLiveSearchIndicators(message) {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;

    // Create live search indicators container
    const indicatorsDiv = document.createElement('div');
    indicatorsDiv.id = 'liveSearchIndicators';
    indicatorsDiv.className = 'live-search-indicators';
    
    // Determine what the AI will search for based on query
    const searchSteps = this.generateSearchSteps(message);
    
    let html = '';
    searchSteps.forEach((step, index) => {
      html += `
        <div class="search-step" data-step="${index}">
          <span class="search-icon">${step.icon}</span>
          <span class="search-text">${step.text}</span>
          ${step.highlight ? `<span class="search-highlight">${step.highlight}</span>` : ''}
        </div>
      `;
    });
    
    indicatorsDiv.innerHTML = html;
    chatMessages.appendChild(indicatorsDiv);
    
    // Animate search steps
    this.animateSearchSteps(searchSteps.length);
  }

  /**
   * Generate search steps based on query type
   */
  generateSearchSteps(message) {
    const lowerMessage = message.toLowerCase();
    const steps = [];

    if (lowerMessage.includes('moic') || lowerMessage.includes('multiple')) {
      steps.push(
        { icon: '🔍', text: 'Search "MOIC calculation"', highlight: null },
        { icon: '📊', text: 'Looking up precedents', highlight: 'FCF!B23' },
        { icon: '👁️', text: 'Looking up values', highlight: 'FCF!B18:I19' }
      );
    } else if (lowerMessage.includes('irr')) {
      steps.push(
        { icon: '🔍', text: 'Search "IRR calculation"', highlight: null },
        { icon: '📊', text: 'Looking up precedents', highlight: 'FCF!B21:B22' },
        { icon: '👁️', text: 'Looking up cash flows', highlight: 'FCF!B19:I19' }
      );
    } else if (lowerMessage.includes('revenue')) {
      steps.push(
        { icon: '🔍', text: 'Search "total revenue"', highlight: null },
        { icon: '📊', text: 'Looking up precedents', highlight: 'Revenue!C430:T444' },
        { icon: '👁️', text: 'Looking up values', highlight: 'Revenue!C310:T325' }
      );
    } else {
      // Generic search steps
      steps.push(
        { icon: '🔍', text: 'Analyzing Excel data', highlight: null },
        { icon: '📊', text: 'Looking up relevant metrics', highlight: null },
        { icon: '👁️', text: 'Looking up values', highlight: null }
      );
    }

    return steps;
  }

  /**
   * Animate search steps progressively
   */
  animateSearchSteps(stepCount) {
    let currentStep = 0;
    const interval = setInterval(() => {
      const stepElement = document.querySelector(`[data-step="${currentStep}"]`);
      if (stepElement) {
        stepElement.classList.add('active');
      }
      
      currentStep++;
      if (currentStep >= stepCount) {
        clearInterval(interval);
      }
    }, 800); // 800ms between each step
  }

  /**
   * Hide live search indicators
   */
  hideLiveSearchIndicators() {
    const indicators = document.getElementById('liveSearchIndicators');
    if (indicators) {
      indicators.style.opacity = '0';
      setTimeout(() => {
        indicators.remove();
      }, 300);
    }
  }
}

// Initialize enhanced formatting when DOM is ready
document.addEventListener('DOMContentLoaded', function() {
  const enhancedFormatter = new EnhancedFormattingInjector();
  enhancedFormatter.inject();
});

// Also initialize if DOM is already loaded
if (document.readyState === 'loading') {
  // DOM is still loading
} else {
  // DOM is already loaded
  const enhancedFormatter = new EnhancedFormattingInjector();
  enhancedFormatter.inject();
}

// Export for global use
window.EnhancedFormattingInjector = EnhancedFormattingInjector;