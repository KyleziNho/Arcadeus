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

/**
 * Direct Chat Message Formatter - More reliable approach
 */
class DirectChatFormatter {
  constructor() {
    this.observer = null;
    this.isFormatting = false;
  }

  /**
   * Start watching for new chat messages and format them
   */
  startWatching() {
    console.log('🔍 Starting direct chat message formatting...');
    console.log('Current URL:', window.location.href);
    console.log('Document ready state:', document.readyState);

    // Inject CSS
    this.injectCSS();

    // Watch for new messages being added to the chat
    const chatContainer = document.getElementById('chatMessages');
    if (!chatContainer) {
      console.error('❌ Chat container #chatMessages not found');
      console.log('Available elements with "chat" in ID:');
      document.querySelectorAll('[id*="chat"], [class*="chat"]').forEach(el => {
        console.log('- ', el.tagName, el.id || 'no-id', el.className || 'no-class');
      });
      return;
    }

    console.log('✅ Found chat container:', chatContainer);

    // Create MutationObserver to watch for new messages
    this.observer = new MutationObserver((mutations) => {
      console.log('🔄 MutationObserver detected changes:', mutations.length);
      
      mutations.forEach((mutation) => {
        console.log('Mutation type:', mutation.type, 'Added nodes:', mutation.addedNodes.length);
        
        mutation.addedNodes.forEach((node) => {
          if (node.nodeType === Node.ELEMENT_NODE) {
            console.log('Added element:', node.tagName, node.className);
            
            // Look for assistant messages with more flexible matching
            const isAssistantMessage = node.classList && (
              node.classList.contains('assistant-message') || 
              (node.classList.contains('chat-message-modern') && node.classList.contains('assistant-message')) ||
              node.classList.contains('chat-message')
            );
            
            if (isAssistantMessage) {
              console.log('🎯 Found assistant message element!');
              this.formatAssistantMessage(node);
            }
            
            // Also check child nodes for assistant messages
            const assistantMessages = node.querySelectorAll(
              '.assistant-message, .chat-message-modern.assistant-message, .chat-message.assistant-message'
            );
            console.log('Found child assistant messages:', assistantMessages.length);
            assistantMessages.forEach((msg) => this.formatAssistantMessage(msg));
          }
        });
      });
    });

    // Start observing
    this.observer.observe(chatContainer, {
      childList: true,
      subtree: true
    });

    console.log('✅ Chat message formatting started');
  }

  /**
   * Format an assistant message element
   */
  formatAssistantMessage(messageElement) {
    if (this.isFormatting) return; // Prevent recursive formatting
    
    console.log('💅 Formatting assistant message...');
    this.isFormatting = true;

    try {
      // Find the message text element
      const textElement = messageElement.querySelector('.message-text');
      if (!textElement) {
        console.log('No .message-text found in message');
        return;
      }

      // Get the raw text content
      const rawText = textElement.textContent || textElement.innerHTML;
      console.log('Raw text length:', rawText.length);

      // Apply enhanced formatting
      const formattedHTML = this.enhanceResponseFormatting(rawText);
      console.log('Formatted HTML length:', formattedHTML.length);

      // Replace the content with formatted HTML
      textElement.innerHTML = formattedHTML;
      textElement.classList.add('formatted-response');

      console.log('✅ Message formatted successfully');

    } catch (error) {
      console.error('❌ Error formatting message:', error);
    } finally {
      this.isFormatting = false;
    }
  }

  /**
   * Enhanced response formatting
   */
  enhanceResponseFormatting(content) {
    if (!content || typeof content !== 'string') return content;

    console.log('🎨 Applying enhanced formatting...');

    let formatted = content
      // Remove markdown headers and make them clean
      .replace(/### ([^#\n]+)/g, '<div class="section-header">$1</div>')
      .replace(/## ([^#\n]+)/g, '<div class="section-header">$1</div>')
      
      // Convert **bold** to highlighted values (but be more specific)
      .replace(/\*\*([^*\n]+)\*\*/g, '<span class="value-highlight">$1</span>')
      
      // Convert cell references to highlighted ranges
      .replace(/\b([A-Z]+![A-Z]+\d+(?::[A-Z]+\d+)?)\b/g, '<span class="cell-highlight">$1</span>')
      
      // Convert financial figures to highlighted values
      .replace(/\$(\d{1,3}(?:,\d{3})*(?:\.\d{2})?(?:\s?million|\s?M|\s?k|\s?K)?)\b/g, '<span class="money-highlight">$$$1</span>')
      .replace(/\b(\d+(?:\.\d+)?x)\b/g, '<span class="percentage-highlight">$1</span>')
      .replace(/\b(\d+(?:\.\d+)?%)\b/g, '<span class="percentage-highlight">$1</span>')
      
      // Clean up LaTeX and markdown formulas
      .replace(/\\\[[\s\S]*?\\\]/g, '')
      .replace(/\\\([\s\S]*?\\\)/g, '')
      .replace(/\\\\ /g, ' ')
      
      // Convert numbered lists to better format
      .replace(/(\d+)\.\s+\*\*([^*]+)\*\*:?\s*/g, '<div class="insight-item"><span class="insight-number">$1.</span><span class="insight-label">$2</span></div>')
      
      // Convert bullet points
      .replace(/[-•]\s+\*\*([^*]+)\*\*:?\s*/g, '<div class="bullet-item"><span class="bullet-label">$1</span></div>')
      .replace(/[-•]\s+([^\n]+)/g, '<div class="bullet-item">$1</div>')
      
      // Clean up multiple spaces and line breaks
      .replace(/\n{3,}/g, '\n\n')
      .replace(/\s{3,}/g, ' ')
      .trim();

    // Convert line breaks to HTML
    formatted = formatted.replace(/\n\n/g, '</p><p>').replace(/\n/g, '<br>');
    if (formatted && !formatted.startsWith('<p>')) {
      formatted = '<p>' + formatted + '</p>';
    }

    return formatted;
  }

  /**
   * Inject CSS styles
   */
  injectCSS() {
    // Remove existing styles
    const existing = document.getElementById('direct-chat-formatting');
    if (existing) existing.remove();

    const style = document.createElement('style');
    style.id = 'direct-chat-formatting';
    
    style.textContent = `
      /* Direct Chat Message Formatting */
      .formatted-response {
        line-height: 1.6 !important;
        font-size: 15px !important;
      }

      .formatted-response p {
        margin: 8px 0 !important;
        padding: 0 !important;
      }

      /* Value highlighting */
      .formatted-response .value-highlight {
        background: #dbeafe !important;
        color: #1e40af !important;
        padding: 2px 6px !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        display: inline-block !important;
        margin: 0 2px !important;
      }

      .formatted-response .cell-highlight {
        background: #d1fae5 !important;
        color: #065f46 !important;
        padding: 2px 6px !important;
        border-radius: 6px !important;
        font-family: 'Monaco', 'Consolas', monospace !important;
        font-size: 13px !important;
        font-weight: 600 !important;
        display: inline-block !important;
        margin: 0 2px !important;
      }

      .formatted-response .money-highlight {
        background: #fef3c7 !important;
        color: #92400e !important;
        padding: 2px 6px !important;
        border-radius: 6px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        display: inline-block !important;
        margin: 0 2px !important;
      }

      .formatted-response .percentage-highlight {
        background: #fce7f3 !important;
        color: #be185d !important;
        padding: 2px 6px !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        display: inline-block !important;
        margin: 0 2px !important;
      }

      /* Section headers */
      .formatted-response .section-header {
        font-weight: 700 !important;
        font-size: 16px !important;
        color: #111827 !important;
        margin: 16px 0 8px 0 !important;
        padding: 8px 0 4px 0 !important;
        border-bottom: 2px solid #e5e7eb !important;
        display: block !important;
      }

      /* Insight items */
      .formatted-response .insight-item {
        margin: 8px 0 !important;
        padding: 10px !important;
        background: #f9fafb !important;
        border-radius: 8px !important;
        border-left: 3px solid #10b981 !important;
        display: block !important;
      }

      .formatted-response .insight-number {
        font-weight: 700 !important;
        color: #059669 !important;
        margin-right: 8px !important;
      }

      .formatted-response .insight-label {
        font-weight: 600 !important;
        color: #059669 !important;
        margin-right: 8px !important;
      }

      .formatted-response .bullet-item {
        margin: 6px 0 !important;
        padding: 4px 0 4px 16px !important;
        position: relative !important;
        color: #4b5563 !important;
        display: block !important;
      }

      .formatted-response .bullet-item:before {
        content: "•" !important;
        color: #10b981 !important;
        font-weight: bold !important;
        position: absolute !important;
        left: 0 !important;
      }

      .formatted-response .bullet-label {
        font-weight: 600 !important;
        color: #059669 !important;
      }
    `;

    document.head.appendChild(style);
    console.log('✅ Direct formatting CSS injected');
  }

  /**
   * Stop watching
   */
  stopWatching() {
    if (this.observer) {
      this.observer.disconnect();
      this.observer = null;
    }
  }
}

// Initialize the direct formatter
let directFormatter = null;

function initializeDirectChatFormatter() {
  if (directFormatter) {
    directFormatter.stopWatching();
  }
  
  directFormatter = new DirectChatFormatter();
  directFormatter.startWatching();
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializeDirectChatFormatter);
} else {
  // DOM is already loaded, initialize immediately
  setTimeout(initializeDirectChatFormatter, 100);
}

// Also initialize when chat page becomes visible
function initializeWhenChatVisible() {
  const chatPage = document.getElementById('chatPage');
  if (chatPage) {
    const observer = new MutationObserver((mutations) => {
      mutations.forEach((mutation) => {
        if (mutation.type === 'attributes' && mutation.attributeName === 'style') {
          const isVisible = chatPage.style.display !== 'none';
          if (isVisible && !directFormatter) {
            setTimeout(initializeDirectChatFormatter, 100);
          }
        }
      });
    });
    
    observer.observe(chatPage, { attributes: true });
    
    // Also check if it's already visible
    if (chatPage.style.display !== 'none') {
      setTimeout(initializeDirectChatFormatter, 100);
    }
  }
}

// Start watching for chat page visibility
setTimeout(initializeWhenChatVisible, 500);

// Backup: Initialize on any chat button click
document.addEventListener('click', function(e) {
  if (e.target && (e.target.id === 'chatTab' || e.target.textContent === 'Chat')) {
    setTimeout(initializeDirectChatFormatter, 200);
  }
});

// Initialize enhanced formatting when DOM is ready - LEGACY FALLBACK
document.addEventListener('DOMContentLoaded', function() {
  const enhancedFormatter = new EnhancedFormattingInjector();
  enhancedFormatter.inject();
});

// Also initialize if DOM is already loaded - LEGACY FALLBACK  
if (document.readyState === 'loading') {
  // DOM is still loading
} else {
  // DOM is already loaded
  const enhancedFormatter = new EnhancedFormattingInjector();
  enhancedFormatter.inject();
}

// Export for global use
window.EnhancedFormattingInjector = EnhancedFormattingInjector;