/**
 * Direct Response Formatter - Aggressively format ugly responses
 * This is a last resort to ensure formatting works
 */

class DirectResponseFormatter {
  constructor() {
    this.pollingInterval = null;
    this.lastProcessedCount = 0;
  }

  start() {
    console.log('🚨 Starting aggressive response formatting...');
    
    // Inject CSS first
    this.injectCSS();
    
    // Start polling for new messages every 500ms
    this.pollingInterval = setInterval(() => {
      this.scanAndFormatMessages();
    }, 500);
    
    console.log('✅ Aggressive formatter started');
  }

  stop() {
    if (this.pollingInterval) {
      clearInterval(this.pollingInterval);
      this.pollingInterval = null;
    }
  }

  scanAndFormatMessages() {
    // Look for any message elements that contain ugly formatting
    const messageElements = document.querySelectorAll(
      '.message-text, .chat-content, [class*="message"]'
    );

    let processed = 0;
    messageElements.forEach(el => {
      const text = el.textContent || el.innerHTML;
      
      // Check if this looks like an ugly AI response
      if (this.looksLikeUglyResponse(text) && !el.classList.contains('already-formatted')) {
        console.log('🎯 Found ugly response, formatting...', el);
        this.formatElement(el);
        processed++;
      }
    });

    if (processed > this.lastProcessedCount) {
      console.log(`✅ Formatted ${processed - this.lastProcessedCount} new messages`);
      this.lastProcessedCount = processed;
    }
  }

  looksLikeUglyResponse(text) {
    if (!text || text.length < 50) return false;
    
    // Check for ugly formatting patterns
    const uglyPatterns = [
      /\*\*[^*]+\*\*/,  // **bold**
      /\\\[[\s\S]*?\\\]/, // LaTeX
      /### /,           // Markdown headers
      /FCF![A-Z]\d+/,   // Cell references
      /\$[\d,]+\.\d{2}/, // Money amounts
    ];

    return uglyPatterns.some(pattern => pattern.test(text));
  }

  formatElement(element) {
    try {
      const originalText = element.textContent || element.innerHTML;
      const formattedHTML = this.enhanceResponseFormatting(originalText);
      
      element.innerHTML = formattedHTML;
      element.classList.add('already-formatted', 'formatted-response');
      
      console.log('✅ Successfully formatted element');
    } catch (error) {
      console.error('❌ Error formatting element:', error);
    }
  }

  enhanceResponseFormatting(content) {
    if (!content || typeof content !== 'string') return content;

    let formatted = content
      // Remove markdown headers
      .replace(/### ([^#\n]+)/g, '<div class="section-header">$1</div>')
      .replace(/## ([^#\n]+)/g, '<div class="section-header">$1</div>')
      
      // Convert **bold** to highlighted values
      .replace(/\*\*([^*\n]{1,50})\*\*/g, '<span class="value-highlight">$1</span>')
      
      // Convert cell references to clickable highlighted ranges with hover tooltips
      .replace(/\b([A-Z]+![A-Z]+\d+(?::[A-Z]+\d+)?)\b/g, '<span class="cell-highlight clickable-cell" data-cell="$1" onclick="navigateToExcelCell(\'$1\')" onmouseenter="showCellPreview(\'$1\', this)" onmouseleave="hideCellPreview(this)">$1</span>')
      
      // Convert financial figures
      .replace(/\$(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)\b/g, '<span class="money-highlight">$$$1</span>')
      .replace(/\b(\d+(?:\.\d+)?x)\b/g, '<span class="percentage-highlight">$1</span>')
      .replace(/\b(\d+(?:\.\d+)?%)\b/g, '<span class="percentage-highlight">$1</span>')
      
      // Remove LaTeX
      .replace(/\\\[[\s\S]*?\\\]/g, '')
      .replace(/\\\([\s\S]*?\\\)/g, '')
      
      // Convert numbered lists
      .replace(/(\d+)\.\s+\*\*([^*]+)\*\*:?\s*/g, '<div class="insight-item"><strong>$1.</strong> $2</div>')
      
      // Clean up
      .replace(/\n{3,}/g, '\n\n')
      .trim();

    // Convert line breaks to paragraphs
    const paragraphs = formatted.split('\n\n').filter(p => p.trim());
    if (paragraphs.length > 1) {
      formatted = paragraphs.map(p => `<p>${p.replace(/\n/g, '<br>')}</p>`).join('');
    } else {
      formatted = formatted.replace(/\n/g, '<br>');
    }

    return formatted;
  }

  injectCSS() {
    const existingStyle = document.getElementById('direct-response-formatting');
    if (existingStyle) return;

    const style = document.createElement('style');
    style.id = 'direct-response-formatting';
    
    style.textContent = `
      .formatted-response {
        line-height: 1.6 !important;
        font-size: 15px !important;
      }

      .formatted-response p {
        margin: 10px 0 !important;
      }

      .formatted-response .value-highlight {
        background: #dbeafe !important;
        color: #1e40af !important;
        padding: 2px 6px !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
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
        display: inline-block !important;
        margin: 0 2px !important;
      }

      .formatted-response .percentage-highlight {
        background: #fce7f3 !important;
        color: #be185d !important;
        padding: 2px 6px !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        display: inline-block !important;
        margin: 0 2px !important;
      }

      .formatted-response .section-header {
        font-weight: 700 !important;
        font-size: 16px !important;
        color: #111827 !important;
        margin: 16px 0 8px 0 !important;
        padding-bottom: 4px !important;
        border-bottom: 2px solid #e5e7eb !important;
      }

      .formatted-response .insight-item {
        margin: 8px 0 !important;
        padding: 10px !important;
        background: #f9fafb !important;
        border-radius: 8px !important;
        border-left: 3px solid #10b981 !important;
      }
    `;

    document.head.appendChild(style);
    console.log('✅ Direct formatting CSS injected');
  }
}

// Start the aggressive formatter immediately
let directFormatter = null;

function startDirectFormatter() {
  if (directFormatter) {
    directFormatter.stop();
  }
  
  directFormatter = new DirectResponseFormatter();
  directFormatter.start();
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', startDirectFormatter);
} else {
  setTimeout(startDirectFormatter, 100);
}

// Also start when chat becomes visible
setTimeout(() => {
  const chatTab = document.getElementById('chatTab');
  if (chatTab) {
    chatTab.addEventListener('click', () => {
      setTimeout(startDirectFormatter, 500);
    });
  }
}, 1000);

console.log('💪 Direct Response Formatter loaded and ready');

window.DirectResponseFormatter = DirectResponseFormatter;