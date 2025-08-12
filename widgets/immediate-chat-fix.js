/**
 * Immediate Chat Fix
 * Runs immediately to fix any existing broken chat responses
 */

(function() {
  'use strict';
  
  /**
   * Fix broken chat responses immediately
   */
  function fixExistingChatResponses() {
    console.log('🔧 Running immediate chat response fixes...');
    
    try {
      // Find all chat message elements
      const messageElements = document.querySelectorAll('.chat-message .message-text, .assistant-message .message-text, .formatted-response');
      
      let fixedCount = 0;
      
      messageElements.forEach((element, index) => {
        const originalHTML = element.innerHTML;
        
        if (originalHTML && (originalHTML.includes('item">') || 
                            originalHTML.includes('header">') || 
                            originalHTML.includes('">'))) {
          
          // Apply the same cleaning logic as the ChatResponseCleaner
          let fixed = originalHTML;
          
          // Remove broken markup
          fixed = fixed.replace(/item"\s*>\s*header"\s*>/g, '');
          fixed = fixed.replace(/item"\s*>/g, '');
          fixed = fixed.replace(/">\s*header"\s*>/g, '');
          fixed = fixed.replace(/["\s]*>\s*["\s]*>/g, '');
          fixed = fixed.replace(/item["\s]*>/g, '');
          fixed = fixed.replace(/header["\s]*>/g, '');
          
          // Create proper structure for IRR comparison content
          if (fixed.includes('Levered IRR') && fixed.includes('Unlevered IRR')) {
            const leveredMatch = fixed.match(/Levered\s+IRR[^\d]*(\d+\.?\d*%?)/i);
            const unlevereedMatch = fixed.match(/Unlevered\s+IRR[^\d]*(\d+\.?\d*%?)/i);
            
            const leveredValue = leveredMatch ? leveredMatch[1] : 'Not found';
            const unlevereedValue = unlevereedMatch ? unlevereedMatch[1] : 'Not found';
            
            fixed = `
              <div class="metrics-comparison-table">
                <div class="metrics-header">
                  <h4>📊 IRR Analysis</h4>
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
                    The significant difference between levered and unlevered IRR is due to the impact of debt financing on overall returns. Here's a detailed breakdown of the factors influencing this discrepancy:
                  </div>
                </div>
              </div>
            `;
          }
          
          // Format financial values
          fixed = fixed.replace(/\$\s*(\d+(?:,\d{3})*(?:\.\d{2})?)\s*(M|B|K)?/g, 
            '<span class="currency-value">$$1$2</span>');
          fixed = fixed.replace(/(\d+\.?\d*)\s*%/g, 
            '<span class="percentage-value">$1%</span>');
          fixed = fixed.replace(/(\d+\.?\d*)\s*x/gi, 
            '<span class="multiple-value">$1x</span>');
          
          // Apply the fix if it's different
          if (fixed !== originalHTML) {
            element.innerHTML = fixed;
            fixedCount++;
            console.log(`✅ Fixed chat message ${index + 1}`);
          }
        }
      });
      
      if (fixedCount > 0) {
        console.log(`🎉 Successfully fixed ${fixedCount} chat response(s)`);
        
        // Trigger a small visual notification
        const notification = document.createElement('div');
        notification.style.cssText = `
          position: fixed;
          top: 20px;
          right: 20px;
          background: #10b981;
          color: white;
          padding: 12px 16px;
          border-radius: 8px;
          font-size: 14px;
          font-weight: 500;
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
          z-index: 10000;
          animation: slideIn 0.3s ease-out;
        `;
        notification.textContent = `✅ Fixed ${fixedCount} chat response${fixedCount > 1 ? 's' : ''}`;
        
        document.body.appendChild(notification);
        
        // Remove notification after 3 seconds
        setTimeout(() => {
          if (notification.parentNode) {
            notification.style.opacity = '0';
            setTimeout(() => notification.remove(), 300);
          }
        }, 3000);
      } else {
        console.log('ℹ️ No broken chat responses found to fix');
      }
      
    } catch (error) {
      console.error('❌ Error fixing chat responses:', error);
    }
  }
  
  /**
   * Ensure styles are available
   */
  function ensureFixStyles() {
    const styleId = 'immediate-chat-fix-styles';
    if (document.getElementById(styleId)) return;

    const style = document.createElement('style');
    style.id = styleId;
    style.textContent = `
      @keyframes slideIn {
        from {
          transform: translateX(100%);
          opacity: 0;
        }
        to {
          transform: translateX(0);
          opacity: 1;
        }
      }
      
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
    `;

    document.head.appendChild(style);
  }
  
  // Run the fix when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      ensureFixStyles();
      setTimeout(fixExistingChatResponses, 500);
    });
  } else {
    ensureFixStyles();
    fixExistingChatResponses();
  }
  
  // Also run the fix after a delay to catch any dynamically loaded content
  setTimeout(() => {
    fixExistingChatResponses();
  }, 2000);
  
  // Expose function for manual use
  window.fixChatResponses = fixExistingChatResponses;
  
  console.log('🔧 Immediate chat fix script loaded');
  
})();