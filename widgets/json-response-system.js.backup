/**
 * JSON Response System - Ultimate Formatting Solution
 * Gets structured JSON from OpenAI and renders it professionally
 */

class JsonResponseSystem {
  constructor() {
    this.initializeSystem();
    this.setupStyles();
    this.hookIntoChat();
    console.log('🎯 JSON Response System initialized');
  }

  /**
   * Initialize the JSON response system
   */
  initializeSystem() {
    // Response templates for different types of M&A queries
    this.responseTemplates = {
      financial_analysis: {
        summary: "Brief overview of findings",
        key_metrics: [
          { label: "IRR", value: "25.3%", location: "FCF!B22", interpretation: "Strong returns" }
        ],
        insights: [
          { title: "Key Finding", content: "Detailed insight", type: "positive" }
        ],
        recommendations: [
          { priority: "high", action: "Specific recommendation", rationale: "Why this matters" }
        ],
        cell_references: ["FCF!B22", "FCF!B18"]
      },
      model_validation: {
        summary: "Model quality assessment",
        score: 85,
        errors: [],
        warnings: [],
        suggestions: []
      },
      general: {
        summary: "Response summary",
        content: "Main response content",
        highlights: [],
        references: []
      }
    };
  }

  /**
   * Setup professional styles for JSON responses
   */
  setupStyles() {
    const styleId = 'json-response-styles';
    if (document.getElementById(styleId)) return;

    const style = document.createElement('style');
    style.id = styleId;
    style.textContent = `
      /* JSON Response Professional Styles */
      .json-response {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        background: #ffffff;
        border-radius: 12px;
        padding: 0;
        margin: 8px 0;
        border: 1px solid #e2e8f0;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
      }

      .response-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 16px 20px;
        font-weight: 600;
        font-size: 16px;
        display: flex;
        align-items: center;
        gap: 8px;
      }

      .response-body {
        padding: 20px;
      }

      .response-summary {
        font-size: 16px;
        color: #2d3748;
        line-height: 1.6;
        margin-bottom: 20px;
        padding: 16px;
        background: #f7fafc;
        border-radius: 8px;
        border-left: 4px solid #667eea;
      }

      .metrics-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 16px;
        margin: 20px 0;
      }

      .metric-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px;
        transition: all 0.2s ease;
        cursor: pointer;
      }

      .metric-card:hover {
        border-color: #667eea;
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.15);
      }

      .metric-label {
        font-size: 12px;
        color: #64748b;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
      }

      .metric-value {
        font-size: 24px;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 4px;
      }

      .metric-location {
        font-size: 11px;
        color: #0369a1;
        font-family: 'Courier New', monospace;
        background: #f0f9ff;
        padding: 2px 6px;
        border-radius: 4px;
        display: inline-block;
        margin-bottom: 4px;
        cursor: pointer;
        transition: all 0.2s ease;
      }

      .metric-location:hover {
        background: #0369a1;
        color: white;
        transform: translateY(-1px);
      }

      .metric-interpretation {
        font-size: 12px;
        color: #4b5563;
        font-style: italic;
      }

      .insights-section {
        margin: 24px 0;
      }

      .section-title {
        font-size: 18px;
        font-weight: 600;
        color: #1e293b;
        margin-bottom: 16px;
        display: flex;
        align-items: center;
        gap: 8px;
      }

      .insight-item {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 12px;
        border-left: 4px solid transparent;
      }

      .insight-item.positive {
        border-left-color: #10b981;
        background: #f0fdf4;
      }

      .insight-item.warning {
        border-left-color: #f59e0b;
        background: #fffbeb;
      }

      .insight-item.negative {
        border-left-color: #ef4444;
        background: #fef2f2;
      }

      .insight-title {
        font-weight: 600;
        color: #1e293b;
        margin-bottom: 6px;
        font-size: 14px;
      }

      .insight-content {
        color: #4b5563;
        line-height: 1.5;
        font-size: 14px;
      }

      .recommendations-section {
        margin: 24px 0;
      }

      .recommendation-item {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 12px;
        border-left: 4px solid transparent;
      }

      .recommendation-item.high {
        border-left-color: #ef4444;
      }

      .recommendation-item.medium {
        border-left-color: #f59e0b;
      }

      .recommendation-item.low {
        border-left-color: #10b981;
      }

      .recommendation-priority {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 8px;
      }

      .recommendation-priority.high {
        background: #fee2e2;
        color: #dc2626;
      }

      .recommendation-priority.medium {
        background: #fef3c7;
        color: #d97706;
      }

      .recommendation-priority.low {
        background: #dcfce7;
        color: #16a34a;
      }

      .recommendation-action {
        font-weight: 600;
        color: #1e293b;
        margin-bottom: 4px;
      }

      .recommendation-rationale {
        color: #64748b;
        font-size: 13px;
        line-height: 1.4;
      }

      .references-section {
        margin: 20px 0;
        padding: 16px;
        background: #f8fafc;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
      }

      .reference-links {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-top: 8px;
      }

      .cell-reference {
        background: #0369a1;
        color: white;
        padding: 6px 12px;
        border-radius: 6px;
        font-family: 'Courier New', monospace;
        font-size: 12px;
        font-weight: 500;
        cursor: pointer;
        transition: all 0.2s ease;
        text-decoration: none;
      }

      .cell-reference:hover {
        background: #075985;
        transform: translateY(-1px);
        box-shadow: 0 2px 8px rgba(3, 105, 161, 0.3);
        color: white;
        text-decoration: none;
      }

      .loading-animation {
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 40px;
        color: #64748b;
      }

      .loading-dots {
        display: flex;
        gap: 4px;
      }

      .loading-dot {
        width: 8px;
        height: 8px;
        background: #667eea;
        border-radius: 50%;
        animation: pulse 1.5s infinite ease-in-out;
      }

      .loading-dot:nth-child(2) {
        animation-delay: 0.2s;
      }

      .loading-dot:nth-child(3) {
        animation-delay: 0.4s;
      }

      @keyframes pulse {
        0%, 80%, 100% {
          transform: scale(0.8);
          opacity: 0.5;
        }
        40% {
          transform: scale(1);
          opacity: 1;
        }
      }

      /* Mobile Responsive */
      @media (max-width: 768px) {
        .metrics-grid {
          grid-template-columns: 1fr;
        }
        
        .response-body {
          padding: 16px;
        }
        
        .metric-value {
          font-size: 20px;
        }
      }
    `;
    
    document.head.appendChild(style);
  }

  /**
   * Hook into the existing chat system
   */
  hookIntoChat() {
    // Override the existing chat methods to use JSON responses
    if (window.chatHandler) {
      const originalProcessWithAI = window.chatHandler.processWithAI?.bind(window.chatHandler);
      
      if (originalProcessWithAI) {
        window.chatHandler.processWithAI = async (message) => {
          try {
            console.log('🎯 Processing with JSON response system');
            return await this.processWithJsonResponse(message);
          } catch (error) {
            console.error('JSON response failed, falling back:', error);
            return await originalProcessWithAI(message);
          }
        };
        
        console.log('✅ Hooked into ChatHandler.processWithAI for JSON responses');
      }

      // Also override the message display
      const originalAddFormattedMessage = window.chatHandler.addFormattedChatMessage?.bind(window.chatHandler);
      
      if (originalAddFormattedMessage) {
        window.chatHandler.addFormattedChatMessage = (role, content) => {
          if (role === 'assistant' && this.isJsonResponse(content)) {
            this.renderJsonResponse(content);
          } else {
            originalAddFormattedMessage(role, content);
          }
        };
        
        console.log('✅ Hooked into ChatHandler.addFormattedChatMessage for JSON rendering');
      }
    }
  }

  /**
   * Process message with JSON response system
   */
  async processWithJsonResponse(message) {
    // Show loading animation
    this.showLoadingAnimation();
    
    try {
      // Determine query type
      const queryType = this.analyzeQueryType(message);
      
      // Create JSON-structured prompt
      const jsonPrompt = this.createJsonPrompt(message, queryType);
      
      // Get Excel context if available
      let excelContext = null;
      if (window.chatHandler?.excelAnalyzer) {
        try {
          excelContext = await window.chatHandler.excelAnalyzer.getOptimizedContextForAI();
        } catch (error) {
          console.log('Could not get Excel context:', error);
        }
      }
      
      // Make API call with structured prompt
      const response = await this.callJsonAPI(jsonPrompt, excelContext);
      
      // Hide loading animation
      this.hideLoadingAnimation();
      
      return response;
      
    } catch (error) {
      this.hideLoadingAnimation();
      throw error;
    }
  }

  /**
   * Analyze query type for appropriate JSON template
   */
  analyzeQueryType(message) {
    const lowerMessage = message.toLowerCase();
    
    if (lowerMessage.includes('irr') || lowerMessage.includes('moic') || 
        lowerMessage.includes('return') || lowerMessage.includes('multiple')) {
      return 'financial_analysis';
    }
    
    if (lowerMessage.includes('validate') || lowerMessage.includes('check') || 
        lowerMessage.includes('error') || lowerMessage.includes('review')) {
      return 'model_validation';
    }
    
    return 'general';
  }

  /**
   * Create structured JSON prompt for OpenAI
   */
  createJsonPrompt(message, queryType) {
    const template = this.responseTemplates[queryType];
    
    return `You are a professional M&A analyst. Respond to the following query with a properly structured JSON response.

User Query: "${message}"

CRITICAL: You must respond with valid JSON only, using this exact structure:
${JSON.stringify(template, null, 2)}

Guidelines:
1. Replace all template values with actual analysis
2. For key_metrics, include specific Excel cell locations (e.g., "FCF!B22")
3. Use professional M&A terminology
4. Make insights actionable and specific
5. Prioritize recommendations (high/medium/low)
6. Include cell references that users can click
7. NO markdown formatting - use plain text in JSON values
8. NO LaTeX or special formatting characters

Respond with JSON only - no explanatory text before or after.`;
  }

  /**
   * Call API with JSON-structured request
   */
  async callJsonAPI(jsonPrompt, excelContext) {
    const context = {
      message: jsonPrompt,
      systemPrompt: 'You are a professional M&A financial analyst. Always respond with properly formatted JSON as requested. Never include explanatory text outside the JSON structure.',
      temperature: 0.3,
      maxTokens: 3000,
      excelContext: excelContext,
      requireJson: true
    };

    const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
    const apiUrl = isLocal ? 'http://localhost:8888/.netlify/functions/chat' : '/.netlify/functions/chat';

    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(context)
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    
    if (data.error) {
      throw new Error(data.error);
    }

    return data.response || 'No response received';
  }

  /**
   * Check if content is a JSON response
   */
  isJsonResponse(content) {
    try {
      JSON.parse(content);
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Render JSON response with professional UI
   */
  renderJsonResponse(jsonContent) {
    try {
      const data = JSON.parse(jsonContent);
      const html = this.createJsonResponseHTML(data);
      
      // Add to chat
      const chatMessages = document.getElementById('chatMessages');
      if (chatMessages) {
        const messageDiv = document.createElement('div');
        messageDiv.className = 'chat-message assistant-message';
        messageDiv.innerHTML = html;
        chatMessages.appendChild(messageDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;
      }
      
    } catch (error) {
      console.error('Failed to render JSON response:', error);
      // Fallback to regular text rendering
      if (window.chatHandler?.addChatMessage) {
        window.chatHandler.addChatMessage('assistant', jsonContent);
      }
    }
  }

  /**
   * Create professional HTML from JSON response
   */
  createJsonResponseHTML(data) {
    let html = '<div class="json-response">';
    
    // Header
    html += `
      <div class="response-header">
        <span>🎯</span>
        <span>M&A Analysis</span>
      </div>
      <div class="response-body">
    `;
    
    // Summary
    if (data.summary) {
      html += `<div class="response-summary">${data.summary}</div>`;
    }
    
    // Key Metrics
    if (data.key_metrics && data.key_metrics.length > 0) {
      html += '<div class="section-title">💰 Key Metrics</div>';
      html += '<div class="metrics-grid">';
      
      data.key_metrics.forEach(metric => {
        html += `
          <div class="metric-card" onclick="navigateToExcelCell('${metric.location}')">
            <div class="metric-label">${metric.label}</div>
            <div class="metric-value">${metric.value}</div>
            <div class="metric-location" onclick="event.stopPropagation(); navigateToExcelCell('${metric.location}')">${metric.location}</div>
            <div class="metric-interpretation">${metric.interpretation}</div>
          </div>
        `;
      });
      
      html += '</div>';
    }
    
    // Insights
    if (data.insights && data.insights.length > 0) {
      html += '<div class="insights-section">';
      html += '<div class="section-title">💡 Key Insights</div>';
      
      data.insights.forEach(insight => {
        html += `
          <div class="insight-item ${insight.type || 'neutral'}">
            <div class="insight-title">${insight.title}</div>
            <div class="insight-content">${insight.content}</div>
          </div>
        `;
      });
      
      html += '</div>';
    }
    
    // Recommendations
    if (data.recommendations && data.recommendations.length > 0) {
      html += '<div class="recommendations-section">';
      html += '<div class="section-title">🎯 Recommendations</div>';
      
      data.recommendations.forEach(rec => {
        html += `
          <div class="recommendation-item ${rec.priority}">
            <div class="recommendation-priority ${rec.priority}">${rec.priority} priority</div>
            <div class="recommendation-action">${rec.action}</div>
            <div class="recommendation-rationale">${rec.rationale}</div>
          </div>
        `;
      });
      
      html += '</div>';
    }
    
    // Cell References
    if (data.cell_references && data.cell_references.length > 0) {
      html += '<div class="references-section">';
      html += '<div class="section-title">📊 Excel References</div>';
      html += '<div class="reference-links">';
      
      data.cell_references.forEach(ref => {
        html += `<a href="#" class="cell-reference" onclick="navigateToExcelCell('${ref}'); return false;">${ref}</a>`;
      });
      
      html += '</div></div>';
    }
    
    html += '</div></div>';
    
    return html;
  }

  /**
   * Show loading animation
   */
  showLoadingAnimation() {
    const chatMessages = document.getElementById('chatMessages');
    if (chatMessages) {
      const loadingDiv = document.createElement('div');
      loadingDiv.id = 'json-loading-animation';
      loadingDiv.className = 'chat-message assistant-message';
      loadingDiv.innerHTML = `
        <div class="json-response">
          <div class="response-header">
            <span>🎯</span>
            <span>Analyzing your M&A model...</span>
          </div>
          <div class="loading-animation">
            <div class="loading-dots">
              <div class="loading-dot"></div>
              <div class="loading-dot"></div>
              <div class="loading-dot"></div>
            </div>
          </div>
        </div>
      `;
      
      chatMessages.appendChild(loadingDiv);
      chatMessages.scrollTop = chatMessages.scrollHeight;
    }
  }

  /**
   * Hide loading animation
   */
  hideLoadingAnimation() {
    const loadingDiv = document.getElementById('json-loading-animation');
    if (loadingDiv) {
      loadingDiv.remove();
    }
  }

  /**
   * Test the JSON response system
   */
  async testJsonResponse(testQuery = "What is my IRR and why is it so high?") {
    console.log('🧪 Testing JSON Response System...');
    
    try {
      const response = await this.processWithJsonResponse(testQuery);
      console.log('✅ JSON Response received:', response);
      
      if (this.isJsonResponse(response)) {
        console.log('✅ Response is valid JSON');
        this.renderJsonResponse(response);
      } else {
        console.log('⚠️ Response is not JSON, rendering as text');
      }
      
    } catch (error) {
      console.error('❌ JSON Response test failed:', error);
    }
  }
}

// Initialize the JSON Response System
window.jsonResponseSystem = new JsonResponseSystem();

// Global navigation function
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

console.log('🎯 JSON Response System loaded!');
console.log('💡 Test with: jsonResponseSystem.testJsonResponse()');
console.log('🎯 All chat responses will now use professional JSON formatting');