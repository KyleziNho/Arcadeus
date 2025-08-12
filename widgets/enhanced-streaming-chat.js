/**
 * Enhanced Streaming Chat System with Chain-of-Thought
 * Implements OpenAI SDK streaming with progressive UI updates
 */

class EnhancedStreamingChat {
  constructor() {
    this.renderedSteps = 0;
    this.interpretationRendered = false;
    this.currentStreamingContainer = null;
    this.setupStyles();
    this.hookIntoChat();
    console.log('🚀 Enhanced Streaming Chat System initialized');
  }

  /**
   * Setup enhanced styles for streaming UI
   */
  setupStyles() {
    const styleId = 'enhanced-streaming-styles';
    if (document.getElementById(styleId)) return;

    const style = document.createElement('style');
    style.id = styleId;
    style.textContent = `
      /* Chain of Thought Streaming Styles */
      .analysis-container {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        border-radius: 16px;
        overflow: hidden;
        margin: 12px 0;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.08);
      }

      .analysis-header {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        padding: 16px 20px;
        font-weight: 600;
        display: flex;
        align-items: center;
        gap: 10px;
      }

      .analysis-header .status-indicator {
        width: 8px;
        height: 8px;
        background: #fff;
        border-radius: 50%;
        animation: pulse 2s infinite;
      }

      .analysis-header.streaming .status-indicator {
        background: #fbbf24;
        animation: pulse 1s infinite;
      }

      @keyframes pulse {
        0%, 100% { opacity: 0.4; transform: scale(0.8); }
        50% { opacity: 1; transform: scale(1.2); }
      }

      .analysis-body {
        padding: 24px;
      }

      /* Query Interpretation */
      .interpretation-section {
        background: white;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 20px;
        border-left: 4px solid #10b981;
        opacity: 0;
        animation: fadeInUp 0.5s ease-out forwards;
      }

      .interpretation-header {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 12px;
        font-weight: 600;
        color: #1e293b;
      }

      .interpretation-content {
        color: #475569;
        line-height: 1.6;
      }

      /* Analysis Steps Container */
      .analysis-steps {
        margin: 20px 0;
      }

      .steps-header {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 16px;
        font-size: 18px;
        font-weight: 600;
        color: #1e293b;
      }

      /* Individual Analysis Step */
      .analysis-step {
        display: flex;
        gap: 16px;
        margin: 16px 0;
        padding: 20px;
        background: white;
        border-radius: 12px;
        border-left: 3px solid #10b981;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
        opacity: 0;
        transform: translateX(-20px);
        transition: all 0.5s ease-out;
      }

      .analysis-step.visible {
        opacity: 1;
        transform: translateX(0);
      }

      .step-number {
        min-width: 48px;
        height: 48px;
        background: linear-gradient(135deg, #10b981, #059669);
        color: white;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: bold;
        font-size: 18px;
        flex-shrink: 0;
      }

      .step-content {
        flex: 1;
      }

      .step-header {
        display: flex;
        justify-content: space-between;
        align-items: start;
        margin-bottom: 12px;
      }

      .step-action {
        font-weight: 600;
        color: #1e293b;
        font-size: 16px;
      }

      .step-status {
        font-size: 12px;
        padding: 4px 8px;
        border-radius: 12px;
        background: #f0f9ff;
        color: #0369a1;
        display: flex;
        align-items: center;
        gap: 4px;
      }

      .step-status.streaming {
        background: #fef3c7;
        color: #d97706;
      }

      .step-status.complete {
        background: #dcfce7;
        color: #16a34a;
      }

      .pulse-dot {
        width: 6px;
        height: 6px;
        background: currentColor;
        border-radius: 50%;
        animation: pulse 1.5s infinite;
      }

      /* Step Details */
      .step-reference {
        margin: 8px 0;
        padding: 8px 12px;
        background: #f8fafc;
        border-radius: 8px;
        display: flex;
        align-items: center;
        gap: 8px;
      }

      .step-reference .label {
        color: #64748b;
        font-size: 13px;
        font-weight: 500;
      }

      .cell-reference-clickable {
        background: #dcfce7;
        color: #15803d;
        padding: 4px 10px;
        border-radius: 6px;
        font-family: 'SF Mono', Monaco, monospace;
        font-size: 13px;
        font-weight: 600;
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

      .step-observation {
        margin: 12px 0;
        padding: 12px;
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        color: #334155;
        line-height: 1.6;
      }

      .step-calculation {
        margin: 12px 0;
        padding: 12px;
        background: #f0f9ff;
        border-radius: 8px;
        border-left: 3px solid #0ea5e9;
      }

      .step-calculation code {
        font-family: 'SF Mono', Monaco, monospace;
        color: #0369a1;
        font-size: 14px;
      }

      .step-reasoning {
        margin: 12px 0;
        padding: 12px;
        background: #fafafa;
        border-radius: 8px;
        color: #475569;
        font-style: italic;
        line-height: 1.6;
      }

      /* Typewriter Effect */
      .typewriter-text {
        position: relative;
        display: inline;
      }

      .typewriter-cursor {
        display: inline-block;
        width: 2px;
        height: 1.2em;
        background: #10b981;
        margin-left: 2px;
        animation: blink 1s infinite;
      }

      @keyframes blink {
        0%, 50% { opacity: 1; }
        51%, 100% { opacity: 0; }
      }

      /* Key Metrics Section */
      .metrics-section {
        margin: 24px 0;
        padding: 20px;
        background: white;
        border-radius: 12px;
      }

      .metrics-header {
        font-size: 18px;
        font-weight: 600;
        color: #1e293b;
        margin-bottom: 16px;
        display: flex;
        align-items: center;
        gap: 8px;
      }

      .metrics-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
        gap: 16px;
      }

      .metric-card {
        background: linear-gradient(135deg, #f8fafc, #f1f5f9);
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 16px;
        transition: all 0.3s ease;
        cursor: pointer;
      }

      .metric-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.1);
        border-color: #10b981;
      }

      .metric-card.primary {
        background: linear-gradient(135deg, #dcfce7, #d1fae5);
        border-color: #86efac;
      }

      .metric-name {
        font-size: 12px;
        color: #64748b;
        font-weight: 600;
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
        color: #0369a1;
        font-family: 'SF Mono', Monaco, monospace;
        background: #f0f9ff;
        padding: 2px 6px;
        border-radius: 4px;
        display: inline-block;
      }

      .metric-interpretation {
        display: inline-block;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 600;
        margin-left: 8px;
      }

      .metric-interpretation.excellent,
      .metric-interpretation.strong {
        background: #dcfce7;
        color: #166534;
      }

      .metric-interpretation.good {
        background: #fef3c7;
        color: #92400e;
      }

      .metric-interpretation.fair,
      .metric-interpretation.poor {
        background: #fee2e2;
        color: #991b1b;
      }

      /* Final Answer Section */
      .final-answer-section {
        margin: 24px 0;
        padding: 24px;
        background: linear-gradient(135deg, #f0fdf4, #ecfdf5);
        border-radius: 12px;
        border: 2px solid #86efac;
      }

      .final-answer-header {
        font-size: 20px;
        font-weight: 700;
        color: #166534;
        margin-bottom: 16px;
        display: flex;
        align-items: center;
        gap: 8px;
      }

      .final-answer-content {
        color: #1e293b;
        line-height: 1.8;
        font-size: 16px;
      }

      /* Recommendations */
      .recommendations-section {
        margin: 24px 0;
      }

      .recommendation-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px;
        margin-bottom: 12px;
        border-left: 4px solid transparent;
        transition: all 0.3s ease;
      }

      .recommendation-card.high {
        border-left-color: #dc2626;
      }

      .recommendation-card.medium {
        border-left-color: #f59e0b;
      }

      .recommendation-card.low {
        border-left-color: #10b981;
      }

      .recommendation-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 8px;
      }

      .recommendation-action {
        font-weight: 600;
        color: #1e293b;
      }

      .recommendation-priority {
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
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

      .recommendation-impact {
        color: #64748b;
        font-size: 14px;
        margin: 8px 0;
      }

      .recommendation-cells {
        display: flex;
        gap: 6px;
        flex-wrap: wrap;
        margin-top: 8px;
      }

      .recommendation-cell {
        font-family: 'SF Mono', Monaco, monospace;
        font-size: 11px;
        background: #f8fafc;
        padding: 2px 6px;
        border-radius: 4px;
        color: #0369a1;
        cursor: pointer;
      }

      .recommendation-cell:hover {
        background: #0ea5e9;
        color: white;
      }

      /* Animations */
      @keyframes fadeInUp {
        from {
          opacity: 0;
          transform: translateY(20px);
        }
        to {
          opacity: 1;
          transform: translateY(0);
        }
      }

      @keyframes slideIn {
        from {
          opacity: 0;
          transform: translateX(-20px);
        }
        to {
          opacity: 1;
          transform: translateX(0);
        }
      }

      /* Loading State */
      .step-skeleton {
        background: linear-gradient(90deg, #f0f0f0 25%, transparent 37%, #f0f0f0 63%);
        background-size: 400% 100%;
        animation: skeleton 1.5s ease-in-out infinite;
        border-radius: 8px;
        height: 20px;
        margin: 8px 0;
      }

      @keyframes skeleton {
        0% { background-position: 100% 0; }
        100% { background-position: -100% 0; }
      }
    `;

    document.head.appendChild(style);
  }

  /**
   * Hook into existing chat system
   */
  hookIntoChat() {
    if (window.chatHandler) {
      const originalProcessWithAI = window.chatHandler.processWithAI?.bind(window.chatHandler);
      
      if (originalProcessWithAI) {
        window.chatHandler.processWithAI = async (message) => {
          try {
            console.log('🚀 Processing with Enhanced Streaming System');
            return await this.processWithStreaming(message);
          } catch (error) {
            console.error('Enhanced streaming failed, falling back:', error);
            return await originalProcessWithAI(message);
          }
        };
        
        console.log('✅ Hooked into ChatHandler for enhanced streaming');
      }
    }
  }

  /**
   * Main processing method with streaming
   */
  async processWithStreaming(message) {
    console.log('🎯 Starting enhanced streaming for:', message);
    
    // Reset state
    this.renderedSteps = 0;
    this.interpretationRendered = false;
    
    // Get Excel context
    let excelContext = null;
    if (window.chatHandler?.excelAnalyzer) {
      try {
        excelContext = await window.chatHandler.excelAnalyzer.getOptimizedContextForAI();
      } catch (error) {
        console.log('Could not get Excel context:', error);
      }
    }
    
    // Create streaming UI container
    const streamingContainer = this.createStreamingContainer(message);
    this.currentStreamingContainer = streamingContainer;
    
    try {
      // Call the new streaming API
      const response = await this.callStreamingAPI(message, excelContext, streamingContainer);
      
      // Finalize the streaming UI
      this.finalizeStreamingUI(streamingContainer);
      
      return response;
      
    } catch (error) {
      this.showStreamingError(streamingContainer, error);
      throw error;
    }
  }

  /**
   * Call the streaming API with structured output
   */
  async callStreamingAPI(message, excelContext, container) {
    console.log('📡 Calling streaming API...');
    
    const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
    const apiUrl = isLocal ? 'http://localhost:8888/.netlify/functions/streaming-chat' : '/.netlify/functions/streaming-chat';
    
    try {
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message,
          excelContext,
          streaming: false // Netlify Functions don't support true streaming
        })
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      
      if (data.error) {
        throw new Error(data.error);
      }

      if (data.parsed) {
        // Simulate streaming by progressively rendering the structured response
        await this.simulateStreaming(data.parsed, container, data.queryType);
        return JSON.stringify(data.parsed);
      }

      return data.response || 'No response generated';
      
    } catch (error) {
      console.error('Streaming API error:', error);
      throw error;
    }
  }

  /**
   * Simulate streaming by progressively rendering structured response
   */
  async simulateStreaming(parsedResponse, container, queryType) {
    const body = container.querySelector('.analysis-body');
    
    // Step 1: Render interpretation
    if (parsedResponse.query_interpretation) {
      await this.renderInterpretation(parsedResponse.query_interpretation, body);
      await this.delay(500);
    }
    
    // Step 2: Render analysis steps
    if (parsedResponse.analysis_steps) {
      await this.renderAnalysisSteps(parsedResponse.analysis_steps, body);
    } else if (parsedResponse.structure_analysis) {
      await this.renderAnalysisSteps(parsedResponse.structure_analysis, body);
    }
    
    // Step 3: Render key metrics
    if (parsedResponse.key_metrics) {
      await this.renderKeyMetrics(parsedResponse.key_metrics, body);
      await this.delay(400);
    }
    
    // Step 4: Render insights
    if (parsedResponse.insights) {
      await this.renderInsights(parsedResponse.insights, body);
    }
    
    // Step 5: Render final answer
    if (parsedResponse.final_answer) {
      await this.renderFinalAnswer(parsedResponse.final_answer, body);
      await this.delay(300);
    }
    
    // Step 6: Render recommendations
    if (parsedResponse.recommendations) {
      await this.renderRecommendations(parsedResponse.recommendations, body);
    }
    
    // Step 7: Render next steps
    if (parsedResponse.next_steps) {
      await this.renderNextSteps(parsedResponse.next_steps, body);
    }
  }

  /**
   * Create streaming container
   */
  createStreamingContainer(message) {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return null;

    // Remove welcome message
    const welcomeMsg = document.getElementById('chatWelcome');
    if (welcomeMsg) welcomeMsg.style.display = 'none';

    // Create container
    const container = document.createElement('div');
    container.className = 'chat-message assistant-message';
    container.innerHTML = `
      <div class="analysis-container">
        <div class="analysis-header streaming">
          <span class="status-indicator"></span>
          <span>🎯 Analyzing your M&A model...</span>
          <span style="margin-left: auto; font-size: 12px; opacity: 0.8;">Processing</span>
        </div>
        <div class="analysis-body">
          <div class="step-skeleton"></div>
          <div class="step-skeleton" style="width: 80%;"></div>
          <div class="step-skeleton" style="width: 60%;"></div>
        </div>
      </div>
    `;

    chatMessages.appendChild(container);
    chatMessages.scrollTop = chatMessages.scrollHeight;
    
    return container;
  }

  /**
   * Render interpretation section
   */
  async renderInterpretation(interpretation, body) {
    // Clear skeleton loaders
    body.querySelectorAll('.step-skeleton').forEach(el => el.remove());
    
    const section = document.createElement('div');
    section.className = 'interpretation-section';
    section.innerHTML = `
      <div class="interpretation-header">
        <span>🎯</span>
        <span>Understanding your question</span>
      </div>
      <div class="interpretation-content">
        ${this.typewriterHTML(interpretation)}
      </div>
    `;
    
    body.appendChild(section);
    await this.delay(300);
  }

  /**
   * Render analysis steps progressively
   */
  async renderAnalysisSteps(steps, body) {
    const stepsContainer = document.createElement('div');
    stepsContainer.className = 'analysis-steps';
    stepsContainer.innerHTML = `
      <div class="steps-header">
        <span>🔍</span>
        <span>Walking through the analysis...</span>
      </div>
    `;
    body.appendChild(stepsContainer);
    
    for (const step of steps) {
      await this.renderSingleStep(step, stepsContainer);
      await this.delay(600);
    }
  }

  /**
   * Render a single analysis step
   */
  async renderSingleStep(step, container) {
    const stepElement = document.createElement('div');
    stepElement.className = 'analysis-step';
    
    const stepNumber = step.step_number || this.renderedSteps + 1;
    const excelRef = step.excel_reference || step.location || '';
    const observation = step.observation || step.explanation || '';
    
    stepElement.innerHTML = `
      <div class="step-number">${stepNumber}</div>
      <div class="step-content">
        <div class="step-header">
          <div class="step-action">${step.action}</div>
          <div class="step-status streaming">
            <span class="pulse-dot"></span>
            Analyzing...
          </div>
        </div>
        ${excelRef ? `
          <div class="step-reference">
            <span class="label">Looking at:</span>
            <span class="cell-reference-clickable" onclick="navigateToCell('${excelRef}')">
              ${excelRef}
            </span>
          </div>
        ` : ''}
        <div class="step-observation">
          ${this.typewriterHTML(observation)}
        </div>
        ${step.calculation ? `
          <div class="step-calculation">
            <code>${step.calculation}</code>
          </div>
        ` : ''}
        ${step.reasoning ? `
          <div class="step-reasoning">
            ${this.typewriterHTML(step.reasoning)}
          </div>
        ` : ''}
      </div>
    `;
    
    container.appendChild(stepElement);
    
    // Animate in
    requestAnimationFrame(() => {
      stepElement.classList.add('visible');
      
      // Update status after animation
      setTimeout(() => {
        const statusEl = stepElement.querySelector('.step-status');
        if (statusEl) {
          statusEl.className = 'step-status complete';
          statusEl.innerHTML = '✓ Complete';
        }
      }, 800);
    });
    
    this.renderedSteps++;
    
    // Scroll to show new step
    const chatMessages = document.getElementById('chatMessages');
    if (chatMessages) {
      chatMessages.scrollTop = chatMessages.scrollHeight;
    }
  }

  /**
   * Render key metrics
   */
  async renderKeyMetrics(metrics, body) {
    const section = document.createElement('div');
    section.className = 'metrics-section';
    section.innerHTML = `
      <div class="metrics-header">
        <span>📊</span>
        <span>Key Metrics</span>
      </div>
      <div class="metrics-grid"></div>
    `;
    
    body.appendChild(section);
    const grid = section.querySelector('.metrics-grid');
    
    // Render primary metric
    if (metrics.primary) {
      const card = this.createMetricCard(metrics.primary, true);
      grid.appendChild(card);
      await this.delay(300);
    }
    
    // Render supporting metrics
    if (metrics.supporting) {
      for (const metric of metrics.supporting) {
        const card = this.createMetricCard(metric, false);
        grid.appendChild(card);
        await this.delay(200);
      }
    }
  }

  /**
   * Create metric card element
   */
  createMetricCard(metric, isPrimary) {
    const card = document.createElement('div');
    card.className = `metric-card ${isPrimary ? 'primary' : ''}`;
    card.onclick = () => {
      if (metric.location) {
        window.navigateToCell?.(metric.location);
      }
    };
    
    card.innerHTML = `
      <div class="metric-name">${metric.name}</div>
      <div class="metric-value">${metric.value}</div>
      <div>
        <span class="metric-location">${metric.location}</span>
        ${metric.interpretation ? `
          <span class="metric-interpretation ${metric.interpretation.toLowerCase()}">
            ${metric.interpretation}
          </span>
        ` : ''}
      </div>
      ${metric.formula ? `
        <div style="margin-top: 8px; font-size: 11px; color: #64748b;">
          Formula: <code>${metric.formula}</code>
        </div>
      ` : ''}
    `;
    
    return card;
  }

  /**
   * Render insights
   */
  async renderInsights(insights, body) {
    const section = document.createElement('div');
    section.className = 'insights-section';
    section.innerHTML = `
      <div class="section-header">
        <span>💡</span>
        <span>Key Insights</span>
      </div>
    `;
    
    body.appendChild(section);
    
    for (const insight of insights) {
      const card = document.createElement('div');
      card.className = `insight-card ${insight.type}`;
      card.innerHTML = `
        <div class="insight-title">${insight.title}</div>
        <div class="insight-content">${insight.content}</div>
        <div class="insight-impact ${insight.impact}">${insight.impact} impact</div>
      `;
      section.appendChild(card);
      await this.delay(300);
    }
  }

  /**
   * Render final answer
   */
  async renderFinalAnswer(answer, body) {
    const section = document.createElement('div');
    section.className = 'final-answer-section';
    section.innerHTML = `
      <div class="final-answer-header">
        <span>✅</span>
        <span>Answer</span>
      </div>
      <div class="final-answer-content">
        ${this.typewriterHTML(answer)}
      </div>
    `;
    
    body.appendChild(section);
    
    // Scroll to show answer
    const chatMessages = document.getElementById('chatMessages');
    if (chatMessages) {
      chatMessages.scrollTop = chatMessages.scrollHeight;
    }
  }

  /**
   * Render recommendations
   */
  async renderRecommendations(recommendations, body) {
    const section = document.createElement('div');
    section.className = 'recommendations-section';
    section.innerHTML = `
      <div class="section-header">
        <span>🎯</span>
        <span>Recommendations</span>
      </div>
    `;
    
    body.appendChild(section);
    
    for (const rec of recommendations) {
      const card = document.createElement('div');
      card.className = `recommendation-card ${rec.priority}`;
      card.innerHTML = `
        <div class="recommendation-header">
          <div class="recommendation-action">${rec.action}</div>
          <div class="recommendation-priority ${rec.priority}">${rec.priority} priority</div>
        </div>
        <div class="recommendation-impact">${rec.expected_impact}</div>
        ${rec.cells_to_modify && rec.cells_to_modify.length > 0 ? `
          <div class="recommendation-cells">
            ${rec.cells_to_modify.map(cell => `
              <span class="recommendation-cell" onclick="navigateToCell('${cell}')">
                ${cell}
              </span>
            `).join('')}
          </div>
        ` : ''}
      `;
      section.appendChild(card);
      await this.delay(300);
    }
  }

  /**
   * Render next steps
   */
  async renderNextSteps(steps, body) {
    const section = document.createElement('div');
    section.className = 'next-steps-section';
    section.innerHTML = `
      <div class="section-header">
        <span>🚀</span>
        <span>Next Steps</span>
      </div>
      <ul class="next-steps-list">
        ${steps.map(step => `<li>${step}</li>`).join('')}
      </ul>
    `;
    
    body.appendChild(section);
  }

  /**
   * Finalize streaming UI
   */
  finalizeStreamingUI(container) {
    if (!container) return;
    
    const header = container.querySelector('.analysis-header');
    if (header) {
      header.classList.remove('streaming');
      const statusText = header.querySelector('span:last-child');
      if (statusText) {
        statusText.textContent = 'Complete';
      }
    }
  }

  /**
   * Show streaming error
   */
  showStreamingError(container, error) {
    if (!container) return;
    
    const body = container.querySelector('.analysis-body');
    if (body) {
      body.innerHTML = `
        <div class="error-section">
          <div class="section-header">
            <span>⚠️</span>
            <span>Error</span>
          </div>
          <div style="background: #fee2e2; color: #dc2626; padding: 16px; border-radius: 8px;">
            Failed to complete analysis: ${error.message}
          </div>
        </div>
      `;
    }
  }

  /**
   * Create typewriter effect HTML
   */
  typewriterHTML(text) {
    return `<span class="typewriter-text">${text}</span>`;
  }

  /**
   * Utility delay function
   */
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

// Initialize the enhanced streaming system
window.enhancedStreamingChat = new EnhancedStreamingChat();

// Global navigation function
window.navigateToCell = function(cellReference) {
  console.log('🎯 Navigating to cell:', cellReference);
  
  if (window.excelNavigator?.navigateToCell) {
    window.excelNavigator.navigateToCell(cellReference)
      .then(() => console.log('✅ Navigation successful'))
      .catch(error => console.error('❌ Navigation failed:', error));
  } else {
    console.warn('⚠️ Excel navigator not available');
  }
};

console.log('🚀 Enhanced Streaming Chat System loaded!');
console.log('💡 Features: Chain-of-thought reasoning, progressive UI updates, structured outputs');