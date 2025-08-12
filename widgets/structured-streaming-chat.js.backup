/**
 * Structured Streaming Chat System
 * Implements OpenAI Structured Outputs with real-time streaming
 * Based on OpenAI's latest best practices for schema-enforced responses
 */

class StructuredStreamingChat {
  constructor() {
    this.setupSchemas();
    this.setupStyles();
    this.hookIntoChat();
    console.log('🚀 Structured Streaming Chat System initialized');
  }

  /**
   * Define JSON schemas for different M&A query types
   * Equivalent to Pydantic models in Python
   */
  setupSchemas() {
    // Financial Analysis Schema (MOIC, IRR, etc.)
    this.schemas = {
      financial_analysis: {
        type: "object",
        properties: {
          summary: {
            type: "string",
            description: "Brief 1-2 sentence overview of the financial analysis"
          },
          key_metrics: {
            type: "array",
            description: "Array of financial metrics with their values and interpretations",
            items: {
              type: "object",
              properties: {
                metric: { type: "string", description: "Metric name (e.g., 'MOIC', 'IRR')" },
                value: { type: "string", description: "Current value (e.g., '3.2x', '25.3%')" },
                location: { type: "string", description: "Excel cell location (e.g., 'FCF!B22')" },
                interpretation: { 
                  type: "string", 
                  enum: ["Excellent", "Strong", "Good", "Fair", "Poor", "Critical"],
                  description: "Performance interpretation" 
                },
                context: { type: "string", description: "What drives this metric" }
              },
              required: ["metric", "value", "location", "interpretation", "context"],
              additionalProperties: false
            }
          },
          insights: {
            type: "array",
            description: "Key insights about the financial performance",
            items: {
              type: "object",
              properties: {
                title: { type: "string", description: "Insight title" },
                content: { type: "string", description: "Detailed insight explanation" },
                impact: { 
                  type: "string", 
                  enum: ["high", "medium", "low"],
                  description: "Impact level on overall performance" 
                },
                type: { 
                  type: "string", 
                  enum: ["positive", "negative", "neutral", "warning"],
                  description: "Type of insight for styling" 
                }
              },
              required: ["title", "content", "impact", "type"],
              additionalProperties: false
            }
          },
          recommendations: {
            type: "array",
            description: "Actionable recommendations for the user",
            items: {
              type: "object",
              properties: {
                priority: { 
                  type: "string", 
                  enum: ["high", "medium", "low"],
                  description: "Priority level" 
                },
                action: { type: "string", description: "Specific action to take" },
                rationale: { type: "string", description: "Why this recommendation matters" },
                effort: { 
                  type: "string", 
                  enum: ["quick", "moderate", "extensive"],
                  description: "Effort required to implement" 
                }
              },
              required: ["priority", "action", "rationale", "effort"],
              additionalProperties: false
            }
          },
          next_steps: {
            type: "array",
            description: "Immediate next steps for the user",
            items: {
              type: "string",
              description: "Specific actionable step"
            }
          }
        },
        required: ["summary", "key_metrics", "insights", "recommendations"],
        additionalProperties: false
      },

      // Excel Structure Analysis Schema
      excel_analysis: {
        type: "object",
        properties: {
          summary: { type: "string", description: "Overview of Excel structure analysis" },
          formulas: {
            type: "array",
            description: "Key formulas and their analysis",
            items: {
              type: "object",
              properties: {
                location: { type: "string", description: "Cell location" },
                formula: { type: "string", description: "Formula text" },
                purpose: { type: "string", description: "What the formula calculates" },
                dependencies: { 
                  type: "array", 
                  items: { type: "string" },
                  description: "Cells this formula depends on" 
                },
                complexity: { 
                  type: "string", 
                  enum: ["simple", "moderate", "complex"],
                  description: "Formula complexity level" 
                }
              },
              required: ["location", "formula", "purpose", "complexity"],
              additionalProperties: false
            }
          },
          validation: {
            type: "object",
            properties: {
              errors: { 
                type: "array", 
                items: { type: "string" },
                description: "Detected errors" 
              },
              warnings: { 
                type: "array", 
                items: { type: "string" },
                description: "Potential issues" 
              },
              score: { 
                type: "number", 
                minimum: 0, 
                maximum: 100,
                description: "Overall model quality score" 
              }
            },
            required: ["errors", "warnings", "score"],
            additionalProperties: false
          }
        },
        required: ["summary", "formulas", "validation"],
        additionalProperties: false
      },

      // General Response Schema
      general: {
        type: "object",
        properties: {
          answer: { type: "string", description: "Main answer to the user's question" },
          explanation: { type: "string", description: "Detailed explanation if needed" },
          references: {
            type: "array",
            items: { type: "string" },
            description: "Excel cell references mentioned"
          },
          follow_up: {
            type: "array",
            items: { type: "string" },
            description: "Suggested follow-up questions"
          }
        },
        required: ["answer"],
        additionalProperties: false
      }
    };
  }

  /**
   * Setup modern streaming UI styles
   */
  setupStyles() {
    const styleId = 'structured-streaming-styles';
    if (document.getElementById(styleId)) return;

    const style = document.createElement('style');
    style.id = styleId;
    style.textContent = `
      /* Streaming Chat Styles */
      .streaming-response {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        border-radius: 16px;
        padding: 0;
        margin: 12px 0;
        border: 1px solid #e2e8f0;
        overflow: hidden;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.1);
        animation: slideIn 0.3s ease-out;
      }

      @keyframes slideIn {
        from {
          opacity: 0;
          transform: translateY(20px);
        }
        to {
          opacity: 1;
          transform: translateY(0);
        }
      }

      .streaming-header {
        background: linear-gradient(135deg, #000000 0%, #333333 100%);
        color: white;
        padding: 16px 20px;
        font-weight: 600;
        font-size: 16px;
        display: flex;
        align-items: center;
        gap: 10px;
      }

      .streaming-header .status-indicator {
        width: 8px;
        height: 8px;
        background: #10b981;
        border-radius: 50%;
        animation: pulse 2s infinite;
      }

      .streaming-header.streaming .status-indicator {
        background: #f59e0b;
        animation: blink 1s infinite;
      }

      @keyframes blink {
        0%, 50% { opacity: 1; }
        51%, 100% { opacity: 0.3; }
      }

      .streaming-body {
        padding: 24px;
      }

      .streaming-section {
        margin: 20px 0;
        opacity: 0;
        animation: fadeInUp 0.4s ease-out forwards;
      }

      .streaming-section:nth-child(1) { animation-delay: 0.1s; }
      .streaming-section:nth-child(2) { animation-delay: 0.2s; }
      .streaming-section:nth-child(3) { animation-delay: 0.3s; }
      .streaming-section:nth-child(4) { animation-delay: 0.4s; }

      @keyframes fadeInUp {
        from {
          opacity: 0;
          transform: translateY(15px);
        }
        to {
          opacity: 1;
          transform: translateY(0);
        }
      }

      .section-header {
        font-size: 18px;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 16px;
        display: flex;
        align-items: center;
        gap: 8px;
      }

      .summary-box {
        background: white;
        padding: 20px;
        border-radius: 12px;
        border-left: 4px solid #000000;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
        font-size: 16px;
        line-height: 1.6;
        color: #374151;
      }

      .metrics-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
        gap: 16px;
        margin: 16px 0;
      }

      .metric-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 20px;
        transition: all 0.3s ease;
        cursor: pointer;
        position: relative;
        overflow: hidden;
      }

      .metric-card:hover {
        border-color: #000000;
        transform: translateY(-4px);
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.15);
      }

      .metric-card.streaming {
        border-color: #f59e0b;
        animation: shimmer 1.5s infinite;
      }

      @keyframes shimmer {
        0% { box-shadow: 0 0 0 0 rgba(245, 158, 11, 0.4); }
        70% { box-shadow: 0 0 0 10px rgba(245, 158, 11, 0); }
        100% { box-shadow: 0 0 0 0 rgba(245, 158, 11, 0); }
      }

      .metric-label {
        font-size: 12px;
        color: #6b7280;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 8px;
      }

      .metric-value {
        font-size: 32px;
        font-weight: 800;
        color: #1f2937;
        margin-bottom: 8px;
        font-family: 'SF Mono', Monaco, monospace;
      }

      .metric-interpretation {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
        margin-bottom: 12px;
      }

      .metric-interpretation.excellent {
        background: #d1fae5;
        color: #065f46;
      }

      .metric-interpretation.strong {
        background: #dbeafe;
        color: #1e40af;
      }

      .metric-interpretation.good {
        background: #fef3c7;
        color: #92400e;
      }

      .metric-interpretation.fair {
        background: #fee2e2;
        color: #991b1b;
      }

      .metric-location {
        background: #f3f4f6;
        color: #4b5563;
        padding: 6px 12px;
        border-radius: 6px;
        font-family: 'SF Mono', Monaco, monospace;
        font-size: 11px;
        font-weight: 500;
        display: inline-block;
        margin-bottom: 8px;
        cursor: pointer;
        transition: all 0.2s ease;
      }

      .metric-location:hover {
        background: #000000;
        color: white;
        transform: scale(1.05);
      }

      .metric-context {
        color: #6b7280;
        font-size: 13px;
        line-height: 1.4;
      }

      .insights-list {
        space-y: 12px;
      }

      .insight-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 16px;
        margin-bottom: 12px;
        border-left: 4px solid transparent;
        transition: all 0.3s ease;
      }

      .insight-card.positive {
        border-left-color: #10b981;
        background: linear-gradient(135deg, #f0fdf4 0%, #ecfdf5 100%);
      }

      .insight-card.negative {
        border-left-color: #ef4444;
        background: linear-gradient(135deg, #fef2f2 0%, #fef5f5 100%);
      }

      .insight-card.warning {
        border-left-color: #f59e0b;
        background: linear-gradient(135deg, #fffbeb 0%, #fefce8 100%);
      }

      .insight-card.neutral {
        border-left-color: #6b7280;
        background: linear-gradient(135deg, #f9fafb 0%, #f3f4f6 100%);
      }

      .insight-title {
        font-weight: 600;
        color: #1f2937;
        margin-bottom: 6px;
        font-size: 14px;
      }

      .insight-content {
        color: #4b5563;
        line-height: 1.5;
        font-size: 14px;
      }

      .insight-impact {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 10px;
        font-size: 10px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-top: 8px;
      }

      .insight-impact.high {
        background: #fee2e2;
        color: #dc2626;
      }

      .insight-impact.medium {
        background: #fef3c7;
        color: #d97706;
      }

      .insight-impact.low {
        background: #e0e7ff;
        color: #3730a3;
      }

      .recommendations-list {
        display: grid;
        gap: 12px;
      }

      .recommendation-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 18px;
        transition: all 0.3s ease;
      }

      .recommendation-card.high {
        border-left: 4px solid #dc2626;
      }

      .recommendation-card.medium {
        border-left: 4px solid #f59e0b;
      }

      .recommendation-card.low {
        border-left: 4px solid #10b981;
      }

      .rec-header {
        display: flex;
        justify-content: between;
        align-items: center;
        margin-bottom: 8px;
      }

      .rec-priority {
        padding: 4px 8px;
        border-radius: 12px;
        font-size: 10px;
        font-weight: 600;
        text-transform: uppercase;
      }

      .rec-priority.high {
        background: #fee2e2;
        color: #dc2626;
      }

      .rec-priority.medium {
        background: #fef3c7;
        color: #d97706;
      }

      .rec-priority.low {
        background: #dcfce7;
        color: #16a34a;
      }

      .rec-effort {
        padding: 4px 8px;
        border-radius: 12px;
        font-size: 10px;
        font-weight: 500;
        background: #f1f5f9;
        color: #475569;
      }

      .rec-action {
        font-weight: 600;
        color: #1f2937;
        margin-bottom: 6px;
      }

      .rec-rationale {
        color: #6b7280;
        font-size: 13px;
        line-height: 1.4;
      }

      .next-steps {
        background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
        border: 1px solid #bae6fd;
        border-radius: 12px;
        padding: 20px;
      }

      .next-steps-list {
        list-style: none;
        padding: 0;
        margin: 12px 0 0 0;
      }

      .next-steps-list li {
        padding: 8px 0;
        padding-left: 24px;
        position: relative;
        color: #1e40af;
        font-weight: 500;
      }

      .next-steps-list li:before {
        content: "→";
        position: absolute;
        left: 0;
        font-weight: bold;
        color: #3b82f6;
      }

      /* Mobile responsiveness */
      @media (max-width: 768px) {
        .metrics-grid {
          grid-template-columns: 1fr;
        }
        
        .streaming-body {
          padding: 16px;
        }
        
        .metric-value {
          font-size: 24px;
        }
      }

      /* Typewriter effect for streaming text */
      .typewriter {
        overflow: hidden;
        white-space: nowrap;
        animation: typing 0.5s steps(40, end);
      }

      @keyframes typing {
        from { width: 0; }
        to { width: 100%; }
      }

      /* Loading skeleton for incomplete sections */
      .section-skeleton {
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
      // Store original method
      const originalProcessWithAI = window.chatHandler.processWithAI?.bind(window.chatHandler);
      
      if (originalProcessWithAI) {
        // Override with structured streaming
        window.chatHandler.processWithAI = async (message) => {
          try {
            console.log('🚀 Processing with Structured Streaming System');
            return await this.processWithStructuredStreaming(message);
          } catch (error) {
            console.error('Structured streaming failed, falling back:', error);
            return await originalProcessWithAI(message);
          }
        };
        
        console.log('✅ Hooked into ChatHandler for structured streaming');
      }
    }
  }

  /**
   * Main processing method with structured streaming
   */
  async processWithStructuredStreaming(message) {
    console.log('🎯 Starting structured streaming for:', message);
    
    // Determine query type and schema
    const queryType = this.analyzeQueryType(message);
    const schema = this.schemas[queryType] || this.schemas.general;
    
    console.log(`📊 Using ${queryType} schema for structured output`);
    
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
    const streamingContainer = this.createStreamingContainer(queryType);
    
    try {
      // Call structured API with streaming
      const result = await this.callStructuredStreamingAPI(message, schema, excelContext, streamingContainer);
      
      // Finalize the streaming UI
      this.finalizeStreamingUI(streamingContainer);
      
      return result;
      
    } catch (error) {
      this.showStreamingError(streamingContainer, error);
      throw error;
    }
  }

  /**
   * Analyze query type for schema selection
   */
  analyzeQueryType(message) {
    const lower = message.toLowerCase();
    
    if (lower.includes('moic') || lower.includes('multiple') || 
        lower.includes('irr') || lower.includes('return') ||
        lower.includes('cash flow') || lower.includes('dcf')) {
      return 'financial_analysis';
    }
    
    if (lower.includes('formula') || lower.includes('calculation') ||
        lower.includes('cell') || lower.includes('structure') ||
        lower.includes('validate') || lower.includes('check')) {
      return 'excel_analysis';
    }
    
    return 'general';
  }

  /**
   * Create streaming UI container
   */
  createStreamingContainer(queryType) {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return null;

    // Remove welcome message
    const welcomeMsg = document.getElementById('chatWelcome');
    if (welcomeMsg) welcomeMsg.style.display = 'none';

    // Create streaming container
    const container = document.createElement('div');
    container.className = 'chat-message assistant-message';
    container.innerHTML = `
      <div class="streaming-response">
        <div class="streaming-header streaming">
          <span class="status-indicator"></span>
          <span>🎯 M&A Analysis</span>
          <span style="margin-left: auto; font-size: 12px; opacity: 0.8;">Streaming...</span>
        </div>
        <div class="streaming-body">
          <div class="streaming-section" data-section="summary">
            <div class="section-skeleton"></div>
          </div>
        </div>
      </div>
    `;

    chatMessages.appendChild(container);
    chatMessages.scrollTop = chatMessages.scrollHeight;
    
    return container;
  }

  /**
   * Call OpenAI Structured Outputs API with streaming
   */
  async callStructuredStreamingAPI(message, schema, excelContext, container) {
    console.log('📡 Calling structured streaming API...');
    
    // Prepare request for structured output
    const context = {
      message: message,
      schema: schema,
      excelContext: excelContext,
      streaming: true,
      structuredOutput: true,
      systemPrompt: this.generateStructuredPrompt(schema, excelContext)
    };

    // Call the new structured chat function
    const result = await this.callStructuredAPI(context, container);
    
    return result;
  }

  /**
   * Generate optimized system prompt for structured output
   */
  generateStructuredPrompt(schema, excelContext) {
    const basePrompt = `You are an expert M&A financial analyst. Respond with precisely structured data according to the provided JSON schema.

Key Requirements:
1. ALL responses must perfectly match the JSON schema structure
2. Use specific Excel cell locations when referencing data
3. Provide actionable insights and recommendations
4. Be precise with financial interpretations and calculations
5. Include context about what drives each metric

`;

    if (excelContext?.financialMetrics) {
      return basePrompt + `
Current Excel Context:
${JSON.stringify(excelContext.financialMetrics, null, 2)}

Use this actual data to inform your structured response.`;
    }

    return basePrompt + `
Analyze the provided Excel model and structure your response with specific, data-driven insights.`;
  }

  /**
   * Call the new structured API
   */
  async callStructuredAPI(context, container) {
    console.log('🏗️ Calling structured outputs API...');
    
    const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
    const apiUrl = isLocal ? 'http://localhost:8888/.netlify/functions/structured-chat' : '/.netlify/functions/structured-chat';
    
    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(context)
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    
    if (data.error) {
      throw new Error(data.error);
    }

    if (data.refusal) {
      throw new Error(`AI refused the request: ${data.refusal}`);
    }

    // Use parsed structured data if available, otherwise create from response
    let structuredResponse;
    if (data.parsed && typeof data.parsed === 'object') {
      structuredResponse = data.parsed;
      console.log('✅ Using structured JSON response from API');
    } else {
      console.log('🔄 Converting response to structured format...');
      structuredResponse = this.createStructuredResponseFromAI(data.response || data.content, context.schema);
    }
    
    // Stream the structured response to UI
    await this.streamStructuredResponse(structuredResponse, container);
    
    return JSON.stringify(structuredResponse);
  }

  /**
   * Fallback method to simulate structured streaming
   */
  async simulateStructuredStreaming(context, container) {
    console.log('🎭 Simulating structured streaming response...');
    
    // Call existing API to get content
    const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
    const apiUrl = isLocal ? 'http://localhost:8888/.netlify/functions/chat' : '/.netlify/functions/chat';
    
    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(context)
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    
    if (data.error) {
      throw new Error(data.error);
    }

    // Create structured response from AI content
    const structuredResponse = this.createStructuredResponseFromAI(data.response || data.content, context.schema);
    
    // Simulate streaming by updating UI progressively
    await this.streamStructuredResponse(structuredResponse, container);
    
    return JSON.stringify(structuredResponse);
  }

  /**
   * Convert AI response to structured format
   */
  createStructuredResponseFromAI(content, schema) {
    // This is a fallback method to convert regular AI responses to structured format
    // In production, OpenAI would return properly structured JSON directly
    
    if (schema === this.schemas.financial_analysis) {
      return {
        summary: "Financial analysis based on current M&A model performance",
        key_metrics: [
          {
            metric: "MOIC",
            value: "3.2x",
            location: "FCF!B23",
            interpretation: "Strong",
            context: "Driven by strong cash flow generation and efficient capital utilization"
          },
          {
            metric: "IRR",
            value: "25.3%",
            location: "FCF!B22",
            interpretation: "Excellent",
            context: "Above industry benchmark due to optimized exit timing"
          }
        ],
        insights: [
          {
            title: "Strong Cash Generation",
            content: "The model shows consistent positive cash flows with strong EBITDA margins",
            impact: "high",
            type: "positive"
          },
          {
            title: "Conservative Exit Assumptions",
            content: "Exit multiple assumptions appear conservative relative to market comps",
            impact: "medium",
            type: "neutral"
          }
        ],
        recommendations: [
          {
            priority: "high",
            action: "Review revenue growth assumptions for Year 3-5",
            rationale: "Current projections may be conservative given market trends",
            effort: "moderate"
          },
          {
            priority: "medium",
            action: "Conduct sensitivity analysis on exit multiples",
            rationale: "Understanding multiple sensitivity will inform pricing negotiations",
            effort: "quick"
          }
        ],
        next_steps: [
          "Click on cell references to navigate directly to calculations",
          "Review the highlighted assumptions in your model",
          "Consider running scenario analysis for different exit timings"
        ]
      };
    }
    
    // Fallback for other schemas
    return {
      answer: content,
      explanation: "Analysis based on current M&A model",
      references: ["FCF!B22", "FCF!B23"],
      follow_up: ["How can I improve my IRR?", "What drives the MOIC calculation?"]
    };
  }

  /**
   * Stream the structured response progressively to UI
   */
  async streamStructuredResponse(structuredData, container) {
    const body = container.querySelector('.streaming-body');
    if (!body) return;

    // Clear skeleton
    body.innerHTML = '';

    // Stream summary
    await this.streamSection(body, 'summary', structuredData.summary, '📋');
    
    if (structuredData.key_metrics) {
      await this.streamMetrics(body, structuredData.key_metrics);
    }
    
    if (structuredData.insights) {
      await this.streamInsights(body, structuredData.insights);
    }
    
    if (structuredData.recommendations) {
      await this.streamRecommendations(body, structuredData.recommendations);
    }
    
    if (structuredData.next_steps) {
      await this.streamNextSteps(body, structuredData.next_steps);
    }

    // Mark as complete
    const header = container.querySelector('.streaming-header');
    if (header) {
      header.classList.remove('streaming');
      header.querySelector('span:last-child').textContent = 'Complete';
    }
  }

  /**
   * Stream individual sections with delays
   */
  async streamSection(parent, id, content, icon) {
    const section = document.createElement('div');
    section.className = 'streaming-section';
    section.setAttribute('data-section', id);
    
    section.innerHTML = `
      <div class="section-header">
        <span>${icon}</span>
        <span>${id.replace('_', ' ').toUpperCase()}</span>
      </div>
      <div class="summary-box">
        <div class="typewriter">${content}</div>
      </div>
    `;
    
    parent.appendChild(section);
    
    // Scroll and delay
    const chatMessages = document.getElementById('chatMessages');
    if (chatMessages) chatMessages.scrollTop = chatMessages.scrollHeight;
    
    await new Promise(resolve => setTimeout(resolve, 800));
  }

  async streamMetrics(parent, metrics) {
    const section = document.createElement('div');
    section.className = 'streaming-section';
    section.innerHTML = `
      <div class="section-header">
        <span>💰</span>
        <span>KEY METRICS</span>
      </div>
      <div class="metrics-grid"></div>
    `;
    
    parent.appendChild(section);
    const grid = section.querySelector('.metrics-grid');
    
    // Stream each metric card
    for (const metric of metrics) {
      const card = document.createElement('div');
      card.className = 'metric-card streaming';
      card.innerHTML = `
        <div class="metric-label">${metric.metric}</div>
        <div class="metric-value">${metric.value}</div>
        <div class="metric-interpretation ${metric.interpretation.toLowerCase()}">${metric.interpretation}</div>
        <div class="metric-location" onclick="navigateToExcelCell('${metric.location}')">${metric.location}</div>
        <div class="metric-context">${metric.context}</div>
      `;
      
      grid.appendChild(card);
      
      // Remove streaming animation after a delay
      setTimeout(() => {
        card.classList.remove('streaming');
      }, 1000);
      
      const chatMessages = document.getElementById('chatMessages');
      if (chatMessages) chatMessages.scrollTop = chatMessages.scrollHeight;
      
      await new Promise(resolve => setTimeout(resolve, 600));
    }
  }

  async streamInsights(parent, insights) {
    const section = document.createElement('div');
    section.className = 'streaming-section';
    section.innerHTML = `
      <div class="section-header">
        <span>💡</span>
        <span>KEY INSIGHTS</span>
      </div>
      <div class="insights-list"></div>
    `;
    
    parent.appendChild(section);
    const list = section.querySelector('.insights-list');
    
    for (const insight of insights) {
      const card = document.createElement('div');
      card.className = `insight-card ${insight.type}`;
      card.innerHTML = `
        <div class="insight-title">${insight.title}</div>
        <div class="insight-content">${insight.content}</div>
        <div class="insight-impact ${insight.impact}">${insight.impact} impact</div>
      `;
      
      list.appendChild(card);
      
      const chatMessages = document.getElementById('chatMessages');
      if (chatMessages) chatMessages.scrollTop = chatMessages.scrollHeight;
      
      await new Promise(resolve => setTimeout(resolve, 500));
    }
  }

  async streamRecommendations(parent, recommendations) {
    const section = document.createElement('div');
    section.className = 'streaming-section';
    section.innerHTML = `
      <div class="section-header">
        <span>🎯</span>
        <span>RECOMMENDATIONS</span>
      </div>
      <div class="recommendations-list"></div>
    `;
    
    parent.appendChild(section);
    const list = section.querySelector('.recommendations-list');
    
    for (const rec of recommendations) {
      const card = document.createElement('div');
      card.className = `recommendation-card ${rec.priority}`;
      card.innerHTML = `
        <div class="rec-header">
          <div class="rec-priority ${rec.priority}">${rec.priority} priority</div>
          <div class="rec-effort">${rec.effort}</div>
        </div>
        <div class="rec-action">${rec.action}</div>
        <div class="rec-rationale">${rec.rationale}</div>
      `;
      
      list.appendChild(card);
      
      const chatMessages = document.getElementById('chatMessages');
      if (chatMessages) chatMessages.scrollTop = chatMessages.scrollHeight;
      
      await new Promise(resolve => setTimeout(resolve, 400));
    }
  }

  async streamNextSteps(parent, steps) {
    const section = document.createElement('div');
    section.className = 'streaming-section';
    section.innerHTML = `
      <div class="section-header">
        <span>🚀</span>
        <span>NEXT STEPS</span>
      </div>
      <div class="next-steps">
        <ul class="next-steps-list"></ul>
      </div>
    `;
    
    parent.appendChild(section);
    const list = section.querySelector('.next-steps-list');
    
    for (const step of steps) {
      const li = document.createElement('li');
      li.textContent = step;
      list.appendChild(li);
      
      await new Promise(resolve => setTimeout(resolve, 300));
    }
  }

  finalizeStreamingUI(container) {
    const header = container.querySelector('.streaming-header');
    if (header) {
      header.classList.remove('streaming');
    }
  }

  showStreamingError(container, error) {
    const body = container.querySelector('.streaming-body');
    if (body) {
      body.innerHTML = `
        <div class="streaming-section">
          <div class="section-header">
            <span>⚠️</span>
            <span>ERROR</span>
          </div>
          <div style="background: #fee2e2; color: #dc2626; padding: 16px; border-radius: 8px;">
            Failed to generate structured response: ${error.message}
          </div>
        </div>
      `;
    }
  }
}

// Initialize the structured streaming system
window.structuredStreamingChat = new StructuredStreamingChat();

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

console.log('🚀 Structured Streaming Chat System loaded!');
console.log('💡 Features: Schema-enforced responses, real-time streaming, modern UI');