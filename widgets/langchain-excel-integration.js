/**
 * LangChain Excel Integration
 * Complete M&A Excel Add-in integration using LangChain for AI-powered financial analysis
 * This file replaces the existing chat system with a LangChain-based approach
 */

(function() {
  'use strict';

  // Check if we're in a proper Excel Add-in environment
  if (typeof Office === 'undefined' || typeof Excel === 'undefined') {
    console.warn('❌ Excel environment not detected. LangChain integration disabled.');
    return;
  }

  console.log('🚀 Initializing LangChain Excel Integration...');

  /**
   * LangChain-like Agent System for Excel
   * Simplified implementation that mimics LangChain functionality without full dependencies
   */
  class ExcelAIAgent {
    constructor() {
      this.tools = [];
      this.memory = [];
      this.isProcessing = false;
      this.maxIterations = 3;
      this.setupStyles();
      this.initializeTools();
      console.log('🤖 Excel AI Agent initialized');
    }

    /**
     * Initialize Excel-specific tools
     */
    initializeTools() {
      this.tools = [
        {
          name: 'read_range',
          description: 'Read values from Excel range',
          execute: this.readRange.bind(this)
        },
        {
          name: 'search_workbook',
          description: 'Search for financial values in workbook',
          execute: this.searchWorkbook.bind(this)
        },
        {
          name: 'get_financial_summary',
          description: 'Get financial metrics summary',
          execute: this.getFinancialSummary.bind(this)
        },
        {
          name: 'navigate_to_cell',
          description: 'Navigate to specific Excel cell',
          execute: this.navigateToCell.bind(this)
        }
      ];
    }


    /**
     * Main processing method - mimics LangChain agent execution
     */
    async processMessage(userMessage) {
      if (this.isProcessing) {
        console.log('Already processing a message');
        return;
      }

      this.isProcessing = true;

      try {
        console.log('🎯 Processing:', userMessage);

        // Add user message to UI and memory
        this.addUserMessage(userMessage);
        this.memory.push({ role: 'user', content: userMessage });

        // Create response container
        const responseContainer = this.createStreamingResponse();

        // Show thinking process
        this.updateThinking(responseContainer, 'Analyzing your Excel model...');

        // Determine which tools to use based on the message
        const plan = this.createExecutionPlan(userMessage);
        
        // Execute the plan
        const result = await this.executePlan(plan, responseContainer);

        // Generate final response
        const finalResponse = await this.generateResponse(userMessage, result, responseContainer);

        // Add to memory
        this.memory.push({ role: 'assistant', content: finalResponse });

        return finalResponse;

      } catch (error) {
        console.error('❌ Error processing message:', error);
        this.showError(error.message);
        return `Sorry, I encountered an error: ${error.message}`;
      } finally {
        this.isProcessing = false;
      }
    }

    /**
     * Create execution plan based on user message
     */
    createExecutionPlan(message) {
      const lowerMessage = message.toLowerCase();
      const plan = [];

      // Always get workbook structure first for context
      plan.push({ tool: 'get_financial_summary', priority: 1 });

      // Add specific tools based on message content
      if (lowerMessage.includes('irr') || lowerMessage.includes('return')) {
        plan.push({ tool: 'search_workbook', params: { searchTerm: 'IRR' }, priority: 2 });
      }

      if (lowerMessage.includes('moic') || lowerMessage.includes('multiple')) {
        plan.push({ tool: 'search_workbook', params: { searchTerm: 'MOIC' }, priority: 2 });
      }

      if (lowerMessage.includes('revenue') || lowerMessage.includes('sales')) {
        plan.push({ tool: 'search_workbook', params: { searchTerm: 'Revenue' }, priority: 2 });
      }

      // If asking about specific cells/ranges
      const cellMatch = message.match(/([A-Z]+!?[A-Z]*\d+)/);
      if (cellMatch) {
        plan.push({ 
          tool: 'read_range', 
          params: { address: cellMatch[1] }, 
          priority: 1 
        });
      }

      return plan.sort((a, b) => a.priority - b.priority);
    }

    /**
     * Execute the planned tools
     */
    async executePlan(plan, container) {
      const results = {};

      for (const step of plan) {
        try {
          this.updateThinking(container, `Using ${step.tool}...`);
          
          const tool = this.tools.find(t => t.name === step.tool);
          if (tool) {
            const result = await tool.execute(step.params || {});
            results[step.tool] = result;
            
            this.updateThinking(container, `✓ ${step.tool} completed`);
            await this.delay(300);
          }
        } catch (error) {
          console.error(`Error executing ${step.tool}:`, error);
          results[step.tool] = `Error: ${error.message}`;
        }
      }

      return results;
    }

    /**
     * Generate final response based on tool results using Netlify OpenAI
     */
    async generateResponse(userMessage, toolResults, container) {
      this.updateThinking(container, 'Generating AI analysis...');

      // Prepare context for OpenAI
      let excelContext = '';
      
      // Add tool results as context
      if (toolResults.get_financial_summary) {
        excelContext += `Financial Summary: ${toolResults.get_financial_summary}\n\n`;
      }

      if (toolResults.search_workbook) {
        excelContext += `Search Results: ${toolResults.search_workbook}\n\n`;
      }

      if (toolResults.read_range) {
        excelContext += `Cell Data: ${toolResults.read_range}\n\n`;
      }

      // Call Netlify function for OpenAI analysis
      try {
        const aiResponse = await this.callNetlifyOpenAI(userMessage, excelContext);
        
        // Format and display the AI response
        const formattedResponse = this.formatResponse(aiResponse);
        this.updateStreamingMessage(container, formattedResponse, true);

        return aiResponse;
      } catch (error) {
        console.error('AI analysis failed:', error);
        
        // Fallback to local analysis
        const fallbackResponse = this.generateContextualAnalysis(userMessage, toolResults);
        const formattedResponse = this.formatResponse(fallbackResponse);
        this.updateStreamingMessage(container, formattedResponse, true);
        
        return fallbackResponse;
      }
    }

    /**
     * Call Netlify function with OpenAI integration
     */
    async callNetlifyOpenAI(userMessage, excelContext) {
      const isLocal = window.location.hostname === 'localhost';
      const apiUrl = isLocal 
        ? 'http://localhost:8888/.netlify/functions/streaming-chat' 
        : '/.netlify/functions/streaming-chat';

      console.log('📡 Calling Netlify OpenAI function:', apiUrl);

      // Enhanced prompt for LangChain-style analysis
      const enhancedMessage = `As an expert M&A financial analyst using Excel, please analyze this question: "${userMessage}"

Excel Data Context:
${excelContext}

Please provide a comprehensive M&A analysis including:
1. Direct answer to the question
2. Relevant financial metrics and their implications
3. Specific Excel cell references where possible
4. M&A industry insights and recommendations
5. Any assumptions or considerations

Focus on M&A concepts like IRR, MOIC, DCF valuation, exit multiples, and leverage analysis.`;

      const requestBody = {
        message: enhancedMessage,
        excelContext: {
          summary: excelContext,
          available: true
        },
        streaming: false // Get structured response
      };

      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        body: JSON.stringify(requestBody)
      });

      if (!response.ok) {
        throw new Error(`Netlify API error: ${response.status} ${response.statusText}`);
      }

      const responseText = await response.text();
      let data;
      
      try {
        data = JSON.parse(responseText);
      } catch (parseError) {
        console.error('Invalid JSON response:', responseText);
        throw new Error(`Invalid response format from AI service`);
      }

      if (data.error) {
        throw new Error(data.error);
      }

      // Return the AI response
      if (data.parsed && data.parsed.final_answer) {
        return data.parsed.final_answer;
      } else if (data.response) {
        return data.response;
      } else {
        throw new Error('No valid response from AI service');
      }
    }

    /**
     * Generate contextual analysis based on the question and results
     */
    generateContextualAnalysis(userMessage, results) {
      const lowerMessage = userMessage.toLowerCase();
      let analysis = '**Analysis:**\n';

      if (lowerMessage.includes('irr')) {
        analysis += '• IRR (Internal Rate of Return) measures the annualized return rate for your investment\n';
        analysis += '• Levered IRR typically exceeds unlevered IRR due to debt financing benefits\n';
        analysis += '• Target IRR for M&A deals is typically 15-25% depending on risk profile\n\n';
      }

      if (lowerMessage.includes('moic')) {
        analysis += '• MOIC (Multiple on Invested Capital) shows total return as a multiple of initial investment\n';
        analysis += '• MOIC = Exit Value ÷ Initial Equity Investment\n';
        analysis += '• Target MOIC for private equity is typically 2.0x-3.0x over 3-7 years\n\n';
      }

      analysis += '**Recommendations:**\n';
      analysis += '• Review assumptions if returns seem too high or low\n';
      analysis += '• Consider sensitivity analysis on key variables\n';
      analysis += '• Validate exit multiples against market comparables\n';

      return analysis;
    }

    /**
     * Tool: Read Excel Range
     */
    async readRange(params) {
      try {
        let result;
        await Excel.run(async (context) => {
          let range;
          
          if (params.address && params.address.includes('!')) {
            const [sheetName, rangeAddr] = params.address.split('!');
            const sheet = context.workbook.worksheets.getItem(sheetName);
            range = sheet.getRange(rangeAddr);
          } else {
            range = context.workbook.getSelectedRange();
          }
          
          range.load(['values', 'address', 'formulas']);
          await context.sync();
          
          result = {
            address: range.address,
            values: range.values,
            formulas: range.formulas
          };
        });

        return `Range ${result.address}:\nValues: ${JSON.stringify(result.values)}\nFormulas: ${JSON.stringify(result.formulas)}`;
      } catch (error) {
        return `Error reading range: ${error.message}`;
      }
    }

    /**
     * Tool: Search Workbook
     */
    async searchWorkbook(params) {
      try {
        const searchTerm = params.searchTerm || '';
        const results = [];

        await Excel.run(async (context) => {
          const worksheets = context.workbook.worksheets;
          worksheets.load('items');
          await context.sync();

          for (const worksheet of worksheets.items) {
            worksheet.load('name');
            const usedRange = worksheet.getUsedRange();
            usedRange.load(['values', 'address']);
            
            await context.sync();

            if (!usedRange.isNullObject && usedRange.values) {
              for (let row = 0; row < usedRange.values.length; row++) {
                for (let col = 0; col < usedRange.values[row].length; col++) {
                  const cellValue = String(usedRange.values[row][col]).toLowerCase();
                  
                  if (cellValue.includes(searchTerm.toLowerCase())) {
                    const cellAddress = this.getCellAddress(row, col);
                    results.push({
                      sheet: worksheet.name,
                      address: `${worksheet.name}!${cellAddress}`,
                      value: usedRange.values[row][col]
                    });
                  }
                }
              }
            }
          }
        });

        if (results.length === 0) {
          return `No results found for "${searchTerm}"`;
        }

        return results.slice(0, 10).map(r => `${r.address}: ${r.value}`).join('\n');
      } catch (error) {
        return `Error searching: ${error.message}`;
      }
    }

    /**
     * Tool: Get Financial Summary
     */
    async getFinancialSummary() {
      try {
        const metrics = {};

        await Excel.run(async (context) => {
          const worksheets = context.workbook.worksheets;
          worksheets.load('items');
          await context.sync();

          const financialTerms = {
            'IRR': ['irr', 'internal rate', 'return rate'],
            'MOIC': ['moic', 'multiple', 'money on invested capital'],
            'Revenue': ['revenue', 'sales', 'income'],
            'EBITDA': ['ebitda', 'earnings'],
            'Exit Value': ['exit value', 'terminal value']
          };

          for (const worksheet of worksheets.items) {
            worksheet.load('name');
            const usedRange = worksheet.getUsedRange();
            usedRange.load(['values']);
            
            await context.sync();

            if (!usedRange.isNullObject && usedRange.values) {
              for (let row = 0; row < usedRange.values.length; row++) {
                for (let col = 0; col < usedRange.values[row].length; col++) {
                  const cellText = String(usedRange.values[row][col]).toLowerCase();
                  
                  for (const [metric, terms] of Object.entries(financialTerms)) {
                    if (terms.some(term => cellText.includes(term))) {
                      const rightValue = usedRange.values[row][col + 1];
                      if (typeof rightValue === 'number' && rightValue !== 0) {
                        metrics[metric] = {
                          value: rightValue,
                          location: `${worksheet.name}!${this.getCellAddress(row, col + 1)}`
                        };
                      }
                    }
                  }
                }
              }
            }
          }
        });

        if (Object.keys(metrics).length === 0) {
          return 'No financial metrics found in the workbook.';
        }

        return Object.entries(metrics)
          .map(([metric, data]) => `${metric}: ${data.value} (${data.location})`)
          .join('\n');
      } catch (error) {
        return `Error getting financial summary: ${error.message}`;
      }
    }

    /**
     * Tool: Navigate to Cell
     */
    async navigateToCell(params) {
      try {
        await Excel.run(async (context) => {
          const address = params.address || params.cellAddress;
          
          let range;
          if (address.includes('!')) {
            const [sheetName, rangeAddr] = address.split('!');
            const sheet = context.workbook.worksheets.getItem(sheetName);
            range = sheet.getRange(rangeAddr);
          } else {
            range = context.workbook.getSelectedRange().worksheet.getRange(address);
          }
          
          range.select();
          await context.sync();
        });

        return `Navigated to ${params.address}`;
      } catch (error) {
        return `Error navigating: ${error.message}`;
      }
    }

    /**
     * Add user message to UI
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
          <div class="langchain-response">
            <div class="thinking-section" id="thinkingSection"></div>
            <div class="response-content" id="responseContent"></div>
          </div>
        </div>
      `;

      chatMessages.appendChild(assistantDiv);
      chatMessages.scrollTop = chatMessages.scrollHeight;

      return assistantDiv;
    }

    /**
     * Update thinking process display
     */
    updateThinking(container, message) {
      if (!container) return;
      
      const thinkingSection = container.querySelector('#thinkingSection');
      if (thinkingSection) {
        thinkingSection.innerHTML = `<div class="thinking-message">🤔 ${message}</div>`;
      }
    }

    /**
     * Update streaming message
     */
    updateStreamingMessage(container, content, isComplete = false) {
      if (!container) return;

      const responseContent = container.querySelector('#responseContent');
      if (responseContent) {
        responseContent.innerHTML = content;
      }

      if (isComplete) {
        // Hide thinking section
        const thinkingSection = container.querySelector('#thinkingSection');
        if (thinkingSection) {
          thinkingSection.style.display = 'none';
        }

        // Add click handlers for values and cell references
        this.addClickHandlers(container);
      }

      // Scroll to show new content
      const chatMessages = document.getElementById('chatMessages');
      if (chatMessages) {
        chatMessages.scrollTop = chatMessages.scrollHeight;
      }
    }

    /**
     * Format response with highlighting
     */
    formatResponse(text) {
      let formatted = text;

      // Highlight cell references
      formatted = formatted.replace(/([A-Z]+!?[A-Z]+\d+)/g, 
        '<span class="cell-ref" onclick="window.navigateToCell(\'$1\')">$1</span>');

      // Highlight percentages
      formatted = formatted.replace(/(\d+\.?\d*%)/g, 
        '<span class="financial-value">$1</span>');

      // Highlight currency values
      formatted = formatted.replace(/(\$\d+(?:,\d{3})*(?:\.\d{2})?[MB]?)/g, 
        '<span class="financial-value">$1</span>');

      // Highlight multiples
      formatted = formatted.replace(/(\d+\.?\d*x)/gi, 
        '<span class="financial-value">$1</span>');

      // Convert markdown formatting
      formatted = formatted.replace(/\*\*(.*?)\*\*/g, '<span class="bold-title">$1</span>');
      formatted = formatted.replace(/\n/g, '<br>');

      return formatted;
    }

    /**
     * Add click handlers to interactive elements
     */
    addClickHandlers(container) {
      // Cell references
      const cellRefs = container.querySelectorAll('.cell-ref');
      cellRefs.forEach(ref => {
        ref.style.cursor = 'pointer';
        ref.addEventListener('click', () => {
          const address = ref.textContent;
          window.navigateToCell(address);
        });
      });
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
          <div class="langchain-response error">
            <p style="color: #dc2626;">❌ Error: ${this.escapeHtml(message)}</p>
          </div>
        </div>
      `;

      chatMessages.appendChild(errorDiv);
      chatMessages.scrollTop = chatMessages.scrollHeight;
    }

    /**
     * Setup styles
     */
    setupStyles() {
      const styleId = 'langchain-integration-styles';
      if (document.getElementById(styleId)) return;

      const style = document.createElement('style');
      style.id = styleId;
      style.textContent = `
        .langchain-response {
          background: white;
          border-radius: 12px;
          padding: 20px;
          margin: 16px 0;
          border: 1px solid #e2e8f0;
          box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
          line-height: 1.6;
          font-size: 15px;
          color: #374151;
        }

        .thinking-section {
          background: #f8fafc;
          border: 1px solid #e2e8f0;
          border-radius: 8px;
          padding: 12px;
          margin-bottom: 16px;
          font-size: 14px;
          color: #64748b;
        }

        .thinking-message {
          animation: pulse 1.5s infinite;
        }

        @keyframes pulse {
          0%, 100% { opacity: 0.6; }
          50% { opacity: 1; }
        }

        .financial-value {
          background: #dcfce7;
          color: #15803d;
          padding: 2px 6px;
          border-radius: 4px;
          font-weight: 600;
          font-family: 'SF Mono', Monaco, monospace;
        }

        .cell-ref {
          background: #dcfce7;
          color: #15803d;
          padding: 2px 6px;
          border-radius: 4px;
          font-weight: 600;
          cursor: pointer;
          font-family: 'SF Mono', Monaco, monospace;
          border: 1px solid #86efac;
          transition: all 0.2s ease;
        }

        .cell-ref:hover {
          background: #22c55e;
          color: white;
          transform: translateY(-1px);
          box-shadow: 0 2px 8px rgba(34, 197, 94, 0.3);
        }

        .bold-title {
          color: #1e293b;
          font-weight: 700;
          font-size: 16px;
        }

        .langchain-response.error {
          border-color: #ef4444;
          background: #fef2f2;
        }
      `;

      document.head.appendChild(style);
    }

    /**
     * Utility functions
     */
    getCellAddress(row, col) {
      let columnLetter = '';
      let temp = col;
      while (temp >= 0) {
        columnLetter = String.fromCharCode(65 + (temp % 26)) + columnLetter;
        temp = Math.floor(temp / 26) - 1;
      }
      return `${columnLetter}${row + 1}`;
    }

    escapeHtml(text) {
      const div = document.createElement('div');
      div.textContent = text;
      return div.innerHTML;
    }

    delay(ms) {
      return new Promise(resolve => setTimeout(resolve, ms));
    }
  }

  /**
   * Global Integration Manager
   */
  class LangChainIntegration {
    constructor() {
      this.agent = new ExcelAIAgent();
      this.isInitialized = false;
    }

    async initialize() {
      console.log('🔧 Setting up LangChain integration...');

      // Override existing chat system
      this.overrideExistingSystem();
      
      // Setup UI handlers
      this.setupUIHandlers();

      this.isInitialized = true;
      console.log('✅ LangChain integration complete');
    }

    overrideExistingSystem() {
      // Override sendQuickMessage (for pre-selected prompts)
      window.sendQuickMessage = (message) => {
        console.log('📝 LangChain handling quick message:', message);
        const chatInput = document.getElementById('chatInput');
        if (chatInput) {
          chatInput.value = message;
        }
        this.agent.processMessage(message);
      };

      // Override sendMessageToOpenAI (main chat processing)
      window.sendMessageToOpenAI = async (message) => {
        console.log('🤖 LangChain handling OpenAI message:', message);
        return await this.agent.processMessage(message);
      };

      // Override handleSendMessageWithExcelContext
      window.handleSendMessageWithExcelContext = async () => {
        const chatInput = document.getElementById('chatInput');
        const message = chatInput.value.trim();
        
        if (!message) return;
        
        console.log('📊 LangChain handling Excel context message:', message);
        chatInput.value = '';
        
        // Hide welcome elements
        this.hideWelcomeElements();
        
        return await this.agent.processMessage(message);
      };

      // Override chatHandler if it exists
      if (window.chatHandler) {
        window.chatHandler.processWithAI = async (message) => {
          return await this.agent.processMessage(message);
        };
      } else {
        window.chatHandler = {
          processWithAI: async (message) => {
            return await this.agent.processMessage(message);
          }
        };
      }
    }

    setupUIHandlers() {
      // Main send button
      const sendButton = document.getElementById('sendButton');
      if (sendButton) {
        sendButton.onclick = () => window.handleSendMessageWithExcelContext();
      }

      // Chat input Enter key
      const chatInput = document.getElementById('chatInput');
      if (chatInput) {
        chatInput.onkeydown = (e) => {
          if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            window.handleSendMessageWithExcelContext();
          }
        };

        // Auto-expanding textarea
        chatInput.oninput = function() {
          this.style.height = 'auto';
          this.style.height = Math.min(this.scrollHeight, 120) + 'px';
        };
      }
    }

    hideWelcomeElements() {
      const welcome = document.getElementById('chatWelcome');
      if (welcome) welcome.style.display = 'none';

      const tipSection = document.querySelector('.chat-tip-section');
      if (tipSection) tipSection.style.display = 'none';

      const upgradeCard = document.querySelector('.upgrade-card');
      if (upgradeCard) upgradeCard.style.display = 'none';
    }
  }

  // Global navigation functions
  window.navigateToCell = async function(cellAddress) {
    console.log('🎯 Navigating to cell:', cellAddress);
    
    try {
      await Excel.run(async (context) => {
        let range;
        
        if (cellAddress.includes('!')) {
          const [sheetName, rangeAddr] = cellAddress.split('!');
          const sheet = context.workbook.worksheets.getItem(sheetName);
          range = sheet.getRange(rangeAddr);
        } else {
          range = context.workbook.getSelectedRange().worksheet.getRange(cellAddress);
        }
        
        range.select();
        await context.sync();
        console.log('✅ Navigated to', cellAddress);
      });
    } catch (error) {
      console.error('❌ Navigation failed:', error);
    }
  };

  // Initialize the system
  const integration = new LangChainIntegration();
  
  // Auto-initialize when Office is ready
  Office.onReady(async () => {
    try {
      await integration.initialize();
    } catch (error) {
      console.error('Failed to initialize LangChain integration:', error);
    }
  });

  // Expose globally for debugging
  window.langchainIntegration = integration;

  console.log('🤖 LangChain Excel Integration loaded successfully!');

})();