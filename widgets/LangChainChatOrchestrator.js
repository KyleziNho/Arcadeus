/**
 * LangChain Chat Orchestrator
 * Central handler for all chat inputs/outputs using LangChain
 * Integrates Excel search tools and ensures proper formatting
 */

class LangChainChatOrchestrator {
  constructor() {
    this.isInitialized = false;
    this.excelAgent = null;
    this.valueFinder = null;
    this.responseFormatter = null;
    this.chatHistory = [];
    this.maxHistoryLength = 10;
    
    this.initialize();
  }

  async initialize() {
    console.log('🚀 Initializing LangChain Chat Orchestrator...');
    
    try {
      // Initialize Excel AI Agent (LangChain)
      if (typeof ExcelAIAgent !== 'undefined') {
        this.excelAgent = new ExcelAIAgent();
        console.log('✅ Excel AI Agent initialized');
      }
      
      // Initialize Accurate Value Finder
      if (typeof AccurateExcelValueFinder !== 'undefined') {
        this.valueFinder = new AccurateExcelValueFinder();
        console.log('✅ Accurate Excel Value Finder initialized');
      }
      
      // Initialize Response Formatter
      if (typeof EnhancedResponseFormatter !== 'undefined') {
        this.responseFormatter = new EnhancedResponseFormatter();
        console.log('✅ Enhanced Response Formatter initialized');
      }
      
      this.isInitialized = true;
      console.log('✅ LangChain Chat Orchestrator ready');
      
      // Replace existing chat handler
      this.replaceDefaultChatHandler();
      
    } catch (error) {
      console.error('❌ Failed to initialize LangChain Chat Orchestrator:', error);
    }
  }

  /**
   * Replace the default chat handler with LangChain processing
   */
  replaceDefaultChatHandler() {
    // Override the global sendChatMessage function
    if (window.chatHandler) {
      const originalSend = window.chatHandler.sendChatMessage.bind(window.chatHandler);
      
      window.chatHandler.sendChatMessage = async () => {
        const chatInput = document.getElementById('chatInput');
        if (!chatInput || !chatInput.value.trim()) return;
        
        const message = chatInput.value.trim();
        chatInput.value = '';
        
        // Process through LangChain
        await this.processMessage(message);
      };
      
      console.log('✅ Default chat handler replaced with LangChain');
    }
  }

  /**
   * Main entry point for processing messages through LangChain
   */
  async processMessage(userMessage) {
    console.log('🎯 Processing message through LangChain:', userMessage);
    
    try {
      // Show user message in chat
      this.addUserMessage(userMessage);
      
      // Add to history
      this.chatHistory.push({ role: 'user', content: userMessage, timestamp: new Date().toISOString() });
      
      // Show processing indicator
      const processingContainer = this.showProcessingIndicator();
      
      // Step 1: Search Excel for relevant data
      const excelData = await this.searchExcelData(userMessage);
      
      // Step 2: Build comprehensive context
      const context = await this.buildContext(userMessage, excelData);
      
      // Step 3: Process through LangChain with tools
      const response = await this.processWithLangChain(userMessage, context);
      
      // Step 4: Format response
      const formattedResponse = await this.formatResponse(response);
      
      // Step 5: Display formatted response
      this.displayResponse(formattedResponse, processingContainer);
      
      // Add to history
      this.chatHistory.push({ role: 'assistant', content: formattedResponse, timestamp: new Date().toISOString() });
      
      // Trim history if needed
      if (this.chatHistory.length > this.maxHistoryLength) {
        this.chatHistory = this.chatHistory.slice(-this.maxHistoryLength);
      }
      
    } catch (error) {
      console.error('❌ Error processing message:', error);
      this.displayError('I encountered an error processing your request. Please try again.');
    }
  }

  /**
   * Search Excel for data relevant to the user's query
   */
  async searchExcelData(query) {
    console.log('🔍 Searching Excel for relevant data...');
    
    const searchResults = {
      metrics: {},
      cells: [],
      context: {}
    };
    
    try {
      // Use AccurateExcelValueFinder to get all financial metrics
      if (this.valueFinder) {
        const metrics = await this.valueFinder.findAllFinancialMetrics();
        searchResults.metrics = metrics;
        console.log('📊 Found metrics:', Object.keys(metrics));
      }
      
      // Search for specific terms mentioned in the query
      const searchTerms = this.extractSearchTerms(query);
      
      for (const term of searchTerms) {
        try {
          await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load('items');
            await context.sync();
            
            for (const worksheet of worksheets.items) {
              worksheet.load('name');
              const usedRange = worksheet.getUsedRangeOrNullObject();
              usedRange.load(['values', 'address']);
              
              await context.sync();
              
              if (!usedRange.isNullObject) {
                const values = usedRange.values;
                
                for (let row = 0; row < values.length; row++) {
                  for (let col = 0; col < values[row].length; col++) {
                    const cellValue = String(values[row][col]).toLowerCase();
                    
                    if (cellValue.includes(term.toLowerCase())) {
                      searchResults.cells.push({
                        worksheet: worksheet.name,
                        address: this.getCellAddress(row, col),
                        value: values[row][col],
                        matchedTerm: term
                      });
                    }
                  }
                }
              }
            }
          });
        } catch (error) {
          console.error(`Error searching for ${term}:`, error);
        }
      }
      
    } catch (error) {
      console.error('Error searching Excel:', error);
    }
    
    return searchResults;
  }

  /**
   * Extract search terms from user query
   */
  extractSearchTerms(query) {
    const terms = [];
    
    // Financial metric keywords
    const metricKeywords = ['irr', 'moic', 'revenue', 'ebitda', 'debt', 'equity', 'exit', 'value', 'return', 'multiple'];
    
    // Check for each keyword in the query
    for (const keyword of metricKeywords) {
      if (query.toLowerCase().includes(keyword)) {
        terms.push(keyword);
      }
    }
    
    // Extract cell references (e.g., B12, Sheet1!A1)
    const cellRefPattern = /[A-Z]+!?[A-Z]*\d+/gi;
    const cellRefs = query.match(cellRefPattern);
    if (cellRefs) {
      terms.push(...cellRefs);
    }
    
    // Extract numbers that might be values to search for
    const numberPattern = /\d+\.?\d*/g;
    const numbers = query.match(numberPattern);
    if (numbers) {
      terms.push(...numbers);
    }
    
    return [...new Set(terms)]; // Remove duplicates
  }

  /**
   * Build comprehensive context for LangChain
   */
  async buildContext(message, excelData) {
    const context = {
      userMessage: message,
      timestamp: new Date().toISOString(),
      excel: {
        metrics: excelData.metrics,
        searchResults: excelData.cells,
        hasData: Object.keys(excelData.metrics).length > 0
      },
      chatHistory: this.chatHistory.slice(-3), // Last 3 messages for context
      instructions: this.getSystemInstructions()
    };
    
    // Add form data if available
    if (window.formHandler) {
      try {
        context.formData = window.formHandler.collectAllModelData();
      } catch (error) {
        console.log('Could not collect form data:', error);
      }
    }
    
    return context;
  }

  /**
   * Process message through LangChain with all tools
   */
  async processWithLangChain(message, context) {
    console.log('🤖 Processing through LangChain with context:', context);
    
    // If Excel AI Agent is available, use it
    if (this.excelAgent) {
      try {
        const response = await this.excelAgent.processMessage(message);
        return response;
      } catch (error) {
        console.error('Excel AI Agent error:', error);
      }
    }
    
    // Otherwise, call the API with enhanced context
    return await this.callLangChainAPI(message, context);
  }

  /**
   * Call LangChain API endpoint
   */
  async callLangChainAPI(message, context) {
    const isLocal = window.location.hostname === 'localhost';
    const apiUrl = isLocal 
      ? 'http://localhost:8888/.netlify/functions/streaming-chat' 
      : '/.netlify/functions/streaming-chat';
    
    // Prepare the prompt with Excel data
    let excelContext = '';
    
    if (context.excel.metrics && Object.keys(context.excel.metrics).length > 0) {
      excelContext = '\n\nACTUAL EXCEL VALUES (from your workbook):\n';
      
      for (const [metric, data] of Object.entries(context.excel.metrics)) {
        excelContext += `\n${metric}:`;
        excelContext += `\n  • Value: ${data.value}`;
        excelContext += `\n  • Location: ${data.location}`;
        if (data.formula) {
          excelContext += `\n  • Formula: ${data.formula}`;
        }
      }
    }
    
    if (context.excel.searchResults && context.excel.searchResults.length > 0) {
      excelContext += '\n\nSEARCH RESULTS:\n';
      context.excel.searchResults.slice(0, 10).forEach(result => {
        excelContext += `\n• ${result.worksheet}!${result.address}: ${result.value}`;
      });
    }
    
    const enhancedMessage = `
You are an expert M&A financial analyst helping with Excel modeling.

User Question: ${message}

${excelContext}

Instructions:
1. Use ONLY the actual Excel values provided above
2. Always cite specific cell locations when referencing values
3. If a value is not provided, say "Value not found in Excel"
4. Provide actionable insights based on the real data
5. Format your response with clear sections and highlights

Please provide a comprehensive analysis.`;
    
    try {
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: enhancedMessage,
          streaming: false
        })
      });
      
      if (!response.ok) {
        throw new Error(`API error: ${response.status}`);
      }
      
      const data = await response.json();
      return data.response || data.parsed?.final_answer || 'No response received';
      
    } catch (error) {
      console.error('LangChain API error:', error);
      return this.generateFallbackResponse(message, context);
    }
  }

  /**
   * Format the response for display
   */
  async formatResponse(response) {
    if (this.responseFormatter) {
      return this.responseFormatter.formatResponse(response);
    }
    
    // Basic formatting if formatter not available
    return this.basicFormatResponse(response);
  }

  /**
   * Basic response formatting
   */
  basicFormatResponse(response) {
    let formatted = response;
    
    // Convert markdown headers
    formatted = formatted.replace(/### (.+)/g, '<h3 class="response-header">$1</h3>');
    formatted = formatted.replace(/## (.+)/g, '<h3 class="response-header">$1</h3>');
    
    // Convert bold text
    formatted = formatted.replace(/\*\*(.+?)\*\*/g, '<strong class="highlight-value">$1</strong>');
    
    // Convert cell references to clickable links
    formatted = formatted.replace(/([A-Z]+!?[A-Z]*\d+)/g, 
      '<span class="cell-reference clickable" onclick="window.navigateToCell(\'$1\')">$1</span>');
    
    // Convert bullet points
    formatted = formatted.replace(/^• (.+)$/gm, '<li class="response-bullet">$1</li>');
    formatted = formatted.replace(/(<li.*<\/li>\n?)+/g, '<ul class="response-list">$&</ul>');
    
    // Convert line breaks
    formatted = formatted.replace(/\n\n/g, '</p><p class="response-paragraph">');
    formatted = `<p class="response-paragraph">${formatted}</p>`;
    
    return formatted;
  }

  /**
   * Display formatted response in chat
   */
  displayResponse(formattedResponse, processingContainer) {
    if (processingContainer) {
      processingContainer.remove();
    }
    
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;
    
    const messageDiv = document.createElement('div');
    messageDiv.className = 'chat-message assistant-message langchain-response';
    
    messageDiv.innerHTML = `
      <div class="message-avatar">
        <div class="avatar-icon">🤖</div>
      </div>
      <div class="message-content">
        <div class="message-header">
          <span class="message-role">AI Assistant</span>
          <span class="message-badge">LangChain</span>
        </div>
        <div class="message-text formatted-response">
          ${formattedResponse}
        </div>
        <div class="message-footer">
          <span class="message-time">${new Date().toLocaleTimeString()}</span>
        </div>
      </div>
    `;
    
    chatMessages.appendChild(messageDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
    
    // Add click handlers for cell references
    this.addCellClickHandlers(messageDiv);
  }

  /**
   * Add click handlers for cell references
   */
  addCellClickHandlers(container) {
    const cellRefs = container.querySelectorAll('.cell-reference.clickable');
    cellRefs.forEach(ref => {
      ref.addEventListener('click', () => {
        const address = ref.textContent;
        this.navigateToCell(address);
      });
    });
  }

  /**
   * Navigate to Excel cell
   */
  async navigateToCell(address) {
    try {
      await Excel.run(async (context) => {
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
      
      console.log(`✅ Navigated to ${address}`);
    } catch (error) {
      console.error(`Error navigating to ${address}:`, error);
    }
  }

  /**
   * Show processing indicator
   */
  showProcessingIndicator() {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return null;
    
    const processingDiv = document.createElement('div');
    processingDiv.className = 'chat-message assistant-message processing';
    processingDiv.innerHTML = `
      <div class="message-avatar">
        <div class="avatar-icon">🤖</div>
      </div>
      <div class="message-content">
        <div class="processing-indicator">
          <div class="processing-dots">
            <span></span><span></span><span></span>
          </div>
          <div class="processing-text">Analyzing Excel data and preparing response...</div>
        </div>
      </div>
    `;
    
    chatMessages.appendChild(processingDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
    
    return processingDiv;
  }

  /**
   * Add user message to chat
   */
  addUserMessage(message) {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;
    
    const messageDiv = document.createElement('div');
    messageDiv.className = 'chat-message user-message';
    
    messageDiv.innerHTML = `
      <div class="message-content">
        <div class="message-text">${this.escapeHtml(message)}</div>
        <div class="message-time">${new Date().toLocaleTimeString()}</div>
      </div>
      <div class="message-avatar">
        <div class="avatar-icon">👤</div>
      </div>
    `;
    
    chatMessages.appendChild(messageDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }

  /**
   * Display error message
   */
  displayError(errorMessage) {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;
    
    const messageDiv = document.createElement('div');
    messageDiv.className = 'chat-message assistant-message error-message';
    
    messageDiv.innerHTML = `
      <div class="message-avatar">
        <div class="avatar-icon">⚠️</div>
      </div>
      <div class="message-content">
        <div class="message-text error-text">${errorMessage}</div>
      </div>
    `;
    
    chatMessages.appendChild(messageDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }

  /**
   * Generate fallback response when API fails
   */
  generateFallbackResponse(message, context) {
    let response = "I'll help you analyze your Excel model.\n\n";
    
    if (context.excel.metrics && Object.keys(context.excel.metrics).length > 0) {
      response += "**Key Metrics Found:**\n";
      
      for (const [metric, data] of Object.entries(context.excel.metrics)) {
        response += `• ${metric}: ${data.value} (Cell: ${data.location})\n`;
      }
      
      response += "\n**Analysis:**\n";
      response += "Based on the data in your Excel model, here are the key insights:\n";
      
      // Add specific insights based on metrics
      if (context.excel.metrics.IRR) {
        const irrValue = context.excel.metrics.IRR.rawValue;
        response += `• Your IRR of ${context.excel.metrics.IRR.value} `;
        response += irrValue > 0.15 ? "exceeds typical target returns\n" : "may need optimization\n";
      }
      
      if (context.excel.metrics.MOIC) {
        const moicValue = context.excel.metrics.MOIC.rawValue;
        response += `• Your MOIC of ${context.excel.metrics.MOIC.value} `;
        response += moicValue > 2.0 ? "shows strong value creation\n" : "has room for improvement\n";
      }
    } else {
      response += "I couldn't find specific metrics in your Excel model. ";
      response += "Please ensure your model contains labeled financial metrics.";
    }
    
    return response;
  }

  /**
   * Get system instructions for consistent behavior
   */
  getSystemInstructions() {
    return `
You are an expert M&A financial analyst integrated with Excel.
Always:
1. Use actual Excel values and cite cell references
2. Provide actionable insights based on real data
3. Format responses with clear sections
4. Highlight important metrics and values
5. Suggest improvements when relevant
Never:
1. Make up values not found in Excel
2. Provide generic advice without context
3. Ignore the user's specific question
    `.trim();
  }

  /**
   * Utility: Get cell address from indices
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

  /**
   * Utility: Escape HTML
   */
  escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }
}

// Initialize and make globally available
if (typeof window !== 'undefined') {
  window.LangChainChatOrchestrator = LangChainChatOrchestrator;
  
  // Auto-initialize when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      window.langChainOrchestrator = new LangChainChatOrchestrator();
      console.log('✅ LangChain Chat Orchestrator initialized globally');
    });
  } else {
    window.langChainOrchestrator = new LangChainChatOrchestrator();
    console.log('✅ LangChain Chat Orchestrator initialized globally');
  }
}