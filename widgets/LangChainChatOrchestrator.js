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
    
    // New LangGraph components
    this.langGraphWorkflow = null;
    this.memory = null;
    this.properTools = [];
    
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
      
      // Initialize new LangGraph system
      await this.initializeLangGraph();
      
      this.isInitialized = true;
      console.log('✅ LangChain Chat Orchestrator ready');
      
      // Replace existing chat handler
      this.replaceDefaultChatHandler();
      
    } catch (error) {
      console.error('❌ Failed to initialize LangChain Chat Orchestrator:', error);
    }
  }

  /**
   * Initialize LangGraph workflow system
   */
  async initializeLangGraph() {
    console.log('🌟 Initializing LangGraph workflow system...');
    
    try {
      // Initialize LangGraph workflow
      if (window.LangGraphExcelWorkflow) {
        this.langGraphWorkflow = new window.LangGraphExcelWorkflow();
        console.log('✅ LangGraph workflow initialized');
      }
      
      // Initialize state management
      if (window.ExcelChatState) {
        // Load saved state from session storage
        const savedState = sessionStorage.getItem("langgraph_chat_state");
        if (savedState) {
          this.chatState = window.ExcelChatState.deserialize(JSON.parse(savedState));
          console.log('📚 Loaded chat state:', this.chatState.messages.length, 'messages');
        } else {
          this.chatState = new window.ExcelChatState();
        }
        
        console.log('✅ LangGraph state management initialized');
      }
      
    } catch (error) {
      console.error('❌ Failed to initialize LangGraph:', error);
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
   * Main entry point for processing messages through LangGraph
   */
  async processMessage(userMessage) {
    console.log('🌟 Processing message through LangGraph:', userMessage);
    
    try {
      // Show user message in chat
      this.addUserMessage(userMessage);
      
      // Show processing indicator
      const processingContainer = this.showProcessingIndicator();
      
      // ALWAYS use LangGraph - NO FALLBACKS
      if (!this.langGraphWorkflow) {
        console.error('❌ LangGraph not initialized!');
        console.log('Available components:');
        console.log('- window.LangGraphExcelWorkflow:', typeof window.LangGraphExcelWorkflow);
        console.log('- window.ExcelChatState:', typeof window.ExcelChatState);
        throw new Error("LangGraph not initialized. Cannot process message without proper workflow.");
      }
      
      console.log('🌟 Using LangGraph for processing...');
      
      // Add user message to state
      this.chatState.addMessage({
        role: 'user',
        content: userMessage
      });
      
      // Stream through LangGraph workflow for real-time updates
      let finalState = null;
      const streamSteps = [];
      
      for await (const update of this.langGraphWorkflow.stream(this.chatState)) {
        console.log('📊 LangGraph update:', update.node, update.step);
        streamSteps.push(update.step);
        finalState = update.state;
      }
      
      console.log('🌟 LangGraph completed with', streamSteps.length, 'steps');
      
      // Display all processing steps
      if (streamSteps.length > 0) {
        console.log('📝 Displaying', streamSteps.length, 'LangGraph steps');
        this.displayLangGraphSteps(streamSteps, processingContainer);
      } else {
        console.log('⚠️ No processing steps found');
        if (processingContainer) {
          processingContainer.remove();
        }
      }
      
      // Get final response from state
      const finalMessage = finalState.messages[finalState.messages.length - 1];
      if (finalMessage && finalMessage.role === 'assistant') {
        const formattedResponse = await this.formatResponse(finalMessage.content);
        this.displayResponse(formattedResponse, null);
      }
      
      // Save state to session storage
      sessionStorage.setItem("langgraph_chat_state", JSON.stringify(finalState.serialize()));
      console.log('💾 Saved LangGraph state with', finalState.messages.length, 'messages');
      
    } catch (error) {
      console.error('❌ Error processing message:', error);
      this.displayLangChainError(error.message, processingContainer);
    }
  }


  /**
   * Display intermediate steps for transparency (following expert plan)
   */
  displayIntermediateSteps(steps, processingContainer) {
    // Remove processing indicator first
    if (processingContainer) {
      processingContainer.remove();
    }
    
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;
    
    // Create thinking steps container
    const stepsDiv = document.createElement('div');
    stepsDiv.className = 'chat-message assistant-message thinking-steps';
    
    let stepsHTML = `
      <div class="message-avatar">
        <div class="avatar-icon">🧠</div>
      </div>
      <div class="message-content">
        <div class="message-header">
          <span class="message-role">AI Thinking</span>
          <span class="message-badge thinking-badge">Step-by-Step</span>
        </div>
        <div class="thinking-container">
          <h4>🔍 Analysis Process:</h4>
    `;
    
    steps.forEach((step, index) => {
      stepsHTML += `
        <div class="thinking-step">
          <div class="step-number">${index + 1}</div>
          <div class="step-content">
            <div class="step-action">
              <strong>🔧 Action:</strong> ${step.action.tool}
            </div>
            <div class="step-input">
              <strong>📝 Input:</strong> <code>${JSON.stringify(step.action.toolInput)}</code>
            </div>
            <div class="step-observation">
              <strong>👀 Result:</strong> 
              <div class="observation-content">
                ${this.formatObservation(step.observation)}
              </div>
            </div>
          </div>
        </div>
      `;
    });
    
    stepsHTML += `
        </div>
        <div class="message-footer">
          <span class="message-time">${new Date().toLocaleTimeString()}</span>
        </div>
      </div>
    `;
    
    stepsDiv.innerHTML = stepsHTML;
    chatMessages.appendChild(stepsDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }

  /**
   * Format observation results for display
   */
  formatObservation(observation) {
    if (observation.success) {
      if (observation.metric && observation.value) {
        return `✅ Found <strong>${observation.metric}: ${observation.value}</strong> at ${observation.location}`;
      } else if (observation.formattedValue) {
        return `✅ Formula result: <strong>${observation.formattedValue}</strong>`;
      } else if (observation.data && observation.data.values) {
        return `✅ Read ${observation.data.values.length} rows from ${observation.data.address}`;
      } else if (observation.message) {
        return `✅ ${observation.message}`;
      }
    } else if (observation.error) {
      return `❌ ${observation.error}`;
    }
    
    // Fallback: show raw JSON for debugging
    return `<pre>${JSON.stringify(observation, null, 2)}</pre>`;
  }

  /**
   * Display LangGraph processing steps for transparency
   */
  displayLangGraphSteps(steps, processingContainer) {
    // Remove processing indicator first
    if (processingContainer) {
      processingContainer.remove();
    }
    
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;
    
    // Create LangGraph steps container
    const stepsDiv = document.createElement('div');
    stepsDiv.className = 'chat-message assistant-message langgraph-steps';
    
    let stepsHTML = `
      <div class="message-avatar">
        <div class="avatar-icon">🌟</div>
      </div>
      <div class="message-content">
        <div class="message-header">
          <span class="message-role">LangGraph Workflow</span>
          <span class="message-badge langgraph-badge">Step-by-Step</span>
        </div>
        <div class="langgraph-container">
          <h4>🔄 Workflow Execution:</h4>
    `;
    
    steps.forEach((step, index) => {
      const statusIcon = step.success ? '✅' : '❌';
      const stepClass = step.success ? 'success' : 'error';
      
      stepsHTML += `
        <div class="langgraph-step ${stepClass}">
          <div class="step-number">${step.stepNumber}</div>
          <div class="step-content">
            <div class="step-node">
              <strong>📍 Node:</strong> ${step.node}
            </div>
            <div class="step-action">
              <strong>🔧 Action:</strong> ${step.action}
            </div>
      `;
      
      if (step.input) {
        stepsHTML += `
            <div class="step-input">
              <strong>📝 Input:</strong> <code>${JSON.stringify(step.input)}</code>
            </div>
        `;
      }
      
      stepsHTML += `
            <div class="step-result">
              <strong>👀 Result:</strong> 
              <div class="result-content">
                ${statusIcon} ${step.result}
              </div>
            </div>
      `;
      
      if (step.details) {
        stepsHTML += `
            <div class="step-details">
              <strong>📋 Details:</strong>
              <div class="details-content">
                ${this.formatLangGraphDetails(step.details)}
              </div>
            </div>
        `;
      }
      
      if (step.error) {
        stepsHTML += `
            <div class="step-error">
              <strong>⚠️ Error:</strong> ${step.error}
            </div>
        `;
      }
      
      stepsHTML += `
          </div>
        </div>
      `;
    });
    
    stepsHTML += `
        </div>
        <div class="message-footer">
          <span class="message-time">${new Date().toLocaleTimeString()}</span>
        </div>
      </div>
    `;
    
    stepsDiv.innerHTML = stepsHTML;
    chatMessages.appendChild(stepsDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }

  /**
   * Format LangGraph step details for display
   */
  formatLangGraphDetails(details) {
    if (details.success && details.cellsFound) {
      return `Found ${details.cellsFound} cells matching "${details.searchTerm}" and formatted ${details.cellsFormatted} successfully`;
    } else if (details.value && details.location) {
      return `Found <strong>${details.metric}: ${details.value}</strong> at ${details.location}`;
    } else if (details.formattedValue) {
      return `Calculation result: <strong>${details.formattedValue}</strong>`;
    }
    
    return JSON.stringify(details, null, 2);
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
   * Process message through LangChain with all tools - ALWAYS use LangChain
   */
  async processWithLangChain(message, context) {
    console.log('🤖 Processing through LangChain with context:', context);
    
    // ALWAYS use LangChain API - no fallbacks or alternatives
    return await this.callLangChainAPI(message, context);
  }

  /**
   * Call LangChain API endpoint with proper tool calling
   */
  async callLangChainAPI(message, context) {
    console.log('🤖 Calling LangChain API with tool calling...');
    
    const isLocal = window.location.hostname === 'localhost';
    const apiUrl = isLocal 
      ? 'http://localhost:8888/.netlify/functions/streaming-chat' 
      : '/.netlify/functions/streaming-chat';
    
    try {
      // Step 1: Determine what Excel tools to use based on the query
      const toolsToUse = this.determineRequiredTools(message);
      console.log('📋 Tools to use:', toolsToUse);
      
      // Step 2: Execute tools to get actual data
      const toolResults = await this.executeTools(toolsToUse, message);
      console.log('🔧 Tool results:', toolResults);
      
      // Step 3: Build enhanced context with tool results
      let toolContext = '';
      if (Object.keys(toolResults).length > 0) {
        toolContext = '\n\nTOOL EXECUTION RESULTS (ACTUAL EXCEL DATA):\n';
        
        for (const [toolName, result] of Object.entries(toolResults)) {
          toolContext += `\n${toolName.toUpperCase()}:\n`;
          
          if (typeof result === 'object') {
            toolContext += JSON.stringify(result, null, 2);
          } else {
            toolContext += result;
          }
          toolContext += '\n';
        }
      }
      
      const enhancedMessage = `
You are an expert M&A financial analyst with access to live Excel data.

User Question: ${message}

${toolContext}

CRITICAL INSTRUCTIONS:
1. Use ONLY the actual data from TOOL EXECUTION RESULTS above
2. Always cite specific cell locations when mentioning values (e.g., "IRR of **25.3%** at B12")
3. If a value wasn't found by tools, say "Value not found in Excel"
4. The tool results contain the REAL Excel data - never make up numbers

FORMATTING REQUIREMENTS:
5. Structure your response with clear sections using ## headers
6. Use **bold** for all financial values and important metrics
7. Use • bullet points for insights and recommendations
8. Include cell references in format: SheetName!A1 or B12
9. Add insight boxes using:
   - 💡 for key insights
   - ⚠️ for warnings or concerns
   - ✅ for recommendations
10. NEVER create tables, grids, or side-by-side layouts
11. ALWAYS list metrics in simple vertical bullet point format
12. Always include an analysis section with actionable insights

CRITICAL: DO NOT CREATE ANY TABLES OR COMPLEX LAYOUTS. Use simple bullet points only.

RESPONSE STRUCTURE:
- ## Key Financial Metrics (simple bullet list with actual values and cell locations)
- ## Analysis (interpretation of the numbers)
- ## Insights and Recommendations (actionable advice)

Provide a comprehensive, well-formatted analysis using ONLY the actual tool results.`;
      
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: enhancedMessage,
          streaming: false,
          toolResults: toolResults // Include tool results for reference
        })
      });
      
      if (!response.ok) {
        throw new Error(`API error: ${response.status}`);
      }
      
      const data = await response.json();
      return data.response || data.parsed?.final_answer || 'No response received';
      
    } catch (error) {
      console.error('❌ LangChain API error:', error);
      // NO FALLBACKS - Always retry LangChain or show proper error
      throw new Error(`LangChain API failed: ${error.message}. Please check your connection and try again.`);
    }
  }

  /**
   * Determine which Excel tools to use based on the user's query
   */
  determineRequiredTools(message) {
    const tools = [];
    const lowerMessage = message.toLowerCase();
    
    // Check for financial metric queries
    const metricKeywords = ['irr', 'moic', 'revenue', 'ebitda', 'debt', 'equity', 'exit', 'deal', 'value'];
    const mentionedMetrics = metricKeywords.filter(keyword => lowerMessage.includes(keyword));
    
    if (mentionedMetrics.length > 0) {
      // If specific metrics mentioned, search for them
      tools.push({ name: 'search_financial_metrics', args: { metricType: 'All' } });
      
      // If asking about relationships between metrics
      if (mentionedMetrics.length > 1) {
        tools.push({ 
          name: 'calculate_metric_relationships', 
          args: { 
            metric1: mentionedMetrics[0].toUpperCase(), 
            metric2: mentionedMetrics[1].toUpperCase() 
          } 
        });
      }
    }
    
    // Check for cell references
    const cellRefPattern = /[A-Z]+!?[A-Z]*\d+/gi;
    const cellRefs = message.match(cellRefPattern);
    if (cellRefs) {
      cellRefs.forEach(cellRef => {
        tools.push({ name: 'read_cell_value', args: { cellAddress: cellRef } });
      });
    }
    
    // Check for value search queries
    const numberPattern = /\d+\.?\d*/g;
    const numbers = message.match(numberPattern);
    if (numbers && (lowerMessage.includes('find') || lowerMessage.includes('search') || lowerMessage.includes('where'))) {
      tools.push({ 
        name: 'search_value_in_workbook', 
        args: { searchTerm: numbers[0], maxResults: 5 } 
      });
    }
    
    // Always get worksheet summary for context
    tools.push({ name: 'get_worksheet_summary', args: {} });
    
    return tools;
  }

  /**
   * Execute the determined tools to get actual Excel data
   */
  async executeTools(toolsToUse, originalMessage) {
    const results = {};
    
    // Check if LangChainExcelTools is available
    if (!window.langChainExcelTools) {
      console.warn('⚠️ LangChainExcelTools not available');
      return results;
    }
    
    for (const toolSpec of toolsToUse) {
      try {
        console.log(`🔧 Executing tool: ${toolSpec.name}`, toolSpec.args);
        
        const result = await window.langChainExcelTools.executeTool(toolSpec.name, toolSpec.args);
        results[toolSpec.name] = JSON.parse(result);
        
        console.log(`✅ Tool ${toolSpec.name} completed`);
      } catch (error) {
        console.error(`❌ Tool ${toolSpec.name} failed:`, error);
        results[toolSpec.name] = { error: error.message };
      }
    }
    
    return results;
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
   * Enhanced response formatting for professional LangChain responses
   */
  basicFormatResponse(response) {
    let formatted = response;
    
    // FIRST: Remove any table HTML that might have been generated
    formatted = formatted.replace(/<table[^>]*>[\s\S]*?<\/table>/gi, '');
    formatted = formatted.replace(/<tr[^>]*>[\s\S]*?<\/tr>/gi, '');
    formatted = formatted.replace(/<td[^>]*>[\s\S]*?<\/td>/gi, '');
    formatted = formatted.replace(/<th[^>]*>[\s\S]*?<\/th>/gi, '');
    formatted = formatted.replace(/\|[^|\n]*\|[\s\S]*?\n/g, '');
    
    // Convert markdown headers with better styling
    formatted = formatted.replace(/### (.+)/g, '<h3 class="response-header">📊 $1</h3>');
    formatted = formatted.replace(/## (.+)/g, '<h2 class="response-header">🎯 $1</h2>');
    formatted = formatted.replace(/# (.+)/g, '<h1 class="response-header">⚡ $1</h1>');
    
    // Convert bold text with enhanced highlighting
    formatted = formatted.replace(/\*\*(.+?)\*\*/g, '<strong class="highlight-value">$1</strong>');
    
    // Convert cell references to clickable links with better formatting
    formatted = formatted.replace(/([A-Z]+!?[A-Z]*\d+)/g, 
      '<span class="cell-reference clickable" onclick="window.langChainOrchestrator.navigateToCell(\'$1\')">📍 $1</span>');
    
    // Convert financial values (numbers with % or $ or x)
    formatted = formatted.replace(/(\\$?[\d,]+\.?\d*%?x?)/g, '<span class="financial-value">$1</span>');
    
    // Convert bullet points with better icons
    formatted = formatted.replace(/^• (.+)$/gm, '<li class="response-bullet">$1</li>');
    formatted = formatted.replace(/^- (.+)$/gm, '<li class="response-bullet">$1</li>');
    
    // Wrap consecutive bullet points in proper lists
    formatted = formatted.replace(/(<li class="response-bullet">.*?<\/li>\s*)+/gs, '<ul class="response-list">$&</ul>');
    
    // Convert recommendations/insights boxes
    formatted = formatted.replace(/💡 (.+?)(?=\\n\\n|$)/gs, '<div class="insight-box"><div class="insight-title">💡 Insight</div><div class="insight-content">$1</div></div>');
    formatted = formatted.replace(/⚠️ (.+?)(?=\\n\\n|$)/gs, '<div class="warning-box"><div class="warning-title">⚠️ Warning</div><div class="warning-content">$1</div></div>');
    formatted = formatted.replace(/✅ (.+?)(?=\\n\\n|$)/gs, '<div class="success-box"><div class="success-title">✅ Recommendation</div><div class="success-content">$1</div></div>');
    
    // Convert line breaks properly
    formatted = formatted.replace(/\n\n/g, '</p><p class="response-paragraph">');
    
    // Wrap in paragraph tags if not already wrapped
    if (!formatted.includes('<p') && !formatted.includes('<h') && !formatted.includes('<div')) {
      formatted = `<p class="response-paragraph">${formatted}</p>`;
    }
    
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
   * Display LangChain error with retry option - NO fallbacks
   */
  displayLangChainError(errorMessage, processingContainer) {
    if (processingContainer) {
      processingContainer.remove();
    }
    
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;
    
    const messageDiv = document.createElement('div');
    messageDiv.className = 'chat-message assistant-message langchain-error';
    
    messageDiv.innerHTML = `
      <div class="message-avatar">
        <div class="avatar-icon">⚠️</div>
      </div>
      <div class="message-content">
        <div class="message-header">
          <span class="message-role">LangChain Error</span>
          <span class="message-badge error-badge">Connection Issue</span>
        </div>
        <div class="message-text error-content">
          <div class="error-title">Unable to process through LangChain</div>
          <div class="error-message">${errorMessage}</div>
          <div class="error-actions">
            <button class="retry-btn" onclick="window.langChainOrchestrator.retryLastMessage()">
              🔄 Retry with LangChain
            </button>
          </div>
        </div>
      </div>
    `;
    
    chatMessages.appendChild(messageDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }
  
  /**
   * Retry the last message
   */
  async retryLastMessage() {
    const lastUserMessage = this.chatHistory
      .slice()
      .reverse()
      .find(msg => msg.role === 'user');
    
    if (lastUserMessage) {
      console.log('🔄 Retrying last message with LangChain...');
      await this.processMessage(lastUserMessage.content);
    }
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