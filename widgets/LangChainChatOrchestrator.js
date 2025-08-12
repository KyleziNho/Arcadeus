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
      // Initialize LangGraph Web workflow (browser-compatible)
      if (window.ExcelLangGraphWorkflow) {
        this.langGraphWorkflow = new window.ExcelLangGraphWorkflow();
        console.log('✅ LangGraph Web workflow initialized');
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
        console.log('- window.ExcelLangGraphWorkflow:', typeof window.ExcelLangGraphWorkflow);
        console.log('- window.ExcelChatState:', typeof window.ExcelChatState);
        console.log('- this.langGraphWorkflow:', this.langGraphWorkflow);
        console.log('- this.chatState:', this.chatState);
        
        // Try to initialize now if components are available
        if (typeof window.ExcelLangGraphWorkflow !== 'undefined') {
          console.log('🔄 Attempting to initialize LangGraph Web now...');
          try {
            this.langGraphWorkflow = new window.ExcelLangGraphWorkflow();
            if (!this.chatState) {
              // Create initial state
              this.chatState = {
                messages: [],
                userIntent: null,
                toolResults: {},
                processingSteps: [],
                needsClarification: false,
                confidence: 0.0
              };
            }
            console.log('✅ LangGraph Web force-initialized successfully');
          } catch (initError) {
            console.error('❌ Force initialization failed:', initError);
          }
        }
        
        if (!this.langGraphWorkflow) {
          throw new Error("LangGraph still not initialized after retry. Please refresh the page and try again.");
        }
      }
      
      console.log('🌟 Using LangGraph Web for processing...');
      
      // Add user message to state
      if (!this.chatState.messages) this.chatState.messages = [];
      this.chatState.messages.push({
        role: 'user',
        content: userMessage,
        timestamp: new Date().toISOString()
      });
      
      // Stream through LangGraph Web workflow for real-time updates
      let finalState = null;
      const streamSteps = [];
      
      for await (const update of this.langGraphWorkflow.stream(this.chatState)) {
        console.log('📊 LangGraph Web update:', update.node, update.step);
        if (update.step) {
          streamSteps.push(update.step);
        }
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
      if (finalState && finalState.messages && finalState.messages.length > 0) {
        const finalMessage = finalState.messages[finalState.messages.length - 1];
        if (finalMessage && finalMessage.role === 'assistant') {
          const formattedResponse = await this.formatResponse(finalMessage.content);
          this.displayResponse(formattedResponse, null);
        }
      }
      
      // Save state to session storage (simplified for Web implementation)
      if (finalState) {
        sessionStorage.setItem("langgraph_web_state", JSON.stringify(finalState));
        console.log('💾 Saved LangGraph Web state with', finalState.messages?.length || 0, 'messages');
      }
      
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

  // OLD METHOD REMOVED - Using LangGraph workflow instead

  // OLD METHODS REMOVED - Using LangGraph workflow instead

  // OLD API METHODS REMOVED - Using LangGraph workflow instead

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