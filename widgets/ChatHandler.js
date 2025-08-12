class ChatHandler {
  constructor() {
    this.chatMessages = [];
    this.isProcessing = false;
    this.excelContext = null; // Store Excel data for all AI interactions
    this.contextLastUpdated = null;
    
    // Initialize SafeExcelContext for reading Excel data
    try {
      if (typeof SafeExcelContext !== 'undefined') {
        this.safeExcelContext = new SafeExcelContext();
        console.log('✅ SafeExcelContext initialized successfully');
      } else {
        console.log('⚠️ SafeExcelContext not available');
        this.safeExcelContext = null;
      }
    } catch (error) {
      console.error('❌ Failed to initialize SafeExcelContext:', error);
      this.safeExcelContext = null;
    }
    
    // Initialize ExcelLiveAnalyzer safely
    try {
      if (typeof ExcelLiveAnalyzer !== 'undefined') {
        this.excelAnalyzer = new ExcelLiveAnalyzer();
        console.log('✅ ExcelLiveAnalyzer initialized successfully');
      } else {
        console.log('⚠️ ExcelLiveAnalyzer not available, using basic mode');
        this.excelAnalyzer = null;
      }
    } catch (error) {
      console.error('❌ Failed to initialize ExcelLiveAnalyzer:', error);
      this.excelAnalyzer = null;
    }
  }

  async initialize() {
    console.log('Initializing Hebbia-inspired chat handler...');
    
    // Load Excel context immediately when chat opens
    await this.loadExcelContext();
    
    // Find chat elements
    const sendChatBtn = document.getElementById('sendChatBtn');
    const chatInput = document.getElementById('chatInput');
    
    if (sendChatBtn) {
      sendChatBtn.addEventListener('click', () => this.sendChatMessage());
      console.log('Send chat button listener added');
    }
    
    if (chatInput) {
      chatInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          this.sendChatMessage();
        }
      });
      
      chatInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          this.sendChatMessage();
        }
      });
      console.log('Chat input listeners added');
    }
    
    // Initialize Hebbia-style live monitoring
    try {
      if (this.excelAnalyzer) {
        console.log('🔄 Starting live Excel monitoring for real-time analysis...');
        
        // Add change listener for proactive insights
        this.excelAnalyzer.addChangeListener((eventType, data) => {
          console.log(`📊 Excel change detected: ${eventType}`, data);
          
          // Cache invalidation and proactive analysis
          if (eventType === 'data_change' && data.updatedData) {
            console.log('💡 Proactive analysis: Excel data changed, updating context cache');
          }
        });
        
        // Start the monitoring system
        await this.excelAnalyzer.startLiveMonitoring();
        console.log('✅ Live Excel monitoring started');
      } else {
        console.log('⚠️ ExcelLiveAnalyzer not available, using basic mode');
      }
    } catch (error) {
      console.error('❌ Failed to start live monitoring:', error);
      console.log('📋 Continuing with basic chat functionality');
    }
    
    console.log('✅ Advanced chat handler initialized with multi-agent architecture');
  }

  /**
   * Load Excel worksheet data to use as context for all AI interactions
   */
  async loadExcelContext() {
    console.log('📊 Loading Excel context for chat...');
    
    if (!this.safeExcelContext) {
      console.log('⚠️ SafeExcelContext not available, skipping Excel context loading');
      return;
    }

    try {
      // Get comprehensive Excel context
      const excelData = await this.safeExcelContext.getComprehensiveContext();
      
      if (excelData && !excelData.error) {
        this.excelContext = excelData;
        this.contextLastUpdated = new Date().toISOString();
        
        console.log('✅ Excel context loaded successfully:', {
          worksheet: excelData.worksheetName,
          hasData: excelData.hasActualData,
          nonEmptyCells: excelData.totalNonEmptyCells,
          usedRange: excelData.usedRange?.address
        });
        
        // Show context loaded indicator to user
        this.showContextStatus('Excel worksheet loaded as context');
      } else {
        console.log('⚠️ No Excel context available:', excelData?.error);
        this.excelContext = null;
      }
    } catch (error) {
      console.error('❌ Failed to load Excel context:', error);
      this.excelContext = null;
    }
  }

  /**
   * Show context status to user
   */
  showContextStatus(message) {
    const statusElement = document.getElementById('chat-context-status');
    if (statusElement) {
      statusElement.textContent = message;
      statusElement.style.display = 'block';
      
      // Auto-hide after 3 seconds
      setTimeout(() => {
        statusElement.style.display = 'none';
      }, 3000);
    } else {
      // Create status indicator if it doesn't exist
      const chatContainer = document.getElementById('chatMessages');
      if (chatContainer) {
        const statusDiv = document.createElement('div');
        statusDiv.id = 'chat-context-status';
        statusDiv.textContent = message;
        statusDiv.style.cssText = `
          background: #e7f5e7;
          color: #2d5a2d;
          padding: 8px 12px;
          border-radius: 6px;
          font-size: 12px;
          margin-bottom: 8px;
          text-align: center;
          border-left: 3px solid #4caf50;
        `;
        chatContainer.insertBefore(statusDiv, chatContainer.firstChild);
        
        // Auto-hide after 3 seconds
        setTimeout(() => {
          statusDiv.style.display = 'none';
        }, 3000);
      }
    }
  }

  /**
   * Refresh Excel context when worksheet changes
   */
  async refreshExcelContext() {
    console.log('🔄 Refreshing Excel context...');
    await this.loadExcelContext();
    this.showContextStatus('Excel context refreshed');
  }

  /**
   * Get Excel context info for display
   */
  getExcelContextInfo() {
    if (!this.excelContext) {
      return null;
    }
    
    return {
      worksheet: this.excelContext.worksheetName,
      hasData: this.excelContext.hasActualData,
      dataPoints: this.excelContext.totalNonEmptyCells,
      lastUpdated: this.contextLastUpdated,
      usedRange: this.excelContext.usedRange?.address
    };
  }

  async sendChatMessage() {
    const chatInput = document.getElementById('chatInput');
    if (!chatInput) {
      console.error('Chat input not found');
      return;
    }

    const message = chatInput.value.trim();
    if (!message) {
      console.log('Empty message, not sending');
      return;
    }

    if (this.isProcessing) {
      console.log('Already processing a message, please wait');
      return;
    }

    console.log('Sending chat message via LangChain orchestrator:', message);
    
    // **ALWAYS use LangChain orchestrator - NO FALLBACKS**
    if (!window.langChainOrchestrator || typeof window.langChainOrchestrator.processMessage !== 'function') {
      console.error('❌ LangChain orchestrator not available! Cannot process message.');
      this.showError('LangChain system not initialized. Please refresh the page and try again.');
      return;
    }
    
    console.log('🌟 Processing message with LangChain orchestrator');
    
    // Clear input immediately
    chatInput.value = '';
    
    try {
      this.isProcessing = true;
      await window.langChainOrchestrator.processMessage(message);
    } catch (error) {
      console.error('❌ LangChain orchestrator failed:', error);
      this.showError(`LangChain processing failed: ${error.message}`);
    } finally {
      this.isProcessing = false;
    }
  }

  /**
   * Show error message to user
   */
  showError(message) {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;
    
    const messageDiv = document.createElement('div');
    messageDiv.className = 'chat-message assistant-message error-message';
    
    messageDiv.innerHTML = `
      <div class="message-avatar">
        <div class="avatar-icon">⚠️</div>
      </div>
      <div class="message-content">
        <div class="message-header">
          <span class="message-role">System Error</span>
          <span class="message-badge error-badge">Critical</span>
        </div>
        <div class="message-text error-text">${message}</div>
        <div class="message-footer">
          <span class="message-time">${new Date().toLocaleTimeString()}</span>
        </div>
      </div>
    `;
    
    chatMessages.appendChild(messageDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }

  // LEGACY METHOD - No longer used since we're using LangChain orchestrator exclusively
  // This method remains for reference but is not called anymore
  async processWithAI(message) {
    console.log('⚠️ LEGACY: This method is no longer used. All processing goes through LangChain orchestrator.');
    console.log('🎯 Processing message with Hebbia-inspired multi-agent approach:', message);
    
    try {
      // Use pre-loaded Excel context for faster response
      let excelContext = this.excelContext;
      
      if (!excelContext) {
        console.log('📊 No pre-loaded Excel context, attempting to load now...');
        await this.loadExcelContext();
        excelContext = this.excelContext;
      }
      
      if (excelContext) {
        console.log('✅ Using Excel context for AI:', {
          worksheet: excelContext.worksheetName,
          hasData: excelContext.hasActualData,
          nonEmptyCells: excelContext.totalNonEmptyCells,
          lastUpdated: this.contextLastUpdated
        });
      } else {
        console.log('⚠️ No Excel context available for AI');
      }

      // Get current form data
      let formData = {};
      if (window.formHandler) {
        try {
          formData = window.formHandler.collectAllModelData();
        } catch (error) {
          console.log('Could not collect form data:', error);
        }
      }

      // Get uploaded files info
      let filesInfo = [];
      if (window.fileUploader) {
        try {
          const files = window.fileUploader.getUploadedFiles();
          filesInfo = files.map(f => ({
            name: f.name,
            type: f.type,
            size: f.size
          }));
        } catch (error) {
          console.log('Could not get files info:', error);
        }
      }

      // Hebbia-inspired agent coordination: Analyze query type and route accordingly
      const queryAnalysis = this.analyzeQueryType(message);
      console.log('🔍 Query analysis:', queryAnalysis);

      // Prepare optimized context for AI (Hebbia-style: send only relevant data)
      const context = {
        message: message,
        queryType: queryAnalysis.type,
        priority: queryAnalysis.priority,
        excelContext: excelContext,
        excelContextInfo: excelContext ? {
          worksheet: excelContext.worksheetName,
          hasData: excelContext.hasActualData,
          dataPoints: excelContext.totalNonEmptyCells,
          sampleData: excelContext.sampleData,
          contextLoadedAt: this.contextLastUpdated
        } : null,
        formData: queryAnalysis.needsFormData ? formData : {},
        uploadedFiles: queryAnalysis.needsFiles ? filesInfo : [],
        chatHistory: this.chatMessages.slice(-3), // Reduced for faster processing
        systemHint: this.generateSystemHint(queryAnalysis, excelContext)
      };

      // Route to multi-agent system for professional M&A analysis
      const response = await this.processWithMultiAgentSystem(message, context);
      
      // Return formatted response for existing chat system
      return response;
      
    } catch (error) {
      console.error('Error processing with multi-agent AI:', error);
      return 'I apologize, but I encountered an error while processing your request. Please try again or contact support if the issue persists.';
    }
  }

  /**
   * Process message with multi-agent system (Stage 2 implementation)
   */
  async processWithMultiAgentSystem(message, context) {
    console.log('🎭 Processing with multi-agent system:', message);
    
    try {
      // Check if multi-agent processor is available
      if (typeof window.multiAgentProcessor === 'undefined') {
        console.warn('⚠️ Multi-agent processor not available, falling back to standard processing');
        return await this.routeToSpecializedAgent(context);
      }
      
      // Process with multi-agent system
      const result = await window.multiAgentProcessor.processQuery(message, {
        excelContext: context.excelContext,
        formData: context.formData,
        chatHistory: context.chatHistory
      });
      
      // Show completion status with results
      if (window.enhancedStatusIndicators && result.metadata) {
        window.enhancedStatusIndicators.showCompletion(result, result.metadata.processingTime);
      }
      
      // Enhance response formatting for display
      return this.enhanceResponseFormatting(result.response);
      
    } catch (error) {
      console.error('❌ Multi-agent processing failed:', error);
      
      // Show error status
      if (window.enhancedStatusIndicators) {
        window.enhancedStatusIndicators.showError(error, 'fallback');
      }
      
      // Fallback to original processing
      console.log('🔄 Falling back to standard processing...');
      return await this.routeToSpecializedAgent(context);
    }
  }

  async callChatAPI(context) {
    console.log('Calling chat API with context:', context);
    
    try {
      // Check if we're running locally or on Netlify
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
      
      return data.response || 'I received your message but couldn\'t generate a proper response.';
      
    } catch (error) {
      console.error('Chat API call failed:', error);
      
      // Fallback to mock response for development
      return this.generateMockResponse(context.message);
    }
  }

  generateMockResponse(message) {
    console.log('Generating mock response for:', message);
    
    const lowerMessage = message.toLowerCase();
    
    // Simple keyword-based responses for development
    if (lowerMessage.includes('revenue') || lowerMessage.includes('income')) {
      return 'To add revenue items, use the "Revenue Items" section below. You can specify different revenue streams and their growth patterns.';
    }
    
    if (lowerMessage.includes('cost') || lowerMessage.includes('expense')) {
      return 'You can add operating expenses and capital expenses in their respective sections. Make sure to include all major cost categories for accurate modeling.';
    }
    
    if (lowerMessage.includes('debt') || lowerMessage.includes('financing')) {
      return 'Configure debt financing in the "Debt Model" section. You can set the LTV ratio and interest rate parameters there.';
    }
    
    if (lowerMessage.includes('generate') || lowerMessage.includes('model') || lowerMessage.includes('excel')) {
      return 'Once you\'ve filled in all the required fields, click "Generate Model in Excel" to create your financial model with P&L and cash flow statements.';
    }
    
    if (lowerMessage.includes('upload') || lowerMessage.includes('file')) {
      return 'You can upload financial documents (PDF, CSV, PNG, JPG) using the file upload area at the top. Then click "Auto Fill with AI" to extract data automatically.';
    }
    
    if (lowerMessage.includes('save') || lowerMessage.includes('load')) {
      return 'Use the "Save Inputs" button to save your current data, and "Load Inputs" to restore previously saved data.';
    }
    
    if (lowerMessage.includes('help') || lowerMessage.includes('how')) {
      return 'I can help you with M&A financial modeling. You can ask me about revenue items, costs, debt financing, generating models, or uploading files for auto-fill.';
    }
    
    // Default response
    return 'I understand you\'re working on your M&A financial model. Could you be more specific about what you need help with? I can assist with revenue items, costs, debt modeling, file uploads, or generating the Excel model.';
  }

  addChatMessage(role, content) {
    console.log(`${role.toUpperCase()}: ${content}`);
    this.chatMessages.push({ 
      role, 
      content, 
      timestamp: new Date().toISOString() 
    });
    
    // Update chat display if chat interface exists
    this.updateChatDisplay();
    
    // Log to console for development
    const prefix = role === 'user' ? '👤 User:' : '🤖 Assistant:';
    console.log(`${prefix} ${content}`);
  }

  updateChatDisplay() {
    // Look for chat display elements
    const chatMessages = document.getElementById('chatMessages');
    const chatContainer = document.getElementById('chatContainer');
    
    if (!chatMessages && !chatContainer) {
      // No chat interface, just log
      return;
    }

    const displayElement = chatMessages || chatContainer;
    
    // Clear existing messages
    displayElement.innerHTML = '';
    
    // Add all messages
    this.chatMessages.forEach(msg => {
      const messageDiv = document.createElement('div');
      messageDiv.className = `chat-message ${msg.role}`;
      
      const roleSpan = document.createElement('span');
      roleSpan.className = 'chat-role';
      roleSpan.textContent = msg.role === 'user' ? 'You:' : 'Assistant:';
      
      const contentDiv = document.createElement('div');
      contentDiv.className = 'chat-content';
      contentDiv.textContent = msg.content;
      
      messageDiv.appendChild(roleSpan);
      messageDiv.appendChild(contentDiv);
      displayElement.appendChild(messageDiv);
    });
    
    // Scroll to bottom
    displayElement.scrollTop = displayElement.scrollHeight;
  }

  // LEGACY METHOD - No longer used with LangChain orchestrator
  showLoading(show) {
    // Look for loading indicators
    const loadingElement = document.getElementById('loading') || 
                          document.getElementById('chatLoading') ||
                          document.querySelector('.loading');
    
    if (loadingElement) {
      loadingElement.style.display = show ? 'block' : 'none';
    }
    
    // Disable send button while loading
    const sendBtn = document.getElementById('sendChatBtn');
    if (sendBtn) {
      sendBtn.disabled = show;
      if (show) {
        sendBtn.textContent = 'Processing...';
      } else {
        sendBtn.textContent = 'Send';
      }
    }
    
    console.log(show ? 'Showing loading indicator' : 'Hiding loading indicator');
  }

  showStatus(message) {
    console.log('Status:', message);
    
    // Look for status display element
    const statusElement = document.getElementById('status') || 
                         document.getElementById('chatStatus') ||
                         document.querySelector('.status');
    
    if (statusElement) {
      statusElement.textContent = message;
      
      // Auto-hide after 3 seconds
      setTimeout(() => {
        statusElement.textContent = '';
      }, 3000);
    }
  }

  async getExcelContext() {
    try {
      if (typeof Excel === 'undefined') {
        return JSON.stringify({ error: 'Excel not available' });
      }

      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const selectedRange = context.workbook.getSelectedRange();
        const usedRange = worksheet.getUsedRange();
        
        worksheet.load(['name']);
        selectedRange.load(['address', 'values', 'formulas']);
        usedRange.load(['address', 'values', 'formulas']);
        
        await context.sync();
        
        return JSON.stringify({
          worksheetName: worksheet.name,
          selectedRange: {
            address: selectedRange.address,
            values: selectedRange.values,
            formulas: selectedRange.formulas
          },
          usedRange: {
            address: usedRange.address,
            values: usedRange.values.slice(0, 10), // Limit to first 10 rows
            formulas: usedRange.formulas.slice(0, 10)
          }
        });
      });
    } catch (error) {
      console.error('Error getting Excel context:', error);
      return JSON.stringify({ error: 'Could not read Excel context' });
    }
  }

  clearChatHistory() {
    this.chatMessages = [];
    this.updateChatDisplay();
    console.log('Chat history cleared');
  }

  getChatHistory() {
    return this.chatMessages;
  }

  exportChatHistory() {
    const chatData = {
      messages: this.chatMessages,
      exportedAt: new Date().toISOString(),
      totalMessages: this.chatMessages.length
    };
    
    const blob = new Blob([JSON.stringify(chatData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = `chat-history-${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    
    URL.revokeObjectURL(url);
    console.log('Chat history exported');
  }

  // Method to handle specific AI-powered tasks
  async processAITask(taskType, data) {
    console.log('Processing AI task:', taskType, data);
    
    switch (taskType) {
      case 'analyze_files':
        return await this.analyzeUploadedFiles(data.files);
      case 'validate_model':
        return await this.validateFinancialModel(data.modelData);
      case 'suggest_improvements':
        return await this.suggestModelImprovements(data.modelData);
      default:
        return 'Unknown task type';
    }
  }

  async analyzeUploadedFiles(files) {
    // Mock analysis - in reality would call AI service
    const analysis = files.map(file => ({
      name: file.name,
      type: file.type,
      extractedData: 'Sample extracted data',
      confidence: 0.85
    }));
    
    return `Analyzed ${files.length} files. Found potential revenue streams and cost structures.`;
  }

  async validateFinancialModel(modelData) {
    // Mock validation - in reality would call AI service
    const issues = [];
    
    if (!modelData.dealName) issues.push('Deal name is missing');
    if (!modelData.revenueItems || modelData.revenueItems.length === 0) issues.push('No revenue items defined');
    if (!modelData.operatingExpenses || modelData.operatingExpenses.length === 0) issues.push('No operating expenses defined');
    
    if (issues.length === 0) {
      return 'Your financial model looks complete and ready for generation!';
    } else {
      return `Please address these issues: ${issues.join(', ')}`;
    }
  }

  async suggestModelImprovements(modelData) {
    // Mock suggestions - in reality would call AI service
    const suggestions = [
      'Consider adding working capital changes to your cash flow model',
      'Include sensitivity analysis for key assumptions',
      'Add debt service calculations if using leverage',
      'Consider seasonal variations in revenue'
    ];
    
    return `Here are some suggestions to improve your model: ${suggestions.join('. ')}.`;
  }

  // Hebbia-inspired multi-agent methods

  /**
   * Analyze query type for intelligent routing (Hebbia's Decomposition Agent)
   */
  analyzeQueryType(message) {
    const lowerMessage = message.toLowerCase();
    
    // Financial Analysis Queries (High Priority)
    if (lowerMessage.includes('moic') || lowerMessage.includes('multiple') || 
        lowerMessage.includes('irr') || lowerMessage.includes('return')) {
      return {
        type: 'financial_analysis',
        priority: 'high',
        needsFormData: false,
        needsFiles: false,
        specialization: 'financial_metrics'
      };
    }

    // Excel Structure Queries
    if (lowerMessage.includes('formula') || lowerMessage.includes('calculation') ||
        lowerMessage.includes('cell') || lowerMessage.includes('range')) {
      return {
        type: 'excel_structure',
        priority: 'high',
        needsFormData: false,
        needsFiles: false,
        specialization: 'excel_formulas'
      };
    }

    // Data Validation Queries
    if (lowerMessage.includes('error') || lowerMessage.includes('wrong') ||
        lowerMessage.includes('validate') || lowerMessage.includes('check')) {
      return {
        type: 'data_validation',
        priority: 'medium',
        needsFormData: true,
        needsFiles: false,
        specialization: 'validation'
      };
    }

    // Form/Input Queries
    if (lowerMessage.includes('revenue') || lowerMessage.includes('cost') ||
        lowerMessage.includes('expense') || lowerMessage.includes('input')) {
      return {
        type: 'form_assistance',
        priority: 'medium',
        needsFormData: true,
        needsFiles: false,
        specialization: 'form_guidance'
      };
    }

    // File Upload Queries
    if (lowerMessage.includes('upload') || lowerMessage.includes('file') ||
        lowerMessage.includes('extract') || lowerMessage.includes('autofill')) {
      return {
        type: 'file_processing',
        priority: 'medium',
        needsFormData: false,
        needsFiles: true,
        specialization: 'data_extraction'
      };
    }

    // General/Conversational
    return {
      type: 'general',
      priority: 'low',
      needsFormData: false,
      needsFiles: false,
      specialization: 'conversation'
    };
  }

  /**
   * Generate optimized system hints (Hebbia's Meta-Prompting Agent)
   */
  generateSystemHint(queryAnalysis, excelContext) {
    const baseHint = "You are an expert M&A financial modeling assistant.";
    
    switch (queryAnalysis.type) {
      case 'financial_analysis':
        return `${baseHint} You specialize in analyzing financial metrics like MOIC, IRR, and cash flows. 
                Provide specific, data-driven insights with exact cell references and calculations. 
                Current Excel structure: ${excelContext?.structure || 'unknown'}.
                ${excelContext?.summary || ''}`;

      case 'excel_structure':
        return `${baseHint} You specialize in Excel formulas and cell relationships. 
                Provide detailed formula explanations and suggest optimizations.
                Focus on calculation logic and dependencies.`;

      case 'data_validation':
        return `${baseHint} You specialize in data validation and error detection. 
                Identify inconsistencies and provide actionable fixes.
                Be precise about what's wrong and how to correct it.`;

      case 'form_assistance':
        return `${baseHint} You specialize in guiding users through M&A model inputs. 
                Provide clear guidance on what data to enter and why.
                Reference industry standards and best practices.`;

      case 'file_processing':
        return `${baseHint} You specialize in document analysis and data extraction. 
                Help users understand what data was extracted and how to verify it.
                Suggest additional data that might be needed.`;

      default:
        return `${baseHint} Provide conversational, helpful responses about M&A financial modeling. 
                Keep responses concise and actionable.`;
    }
  }

  /**
   * Route to specialized processing (Hebbia's Multi-Agent Orchestrator)
   */
  async routeToSpecializedAgent(context) {
    console.log(`🤖 Routing to ${context.queryType} specialist agent...`);

    switch (context.queryType) {
      case 'financial_analysis':
        return await this.processFinancialAnalysis(context);
      
      case 'excel_structure':
        return await this.processExcelStructure(context);
      
      case 'data_validation':
        return await this.processDataValidation(context);
        
      default:
        return await this.callChatAPI(context);
    }
  }

  /**
   * Specialized Financial Analysis Agent (Hebbia-style)
   */
  async processFinancialAnalysis(context) {
    console.log('💰 Financial Analysis Agent processing query...');
    
    // Pre-process financial metrics if available
    if (context.excelContext?.financialMetrics) {
      const metrics = context.excelContext.financialMetrics;
      
      // Add specific financial context to message
      let enhancedMessage = context.message;
      
      if (metrics.moic) {
        enhancedMessage += `\n\nCurrent MOIC: ${metrics.moic.value} (${metrics.moic.interpretation}) at ${metrics.moic.location}`;
      }
      
      if (metrics.irr) {
        enhancedMessage += `\nCurrent IRR: ${metrics.irr.value} (${metrics.irr.interpretation}) at ${metrics.irr.location}`;
      }

      // Create enhanced context for financial analysis
      const enhancedContext = {
        ...context,
        message: enhancedMessage,
        batchType: 'financial_analysis',
        temperature: 0.3 // Lower temperature for more precise financial analysis
      };
      
      return await this.callChatAPI(enhancedContext);
    }
    
    // Fallback to regular processing if no financial metrics available
    return await this.callChatAPI(context);
  }

  /**
   * Specialized Excel Structure Agent
   */
  async processExcelStructure(context) {
    console.log('📊 Excel Structure Agent processing query...');
    
    if (context.excelContext?.keyCalculations) {
      // Add calculation details to context
      const calculations = Object.entries(context.excelContext.keyCalculations)
        .slice(0, 5) // Limit to top 5 calculations for performance
        .map(([cell, data]) => `${cell}: ${data.formula}`)
        .join('\n');
      
      const enhancedContext = {
        ...context,
        message: `${context.message}\n\nKey Calculations:\n${calculations}`,
        temperature: 0.2 // Very precise for formula work
      };
      
      return await this.callChatAPI(enhancedContext);
    }
    
    return await this.callChatAPI(context);
  }

  /**
   * Specialized Data Validation Agent
   */
  async processDataValidation(context) {
    console.log('✅ Data Validation Agent processing query...');
    
    // Include both Excel context and form data for comprehensive validation
    const enhancedContext = {
      ...context,
      systemPrompt: `You are a data validation specialist. Analyze the provided Excel data and form inputs for inconsistencies, errors, or missing critical information. Provide specific, actionable feedback.`,
      temperature: 0.1 // Very low temperature for precise validation
    };
    
    return await this.callChatAPI(enhancedContext);
  }

  // Enhanced UI Methods for Modern Chat Interface

  /**
   * Show live search indicators (like screenshot)
   */
  // LEGACY METHOD - No longer used with LangChain orchestrator  
  showLiveSearchIndicators(message) {
    const chatMessages = document.getElementById('chatMessages') || 
                         document.getElementById('chatContainer');
    
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
    } else if (lowerMessage.includes('formula') || lowerMessage.includes('calculation')) {
      steps.push(
        { icon: '🔍', text: 'Search "Excel formulas"', highlight: null },
        { icon: '🔢', text: 'Analyzing calculations', highlight: null },
        { icon: '👁️', text: 'Looking up dependencies', highlight: null }
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
   * LEGACY METHOD - No longer used with LangChain orchestrator
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

  /**
   * Add formatted chat message with enhanced styling
   */
  addFormattedChatMessage(role, content) {
    console.log(`${role.toUpperCase()}: ${content}`);
    
    // Process content for enhanced formatting
    const formattedContent = this.enhanceResponseFormatting(content);
    
    // Instead of using our internal system, directly integrate with the existing chat
    this.addFormattedMessageToExistingChat(role, formattedContent);
    
    this.chatMessages.push({ 
      role, 
      content: formattedContent, 
      timestamp: new Date().toISOString() 
    });
    
    // Log to console for development
    const prefix = role === 'user' ? '👤 User:' : '🤖 Assistant:';
    console.log(`${prefix} ${content}`);
  }

  /**
   * Integrate with existing chat system in taskpane.html
   */
  addFormattedMessageToExistingChat(role, content) {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;

    // Remove welcome message if it exists
    const welcomeMsg = document.getElementById('chatWelcome');
    if (welcomeMsg) {
      welcomeMsg.style.display = 'none';
    }

    // Create message element using the existing chat structure
    const messageDiv = document.createElement('div');
    messageDiv.className = `chat-message ${role}-message`;
    
    if (role === 'user') {
      messageDiv.innerHTML = `
        <div class="message-content">
          <div class="message-text">${this.escapeHtml(content)}</div>
        </div>
      `;
    } else {
      messageDiv.innerHTML = `
        <div class="message-content">
          <div class="ai-avatar">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M9 12l2 2 4-4"/>
            </svg>
          </div>
          <div class="message-text formatted-response">${content}</div>
        </div>
      `;
    }

    chatMessages.appendChild(messageDiv);
    
    // Scroll to bottom
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }

  /**
   * Enhanced response formatting - convert markdown to modern chat styling
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
   * Update chat display with enhanced formatting
   */
  updateChatDisplay() {
    // Look for chat display elements
    const chatMessages = document.getElementById('chatMessages');
    const chatContainer = document.getElementById('chatContainer');
    
    if (!chatMessages && !chatContainer) {
      // No chat interface, just log
      return;
    }

    const displayElement = chatMessages || chatContainer;
    
    // Clear existing messages
    displayElement.innerHTML = '';
    
    // Add all messages with enhanced formatting
    this.chatMessages.forEach(msg => {
      const messageDiv = document.createElement('div');
      messageDiv.className = `chat-message ${msg.role}`;
      
      if (msg.role === 'user') {
        messageDiv.innerHTML = `
          <div class="message-content user-message">
            ${this.escapeHtml(msg.content)}
          </div>
        `;
      } else {
        messageDiv.innerHTML = `
          <div class="message-content assistant-message">
            ${msg.content} <!-- Already formatted HTML -->
          </div>
        `;
      }
      
      displayElement.appendChild(messageDiv);
    });
    
    // Scroll to bottom
    displayElement.scrollTop = displayElement.scrollHeight;
  }

  /**
   * Escape HTML for user messages
   */
  escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }
}

// Export for use in main application
window.ChatHandler = ChatHandler;