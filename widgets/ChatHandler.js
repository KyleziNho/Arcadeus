class ChatHandler {
  constructor() {
    this.chatMessages = [];
    this.isProcessing = false;
    this.excelAnalyzer = new ExcelLiveAnalyzer();
  }

  async initialize() {
    console.log('Initializing Hebbia-inspired chat handler...');
    
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

    console.log('Sending chat message:', message);
    
    try {
      this.isProcessing = true;
      this.showLoading(true);
      
      // Add user message to chat
      this.addChatMessage('user', message);
      
      // Clear input
      chatInput.value = '';
      
      // Process message with AI
      const response = await this.processWithAI(message);
      
      // Add assistant response
      this.addChatMessage('assistant', response);
      
    } catch (error) {
      console.error('Error sending chat message:', error);
      this.addChatMessage('assistant', 'Sorry, I encountered an error. Please try again.');
    } finally {
      this.isProcessing = false;
      this.showLoading(false);
    }
  }

  async processWithAI(message) {
    console.log('🎯 Processing message with Hebbia-inspired multi-agent approach:', message);
    
    try {
      // Use ExcelLiveAnalyzer for comprehensive, fast context
      let excelContext = null;
      try {
        if (this.excelAnalyzer) {
          console.log('📊 Getting optimized Excel context for AI...');
          excelContext = await this.excelAnalyzer.getOptimizedContextForAI();
          console.log('✅ Excel context retrieved:', {
            structure: excelContext?.structure,
            hasMetrics: !!excelContext?.financialMetrics,
            hasSummary: !!excelContext?.summary
          });
        }
      } catch (error) {
        console.log('Could not get comprehensive Excel context:', error);
        // Fallback to basic context
        try {
          excelContext = await this.getExcelContext();
        } catch (fallbackError) {
          console.log('Fallback context also failed:', fallbackError);
        }
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
        formData: queryAnalysis.needsFormData ? formData : {},
        uploadedFiles: queryAnalysis.needsFiles ? filesInfo : [],
        chatHistory: this.chatMessages.slice(-3), // Reduced for faster processing
        systemHint: this.generateSystemHint(queryAnalysis, excelContext)
      };

      // Route to appropriate processing based on query type
      const response = await this.routeToSpecializedAgent(context);
      
      return response;
      
    } catch (error) {
      console.error('Error processing with multi-agent AI:', error);
      return 'I apologize, but I encountered an error while processing your request. Please try again or contact support if the issue persists.';
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
}

// Export for use in main application
window.ChatHandler = ChatHandler;