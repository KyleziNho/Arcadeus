class ChatHandler {
  constructor() {
    this.chatMessages = [];
    this.isProcessing = false;
  }

  initialize() {
    console.log('Initializing chat handler...');
    
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
    
    console.log('✅ Chat handler initialized');
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
    console.log('Processing message with AI:', message);
    
    try {
      // Get current Excel context if available
      let excelContext = '';
      try {
        if (window.excelGenerator) {
          excelContext = await this.getExcelContext();
        }
      } catch (error) {
        console.log('Could not get Excel context:', error);
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

      // Prepare context for AI
      const context = {
        message: message,
        excelContext: excelContext,
        formData: formData,
        uploadedFiles: filesInfo,
        chatHistory: this.chatMessages.slice(-5) // Last 5 messages for context
      };

      // Call chat API
      const response = await this.callChatAPI(context);
      
      return response;
      
    } catch (error) {
      console.error('Error processing with AI:', error);
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
}

// Export for use in main application
window.ChatHandler = ChatHandler;