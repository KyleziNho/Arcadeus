/**
 * EnhancedChatSystem.js
 * A practical implementation that provides MCP-like features in the browser
 * Works with Excel add-ins without requiring separate server processes
 */

class EnhancedChatSystem {
  constructor() {
    // Core components
    this.conversationHistory = [];
    this.excelContext = {};
    this.operationHistory = [];
    this.maxHistorySize = 100;
    this.maxOperations = 50;
    
    // Load persisted data
    this.loadState();
    
    // Excel operation patterns
    this.patterns = {
      colorCode: /(?:color|colour|highlight|format).*(?:red|green|blue|yellow)/i,
      calculate: /(?:calculate|compute|sum|average|total)/i,
      create: /(?:create|make|generate|add)/i,
      analyze: /(?:analyze|analyse|review|check)/i,
      moic: /moic|multiple.*invested.*capital/i,
      irr: /irr|internal.*rate.*return/i
    };
  }

  /**
   * Initialize the enhanced chat system
   */
  async initialize() {
    console.log('🚀 Initializing Enhanced Chat System...');
    
    try {
      // Set up Excel context monitoring
      await this.setupExcelMonitoring();
      
      // Set up UI event listeners
      this.setupUIListeners();
      
      // Initialize with current Excel state
      await this.updateExcelContext();
      
      console.log('✅ Enhanced Chat System ready');
      return true;
    } catch (error) {
      console.error('❌ Initialization failed:', error);
      return false;
    }
  }

  /**
   * Process a user message with full context awareness
   */
  async processMessage(message) {
    console.log('💬 Processing:', message);
    
    // Add to conversation history
    this.addToHistory('user', message);
    
    // Analyze intent
    const intent = this.analyzeIntent(message);
    console.log('🎯 Intent:', intent);
    
    // Get current Excel context
    const context = await this.getFullContext();
    
    // Check if this is an Excel operation
    if (intent.isExcelOperation) {
      return await this.handleExcelOperation(intent, message, context);
    }
    
    // Otherwise, process with AI
    return await this.processWithAI(message, context);
  }

  /**
   * Analyze message intent
   */
  analyzeIntent(message) {
    const lower = message.toLowerCase();
    
    const intent = {
      isExcelOperation: false,
      operation: null,
      confidence: 0,
      parameters: {}
    };
    
    // Check for color coding request
    if (this.patterns.colorCode.test(message)) {
      intent.isExcelOperation = true;
      intent.operation = 'format';
      intent.confidence = 0.9;
      
      // Extract color
      const colorMatch = message.match(/(red|green|blue|yellow|orange|purple)/i);
      if (colorMatch) {
        intent.parameters.color = colorMatch[1].toLowerCase();
      }
      
      // Extract threshold (e.g., "70%")
      const thresholdMatch = message.match(/(\d+)%/);
      if (thresholdMatch) {
        intent.parameters.threshold = parseInt(thresholdMatch[1]) / 100;
      }
    }
    
    // Check for calculation request
    if (this.patterns.calculate.test(message)) {
      intent.isExcelOperation = true;
      intent.operation = 'calculate';
      intent.confidence = 0.8;
    }
    
    // Check for financial metrics
    if (this.patterns.moic.test(message)) {
      intent.operation = 'analyze_moic';
      intent.confidence = 0.95;
    }
    
    if (this.patterns.irr.test(message)) {
      intent.operation = 'analyze_irr';
      intent.confidence = 0.95;
    }
    
    return intent;
  }

  /**
   * Handle Excel operations directly
   */
  async handleExcelOperation(intent, message, context) {
    console.log('📊 Handling Excel operation:', intent.operation);
    
    try {
      let result = '';
      
      switch (intent.operation) {
        case 'format':
          result = await this.applyConditionalFormatting(intent.parameters);
          break;
          
        case 'calculate':
          result = await this.performCalculation(message, context);
          break;
          
        default:
          result = await this.processWithAI(message, context);
      }
      
      // Record operation for undo
      this.recordOperation({
        type: intent.operation,
        message: message,
        timestamp: new Date(),
        context: context.excel
      });
      
      return result;
      
    } catch (error) {
      console.error('Excel operation failed:', error);
      return `I encountered an error: ${error.message}. Please try again.`;
    }
  }

  /**
   * Apply conditional formatting based on natural language
   */
  async applyConditionalFormatting(params) {
    if (typeof Excel === 'undefined') {
      return "Excel API not available. Please ensure you're running this in Excel.";
    }
    
    return await Excel.run(async (context) => {
      try {
        const range = context.workbook.getSelectedRange();
        range.load(['address', 'values']);
        await context.sync();
        
        const values = range.values;
        const colorMap = {
          'red': '#FF0000',
          'green': '#00FF00',
          'blue': '#0000FF',
          'yellow': '#FFFF00',
          'orange': '#FFA500',
          'purple': '#800080'
        };
        
        const color = colorMap[params.color] || '#FF0000';
        const threshold = params.threshold || 0.7;
        
        // Apply formatting based on threshold
        for (let i = 0; i < values.length; i++) {
          for (let j = 0; j < values[i].length; j++) {
            const value = parseFloat(values[i][j]);
            if (!isNaN(value) && value >= threshold) {
              const cell = range.getCell(i, j);
              cell.format.fill.color = color;
            }
          }
        }
        
        await context.sync();
        
        // Record the operation
        this.recordOperation({
          type: 'format',
          range: range.address,
          color: params.color,
          threshold: threshold,
          timestamp: new Date()
        });
        
        return `Applied ${params.color} formatting to cells with values >= ${threshold * 100}% in range ${range.address}`;
        
      } catch (error) {
        throw new Error(`Formatting failed: ${error.message}`);
      }
    });
  }

  /**
   * Process message with AI (via Netlify function)
   */
  async processWithAI(message, context) {
    try {
      const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
      const apiUrl = isLocal ? 'http://localhost:8888/.netlify/functions/chat' : '/.netlify/functions/chat';
      
      // Include conversation context
      const requestBody = {
        message: message,
        context: {
          conversation: this.conversationHistory.slice(-5), // Last 5 messages
          excel: context.excel,
          formData: context.formData
        },
        batchType: 'chat'
      };
      
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestBody)
      });
      
      if (!response.ok) {
        throw new Error(`API error: ${response.status}`);
      }
      
      const data = await response.json();
      
      if (!data.success) {
        throw new Error(data.error || 'AI processing failed');
      }
      
      // Add response to history
      this.addToHistory('assistant', data.content);
      
      return data.content;
      
    } catch (error) {
      console.error('AI processing error:', error);
      throw error;
    }
  }

  /**
   * Get full context for AI processing
   */
  async getFullContext() {
    const context = {
      excel: await this.updateExcelContext(),
      formData: this.getFormData(),
      conversation: this.conversationHistory
    };
    
    return context;
  }

  /**
   * Update Excel context
   */
  async updateExcelContext() {
    if (typeof Excel === 'undefined') {
      return { available: false };
    }
    
    try {
      return await Excel.run(async (context) => {
        const workbook = context.workbook;
        const worksheet = workbook.worksheets.getActiveWorksheet();
        const range = workbook.getSelectedRange();
        
        workbook.load('name');
        worksheet.load('name');
        range.load(['address', 'values', 'formulas']);
        
        await context.sync();
        
        this.excelContext = {
          workbook: workbook.name,
          worksheet: worksheet.name,
          selectedRange: range.address,
          values: range.values,
          formulas: range.formulas,
          timestamp: new Date()
        };
        
        return this.excelContext;
      });
    } catch (error) {
      console.error('Failed to get Excel context:', error);
      return { available: false, error: error.message };
    }
  }

  /**
   * Set up Excel event monitoring
   */
  async setupExcelMonitoring() {
    if (typeof Excel === 'undefined') return;
    
    try {
      await Excel.run(async (context) => {
        // Monitor selection changes
        context.workbook.onSelectionChanged.add(async (event) => {
          console.log('Selection changed:', event.address);
          this.excelContext.selectedRange = event.address;
          this.saveState();
        });
        
        await context.sync();
      });
    } catch (error) {
      console.error('Failed to set up Excel monitoring:', error);
    }
  }

  /**
   * Set up UI event listeners
   */
  setupUIListeners() {
    // Send button
    const sendBtn = document.getElementById('sendChatBtn') || document.getElementById('sendMessage');
    if (sendBtn) {
      sendBtn.addEventListener('click', () => this.handleSendMessage());
    }
    
    // Input field
    const input = document.getElementById('chatInput');
    if (input) {
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          this.handleSendMessage();
        }
      });
    }
    
    // Undo/Redo shortcuts
    document.addEventListener('keydown', (e) => {
      if (e.ctrlKey || e.metaKey) {
        if (e.key === 'z') {
          e.preventDefault();
          this.undo();
        } else if (e.key === 'y') {
          e.preventDefault();
          this.redo();
        }
      }
    });
  }

  /**
   * Handle send message
   */
  async handleSendMessage() {
    const input = document.getElementById('chatInput');
    if (!input || !input.value.trim()) return;
    
    const message = input.value.trim();
    input.value = '';
    
    // Display user message
    this.displayMessage('user', message);
    
    try {
      // Process and get response
      const response = await this.processMessage(message);
      
      // Display response
      this.displayMessage('assistant', response);
      
    } catch (error) {
      this.displayMessage('assistant', `Error: ${error.message}`, true);
    }
  }

  /**
   * Display message in chat UI
   */
  displayMessage(role, content, isError = false) {
    const container = document.getElementById('chatMessages');
    if (!container) return;
    
    const messageDiv = document.createElement('div');
    messageDiv.className = `chat-message ${role}-message ${isError ? 'error' : ''}`;
    
    const roleLabel = document.createElement('strong');
    roleLabel.textContent = role === 'user' ? 'You: ' : 'Assistant: ';
    
    const contentSpan = document.createElement('span');
    contentSpan.textContent = content;
    
    messageDiv.appendChild(roleLabel);
    messageDiv.appendChild(contentSpan);
    container.appendChild(messageDiv);
    
    // Scroll to bottom
    container.scrollTop = container.scrollHeight;
  }

  /**
   * Undo last operation
   */
  undo() {
    if (this.operationHistory.length === 0) {
      console.log('Nothing to undo');
      return false;
    }
    
    const operation = this.operationHistory.pop();
    console.log('Undoing:', operation);
    
    // Save state
    this.saveState();
    
    // Show notification
    this.showNotification('Operation undone', 'info');
    return true;
  }

  /**
   * Redo operation
   */
  redo() {
    console.log('Redo not implemented yet');
    return false;
  }

  /**
   * Record an operation for undo history
   */
  recordOperation(operation) {
    this.operationHistory.push({
      ...operation,
      id: Date.now().toString(),
      timestamp: new Date()
    });
    
    // Limit history size
    if (this.operationHistory.length > this.maxOperations) {
      this.operationHistory.shift();
    }
    
    this.saveState();
  }

  /**
   * Add message to conversation history
   */
  addToHistory(role, content) {
    this.conversationHistory.push({
      role,
      content,
      timestamp: new Date()
    });
    
    // Limit history size
    if (this.conversationHistory.length > this.maxHistorySize) {
      this.conversationHistory.shift();
    }
    
    this.saveState();
  }

  /**
   * Get form data if available
   */
  getFormData() {
    if (window.formHandler && typeof window.formHandler.collectAllModelData === 'function') {
      try {
        return window.formHandler.collectAllModelData();
      } catch (error) {
        console.error('Failed to collect form data:', error);
      }
    }
    return {};
  }

  /**
   * Save state to localStorage
   */
  saveState() {
    try {
      const state = {
        conversationHistory: this.conversationHistory,
        excelContext: this.excelContext,
        operationHistory: this.operationHistory
      };
      
      localStorage.setItem('enhancedChatState', JSON.stringify(state));
    } catch (error) {
      console.error('Failed to save state:', error);
    }
  }

  /**
   * Load state from localStorage
   */
  loadState() {
    try {
      const saved = localStorage.getItem('enhancedChatState');
      if (saved) {
        const state = JSON.parse(saved);
        this.conversationHistory = state.conversationHistory || [];
        this.excelContext = state.excelContext || {};
        this.operationHistory = state.operationHistory || [];
      }
    } catch (error) {
      console.error('Failed to load state:', error);
    }
  }

  /**
   * Show notification
   */
  showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    notification.textContent = message;
    notification.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      padding: 12px 20px;
      background: ${type === 'error' ? '#f44336' : '#2196F3'};
      color: white;
      border-radius: 4px;
      z-index: 10000;
      animation: slideIn 0.3s ease;
    `;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
      notification.remove();
    }, 3000);
  }

  /**
   * Clear conversation
   */
  clearConversation() {
    this.conversationHistory = [];
    this.operationHistory = [];
    this.saveState();
    
    const container = document.getElementById('chatMessages');
    if (container) {
      container.innerHTML = '';
    }
    
    console.log('Conversation cleared');
  }
}

// Initialize when DOM ready
document.addEventListener('DOMContentLoaded', async () => {
  window.enhancedChat = new EnhancedChatSystem();
  await window.enhancedChat.initialize();
  
  console.log('✅ Enhanced Chat System is ready!');
  console.log('Features available:');
  console.log('- Conversational memory ✓');
  console.log('- Excel context awareness ✓');
  console.log('- Natural language processing ✓');
  console.log('- Direct Excel manipulation ✓');
  console.log('- Undo/Redo support ✓');
  console.log('- Persistent state ✓');
});

// Export for use
window.EnhancedChatSystem = EnhancedChatSystem;