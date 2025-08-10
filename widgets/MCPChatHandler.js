/**
 * MCPChatHandler.js
 * Enhanced chat handler with full MCP integration
 * Replaces the basic ChatHandler with comprehensive MCP functionality
 */

// Import MCP components (Note: In production, these would be proper ES6 imports)
// For now, we'll assume they're available as global objects or loaded via script tags

class MCPChatHandler {
  constructor() {
    this.mcpClient = null;
    this.eventMonitor = null;
    this.notificationManager = null;
    this.isInitialized = false;
    this.isProcessing = false;
    this.currentModel = 'claude-sonnet';
    this.settings = this.loadSettings();
  }

  /**
   * Initialize the enhanced MCP chat system
   */
  async initialize() {
    console.log('🚀 Initializing MCP Chat Handler...');
    
    try {
      // Initialize MCP components
      await this.initializeMCPComponents();
      
      // Set up UI event listeners
      this.setupUIEventListeners();
      
      // Initialize Excel context monitoring
      await this.setupExcelMonitoring();
      
      // Set up notifications
      this.setupNotifications();
      
      // Load conversation history
      await this.loadConversationHistory();
      
      this.isInitialized = true;
      console.log('✅ MCP Chat Handler initialized successfully');
      
      // Show welcome message if no history
      if (this.mcpClient?.getConversationHistory().length === 0) {
        this.showWelcomeMessage();
      }
      
    } catch (error) {
      console.error('❌ Failed to initialize MCP Chat Handler:', error);
      this.showError('Failed to initialize chat system. Some features may not work properly.');
    }
  }

  /**
   * Initialize MCP components
   */
  async initializeMCPComponents() {
    console.log('🔧 Initializing MCP components...');

    // Initialize MCPChatClient (assumes it's available globally)
    if (typeof MCPChatClient !== 'undefined') {
      this.mcpClient = new MCPChatClient();
      await this.mcpClient.initialize();
    } else {
      console.warn('⚠️ MCPChatClient not available, using fallback mode');
    }

    // Initialize Excel Event Monitor
    if (typeof ExcelEventMonitor !== 'undefined') {
      this.eventMonitor = new ExcelEventMonitor({
        enabledEvents: [
          'worksheet-changed',
          'selection-changed', 
          'format-changed',
          'formula-changed'
        ],
        debounceMs: 200,
        maxEventsPerSecond: 10
      });
      await this.eventMonitor.initialize();
    }

    // Initialize Notification Manager
    if (typeof MCPNotificationManager !== 'undefined') {
      this.notificationManager = new MCPNotificationManager({
        position: 'top-right',
        maxVisible: 3,
        defaultAutoHide: 4000
      });
    }

    console.log('✅ MCP components initialized');
  }

  /**
   * Set up UI event listeners
   */
  setupUIEventListeners() {
    // Send message button
    const sendBtn = document.getElementById('sendMessage') || document.getElementById('sendChatBtn');
    if (sendBtn) {
      sendBtn.addEventListener('click', () => this.sendMessage());
    }

    // Chat input
    const chatInput = document.getElementById('chatInput');
    if (chatInput) {
      chatInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          this.sendMessage();
        }
      });

      // Character count
      chatInput.addEventListener('input', () => {
        const charCount = document.getElementById('charCount');
        if (charCount) {
          charCount.textContent = chatInput.value.length;
        }
      });
    }

    // Model selector
    const modelSelector = document.getElementById('modelSelector');
    if (modelSelector) {
      modelSelector.addEventListener('change', (e) => {
        this.currentModel = e.target.value;
        this.settings.selectedModel = e.target.value;
        this.saveSettings();
      });
    }

    // Undo/Redo buttons
    const undoBtn = document.getElementById('undoBtn');
    const redoBtn = document.getElementById('redoBtn');
    
    if (undoBtn) {
      undoBtn.addEventListener('click', () => this.undo());
    }
    
    if (redoBtn) {
      redoBtn.addEventListener('click', () => this.redo());
    }

    // Clear chat button
    const clearBtn = document.getElementById('clearChatBtn');
    if (clearBtn) {
      clearBtn.addEventListener('click', () => this.clearConversation());
    }

    // Settings button
    const settingsBtn = document.getElementById('settingsBtn');
    if (settingsBtn) {
      settingsBtn.addEventListener('click', () => this.showSettings());
    }

    // Refresh context button
    const refreshBtn = document.getElementById('refreshContextBtn');
    if (refreshBtn) {
      refreshBtn.addEventListener('click', () => this.refreshExcelContext());
    }

    // Global keyboard shortcuts
    document.addEventListener('keydown', (e) => {
      if (e.ctrlKey || e.metaKey) {
        switch (e.key) {
          case 'z':
            e.preventDefault();
            this.undo();
            break;
          case 'y':
            e.preventDefault();
            this.redo();
            break;
        }
      }
    });

    console.log('✅ UI event listeners set up');
  }

  /**
   * Set up Excel monitoring
   */
  async setupExcelMonitoring() {
    if (!this.eventMonitor) return;

    // Subscribe to Excel events
    this.eventMonitor.subscribe(
      ['selection-changed', 'worksheet-changed'],
      async (event) => {
        await this.handleExcelEvent(event);
      },
      { priority: 'normal' }
    );

    // Update context display
    await this.refreshExcelContext();
    
    console.log('✅ Excel monitoring set up');
  }

  /**
   * Set up notifications
   */
  setupNotifications() {
    if (!this.notificationManager) return;

    // Subscribe to MCP notifications
    this.notificationManager.subscribe({
      methods: ['*'],
      categories: ['*'],
      priorities: ['*'],
      callback: async (notification) => {
        console.log('📢 MCP Notification:', notification);
      }
    });

    console.log('✅ Notifications set up');
  }

  /**
   * Send a message through MCP system
   */
  async sendMessage() {
    const chatInput = document.getElementById('chatInput');
    if (!chatInput || this.isProcessing) return;

    const message = chatInput.value.trim();
    if (!message) return;

    console.log('💬 Sending message:', message);

    try {
      this.isProcessing = true;
      this.showTypingIndicator(true);
      
      // Clear input
      chatInput.value = '';
      this.updateCharCount(0);
      
      // Add user message to display
      this.addMessageToDisplay('user', message);

      // Process with MCP client
      let response;
      if (this.mcpClient) {
        response = await this.mcpClient.processMessageWithMCPFeatures(message);
      } else {
        // Fallback to simple processing
        response = await this.processFallback(message);
      }

      // Add assistant response
      this.addMessageToDisplay('assistant', response, {
        model: this.currentModel,
        timestamp: new Date()
      });

      // Update undo/redo buttons
      this.updateUndoRedoButtons();

    } catch (error) {
      console.error('❌ Error processing message:', error);
      this.addMessageToDisplay('assistant', 
        'I encountered an error processing your request. Please try again.', 
        { isError: true }
      );
    } finally {
      this.isProcessing = false;
      this.showTypingIndicator(false);
    }
  }

  /**
   * Add message to chat display
   */
  addMessageToDisplay(role, content, metadata = {}) {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;

    // Remove welcome message if present
    const welcomeMsg = chatMessages.querySelector('.welcome-message');
    if (welcomeMsg && role === 'user') {
      welcomeMsg.remove();
    }

    // Create message element
    const messageDiv = document.createElement('div');
    messageDiv.className = `chat-message ${role}-message`;

    const avatar = this.createAvatar(role);
    const content_div = this.createMessageContent(role, content, metadata);

    messageDiv.appendChild(avatar);
    messageDiv.appendChild(content_div);

    chatMessages.appendChild(messageDiv);
    
    // Scroll to bottom
    chatMessages.scrollTop = chatMessages.scrollHeight;

    // Show notification for assistant messages
    if (role === 'assistant' && this.notificationManager && document.hidden) {
      this.notificationManager.sendNotification({
        title: 'New AI Response',
        message: content.substring(0, 50) + '...',
        category: 'chat',
        priority: 'normal'
      });
    }
  }

  /**
   * Create message avatar
   */
  createAvatar(role) {
    const avatar = document.createElement('div');
    avatar.className = 'message-avatar';
    
    if (role === 'user') {
      avatar.innerHTML = `
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
          <circle cx="12" cy="7" r="4"/>
        </svg>
      `;
    } else {
      avatar.innerHTML = `
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M12 2L2 7l10 5 10-5-10-5z"/>
          <path d="M2 17l10 5 10-5"/>
          <path d="M2 12l10 5 10-5"/>
        </svg>
      `;
    }
    
    return avatar;
  }

  /**
   * Create message content
   */
  createMessageContent(role, content, metadata) {
    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content';

    // Message header
    if (role === 'assistant') {
      const header = document.createElement('div');
      header.className = 'message-header';
      header.innerHTML = `
        <span class="model-name">${metadata.model || this.currentModel}</span>
        ${this.isProcessing ? '<span class="typing-indicator">Typing...</span>' : ''}
      `;
      contentDiv.appendChild(header);
    }

    // Message text
    const textDiv = document.createElement('div');
    textDiv.className = 'message-text';
    textDiv.innerHTML = this.formatMessageContent(content);
    contentDiv.appendChild(textDiv);

    // Message actions for assistant messages
    if (role === 'assistant') {
      const actions = this.createMessageActions(content, metadata);
      if (actions) {
        contentDiv.appendChild(actions);
      }
    }

    return contentDiv;
  }

  /**
   * Format message content (markdown-like formatting)
   */
  formatMessageContent(content) {
    // Basic formatting
    return content
      .replace(/`([^`]+)`/g, '<code>$1</code>')
      .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
      .replace(/\*([^*]+)\*/g, '<em>$1</em>')
      .replace(/\n/g, '<br>');
  }

  /**
   * Create message actions
   */
  createMessageActions(content, metadata) {
    if (metadata.isError) return null;

    const actions = document.createElement('div');
    actions.className = 'message-actions';

    // Copy action
    const copyBtn = document.createElement('button');
    copyBtn.className = 'action-btn';
    copyBtn.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <rect x="9" y="9" width="13" height="13" rx="2" ry="2"/>
        <path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/>
      </svg>
      Copy
    `;
    copyBtn.onclick = () => this.copyToClipboard(content);
    actions.appendChild(copyBtn);

    // Execute action (if content contains Excel operations)
    if (content.includes('=') || content.toLowerCase().includes('excel')) {
      const executeBtn = document.createElement('button');
      executeBtn.className = 'action-btn';
      executeBtn.innerHTML = `
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <polygon points="5 3 19 12 5 21 5 3"/>
        </svg>
        Execute
      `;
      executeBtn.onclick = () => this.executeContent(content);
      actions.appendChild(executeBtn);
    }

    return actions;
  }

  /**
   * Handle Excel events
   */
  async handleExcelEvent(event) {
    console.log('📊 Excel event:', event.type, event);

    // Update context display
    if (event.type === 'selection-changed') {
      this.updateSelectionDisplay(event.data.address);
    }

    // Notify MCP client of context changes
    if (this.mcpClient) {
      this.mcpClient.conversationManager?.updateExcelContext({
        selectedRange: event.data.address,
        worksheet: event.worksheet
      });
    }
  }

  /**
   * Update selection display
   */
  updateSelectionDisplay(address) {
    const selectionIndicator = document.getElementById('selectionIndicator');
    const selectedRange = document.getElementById('selectedRange');
    
    if (selectionIndicator) {
      selectionIndicator.textContent = address || 'No cell selected';
    }
    
    if (selectedRange) {
      selectedRange.textContent = address || 'None';
    }
  }

  /**
   * Refresh Excel context
   */
  async refreshExcelContext() {
    try {
      if (typeof Excel === 'undefined') {
        this.updateContextDisplay({
          workbook: 'Excel not available',
          worksheet: 'N/A',
          selection: 'N/A'
        });
        return;
      }

      await Excel.run(async (context) => {
        const workbook = context.workbook;
        const worksheet = workbook.worksheets.getActiveWorksheet();
        const selectedRange = workbook.getSelectedRange();
        
        workbook.load('name');
        worksheet.load('name');
        selectedRange.load('address');
        
        await context.sync();
        
        this.updateContextDisplay({
          workbook: workbook.name || 'Workbook',
          worksheet: worksheet.name || 'Sheet1',
          selection: selectedRange.address || 'None'
        });
      });
    } catch (error) {
      console.error('Failed to refresh Excel context:', error);
    }
  }

  /**
   * Update context display
   */
  updateContextDisplay(context) {
    const elements = {
      workbookName: context.workbook,
      activeSheet: context.worksheet,
      selectedRange: context.selection
    };

    for (const [id, value] of Object.entries(elements)) {
      const element = document.getElementById(id);
      if (element) {
        element.textContent = value;
      }
    }
  }

  /**
   * Show typing indicator
   */
  showTypingIndicator(show) {
    const sendBtn = document.getElementById('sendMessage') || document.getElementById('sendChatBtn');
    
    if (sendBtn) {
      sendBtn.disabled = show;
      if (show) {
        sendBtn.innerHTML = `
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="12" cy="12" r="3"/>
          </svg>
          <span>Processing...</span>
        `;
      } else {
        sendBtn.innerHTML = `
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <line x1="22" y1="2" x2="11" y2="13"/>
            <polygon points="22 2 15 22 11 13 2 9 22 2"/>
          </svg>
          <span>Send</span>
        `;
      }
    }
  }

  /**
   * Update character count
   */
  updateCharCount(count) {
    const charCount = document.getElementById('charCount');
    if (charCount) {
      charCount.textContent = count;
    }
  }

  /**
   * Undo last operation
   */
  async undo() {
    if (this.mcpClient) {
      const success = await this.mcpClient.undo();
      if (success) {
        this.showNotification('Undo successful', 'success');
      } else {
        this.showNotification('Nothing to undo', 'info');
      }
    }
    this.updateUndoRedoButtons();
  }

  /**
   * Redo last undone operation
   */
  async redo() {
    if (this.mcpClient) {
      const success = await this.mcpClient.redo();
      if (success) {
        this.showNotification('Redo successful', 'success');
      } else {
        this.showNotification('Nothing to redo', 'info');
      }
    }
    this.updateUndoRedoButtons();
  }

  /**
   * Update undo/redo button states
   */
  updateUndoRedoButtons() {
    if (!this.mcpClient) return;

    const undoBtn = document.getElementById('undoBtn');
    const redoBtn = document.getElementById('redoBtn');
    
    if (undoBtn) {
      const canUndo = this.mcpClient.operationHistory?.canUndo();
      undoBtn.disabled = !canUndo;
    }
    
    if (redoBtn) {
      const canRedo = this.mcpClient.operationHistory?.canRedo();
      redoBtn.disabled = !canRedo;
    }
  }

  /**
   * Clear conversation
   */
  clearConversation() {
    if (confirm('Are you sure you want to clear the conversation history?')) {
      if (this.mcpClient) {
        this.mcpClient.clearConversation();
      }
      
      const chatMessages = document.getElementById('chatMessages');
      if (chatMessages) {
        chatMessages.innerHTML = '';
        this.showWelcomeMessage();
      }
    }
  }

  /**
   * Show welcome message
   */
  showWelcomeMessage() {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;

    // The welcome message HTML is already in the template
    // We just need to make sure it's visible
    const welcomeHTML = chatMessages.innerHTML;
    if (!welcomeHTML.includes('welcome-message')) {
      chatMessages.innerHTML = `
        <div class="welcome-message">
          <div class="welcome-icon">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
              <path d="M12 2L2 7l10 5 10-5-10-5z"/>
              <path d="M2 17l10 5 10-5"/>
              <path d="M2 12l10 5 10-5"/>
            </svg>
          </div>
          <h3>Welcome to Arcadeus MCP Assistant</h3>
          <p>I can help you with advanced Excel operations, financial modeling, and data analysis.</p>
        </div>
      `;
    }
  }

  /**
   * Show settings modal
   */
  showSettings() {
    const modal = document.getElementById('chatSettingsModal');
    if (modal) {
      modal.style.display = 'flex';
      
      // Load current settings
      const elements = {
        openaiKey: this.settings.apiKeys?.openai || '',
        anthropicKey: this.settings.apiKeys?.anthropic || '',
        googleKey: this.settings.apiKeys?.google || '',
        autoReadExcel: this.settings.autoReadExcel,
        streamResponses: this.settings.streamResponses,
        maxHistory: this.settings.maxHistory
      };

      for (const [id, value] of Object.entries(elements)) {
        const element = document.getElementById(id);
        if (element) {
          if (element.type === 'checkbox') {
            element.checked = value;
          } else {
            element.value = value;
          }
        }
      }
    }
  }

  /**
   * Save settings
   */
  saveSettings() {
    const settings = {
      apiKeys: {
        openai: document.getElementById('openaiKey')?.value || '',
        anthropic: document.getElementById('anthropicKey')?.value || '',
        google: document.getElementById('googleKey')?.value || ''
      },
      autoReadExcel: document.getElementById('autoReadExcel')?.checked || true,
      streamResponses: document.getElementById('streamResponses')?.checked || true,
      maxHistory: parseInt(document.getElementById('maxHistory')?.value) || 20,
      selectedModel: this.currentModel
    };

    this.settings = settings;
    localStorage.setItem('mcpChatSettings', JSON.stringify(settings));
    
    // Close modal
    const modal = document.getElementById('chatSettingsModal');
    if (modal) {
      modal.style.display = 'none';
    }

    this.showNotification('Settings saved', 'success');
  }

  /**
   * Load settings
   */
  loadSettings() {
    try {
      const saved = localStorage.getItem('mcpChatSettings');
      return saved ? JSON.parse(saved) : {
        autoReadExcel: true,
        streamResponses: true,
        maxHistory: 20,
        selectedModel: 'claude-sonnet',
        apiKeys: {}
      };
    } catch (error) {
      console.error('Failed to load settings:', error);
      return {
        autoReadExcel: true,
        streamResponses: true,
        maxHistory: 20,
        selectedModel: 'claude-sonnet',
        apiKeys: {}
      };
    }
  }

  /**
   * Load conversation history
   */
  async loadConversationHistory() {
    if (!this.mcpClient) return;

    try {
      const history = this.mcpClient.getConversationHistory();
      
      // Display last few messages
      const recentMessages = history.slice(-10);
      for (const msg of recentMessages) {
        this.addMessageToDisplay(msg.role, msg.content, {
          timestamp: new Date(msg.timestamp)
        });
      }
      
    } catch (error) {
      console.error('Failed to load conversation history:', error);
    }
  }

  /**
   * Copy content to clipboard
   */
  async copyToClipboard(content) {
    try {
      await navigator.clipboard.writeText(content);
      this.showNotification('Copied to clipboard', 'success');
    } catch (error) {
      console.error('Failed to copy:', error);
      this.showNotification('Failed to copy', 'error');
    }
  }

  /**
   * Execute content in Excel
   */
  async executeContent(content) {
    if (!this.mcpClient) {
      this.showNotification('MCP not available for execution', 'error');
      return;
    }

    try {
      // Extract Excel operations from content
      const operations = this.extractExcelOperations(content);
      
      if (operations.length === 0) {
        this.showNotification('No Excel operations found', 'info');
        return;
      }

      // Execute operations
      for (const op of operations) {
        await this.mcpClient.callTool(op.tool, op.args);
      }

      this.showNotification('Operations executed', 'success');
      
    } catch (error) {
      console.error('Failed to execute content:', error);
      this.showNotification('Execution failed', 'error');
    }
  }

  /**
   * Extract Excel operations from content
   */
  extractExcelOperations(content) {
    const operations = [];
    
    // Simple pattern matching - in production, use more sophisticated parsing
    if (content.includes('=')) {
      operations.push({
        tool: 'excel/write-data',
        args: { data: content, range: 'A1' }
      });
    }
    
    return operations;
  }

  /**
   * Fallback processing when MCP not available
   */
  async processFallback(message) {
    console.log('Using fallback processing for:', message);
    
    // Simple keyword-based responses
    const lower = message.toLowerCase();
    
    if (lower.includes('hello') || lower.includes('hi')) {
      return 'Hello! I\'m your Arcadeus AI assistant. How can I help you with your Excel work today?';
    }
    
    if (lower.includes('help')) {
      return 'I can help you with Excel operations, financial modeling, data analysis, and more. Try asking me to create formulas, analyze data, or generate reports.';
    }
    
    return 'I understand you want help with Excel. Unfortunately, the advanced MCP features are not available right now, but I\'m still here to assist you.';
  }

  /**
   * Show notification
   */
  showNotification(message, type = 'info') {
    if (this.notificationManager) {
      this.notificationManager.sendNotification({
        title: 'Chat',
        message: message,
        category: 'chat',
        priority: type === 'error' ? 'high' : 'normal'
      });
    } else {
      // Fallback notification
      const notification = document.createElement('div');
      notification.className = `chat-notification ${type}`;
      notification.textContent = message;
      document.body.appendChild(notification);
      
      setTimeout(() => {
        notification.remove();
      }, 3000);
    }
  }

  /**
   * Show error message
   */
  showError(message) {
    this.showNotification(message, 'error');
    console.error('MCP Chat Error:', message);
  }

  /**
   * Cleanup when page unloads
   */
  async cleanup() {
    if (this.mcpClient) {
      await this.mcpClient.disconnect();
    }
    
    if (this.eventMonitor) {
      await this.eventMonitor.cleanup();
    }
  }
}

// Global window methods for quick actions (called from HTML)
window.enhancedChat = {
  validateFinancialModel: async function() {
    if (window.mcpChatHandler?.mcpClient) {
      const result = await window.mcpChatHandler.mcpClient.callTool('ai/validate-model', {
        modelData: window.formHandler?.collectAllModelData() || {}
      });
      
      window.mcpChatHandler.addMessageToDisplay('assistant', result.content[0]?.text || 'Model validated');
    }
  },
  
  generateExcelPage: async function(config) {
    console.log('Generating Excel page with config:', config);
    // Implementation would generate Excel content
  },
  
  suggestFormulas: async function() {
    if (window.mcpChatHandler?.mcpClient) {
      const result = await window.mcpChatHandler.mcpClient.callTool('excel/suggest-formulas', {
        context: 'current selection'
      });
      
      window.mcpChatHandler.addMessageToDisplay('assistant', result.content[0]?.text || 'Here are some formula suggestions...');
    }
  }
};

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', async () => {
  window.mcpChatHandler = new MCPChatHandler();
  await window.mcpChatHandler.initialize();
});

// Cleanup on page unload
window.addEventListener('beforeunload', () => {
  if (window.mcpChatHandler) {
    window.mcpChatHandler.cleanup();
  }
});

// Export for use in other scripts
window.MCPChatHandler = MCPChatHandler;