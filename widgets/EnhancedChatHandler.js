class EnhancedChatHandler {
  constructor() {
    this.modelProvider = new ModelProviderManager();
    this.excelReader = new ExcelContextReader();
    this.commandExecutor = new ExcelCommandExecutor();
    this.conversationHistory = [];
    this.currentStreamingMessage = null;
    this.isProcessing = false;
  }

  async initialize() {
    console.log('Initializing Enhanced Chat Handler...');
    
    // Initialize Excel context reader
    await this.excelReader.initialize();
    
    // Set up UI event listeners
    this.setupUIListeners();
    
    // Listen for Excel selection changes
    window.addEventListener('excelSelectionChanged', (event) => {
      this.handleSelectionChange(event.detail);
    });
    
    // Set up keyboard shortcuts for undo/redo
    this.setupKeyboardShortcuts();
    
    console.log('Enhanced Chat Handler initialized');
  }

  setupUIListeners() {
    // Send button
    const sendBtn = document.getElementById('sendMessage');
    if (sendBtn) {
      sendBtn.addEventListener('click', () => this.sendMessage());
    }
    
    // Chat input with Enter key support
    const chatInput = document.getElementById('chatInput');
    if (chatInput) {
      chatInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          this.sendMessage();
        }
      });
    }
    
    // Model selector
    const modelSelector = document.getElementById('modelSelector');
    if (modelSelector) {
      modelSelector.addEventListener('change', (e) => {
        this.modelProvider.setProvider(e.target.value);
        this.showNotification(`Switched to ${e.target.selectedOptions[0].text}`);
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
  }

  setupKeyboardShortcuts() {
    document.addEventListener('keydown', (e) => {
      // Ctrl/Cmd + Z for undo
      if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
        e.preventDefault();
        this.undo();
      }
      
      // Ctrl/Cmd + Shift + Z or Ctrl/Cmd + Y for redo
      if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) {
        e.preventDefault();
        this.redo();
      }
    });
  }

  async sendMessage() {
    const chatInput = document.getElementById('chatInput');
    const message = chatInput?.value.trim();
    
    if (!message || this.isProcessing) {
      return;
    }
    
    this.isProcessing = true;
    chatInput.value = '';
    
    try {
      // Add user message to UI
      this.addMessageToUI('user', message);
      
      // Get current Excel context
      const excelContext = await this.excelReader.getFullContext();
      
      // Get selected cell context for more detail
      const selectedCell = await this.excelReader.getSelectedCellContext();
      
      // Build conversation context
      const context = {
        excel: excelContext,
        selectedCell: selectedCell,
        formData: this.getFormData(),
        history: this.conversationHistory.slice(-10) // Last 10 messages
      };
      
      // Add to conversation history
      this.conversationHistory.push({ role: 'user', content: message });
      
      // Create streaming message container
      const streamingContainer = this.createStreamingMessage();
      
      // Send message with streaming
      const fullResponse = await this.modelProvider.sendMessage(
        this.conversationHistory,
        context,
        (chunk) => this.handleStreamingChunk(chunk, streamingContainer)
      );
      
      // Process any commands in the response
      await this.processAICommands(fullResponse);
      
      // Add to conversation history
      this.conversationHistory.push({ role: 'assistant', content: fullResponse });
      
    } catch (error) {
      console.error('Error sending message:', error);
      this.addMessageToUI('error', `Error: ${error.message}`);
    } finally {
      this.isProcessing = false;
      this.currentStreamingMessage = null;
    }
  }

  createStreamingMessage() {
    const chatMessages = document.getElementById('chatMessages');
    const messageDiv = document.createElement('div');
    messageDiv.className = 'chat-message assistant-message streaming';
    
    messageDiv.innerHTML = `
      <div class="message-avatar">
        <img src="assets/ai-avatar.svg" alt="AI">
      </div>
      <div class="message-content">
        <div class="message-header">
          <span class="model-name">${this.modelProvider.providers[this.modelProvider.currentProvider].name}</span>
          <span class="typing-indicator">Typing...</span>
        </div>
        <div class="message-text"></div>
        <div class="message-actions" style="display: none;">
          <button class="action-btn undo-btn" onclick="window.enhancedChat.undoLastCommand()">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M3 7v6h6"/>
              <path d="M21 17a9 9 0 00-9-9 9 9 0 00-6 2.3L3 13"/>
            </svg>
            Undo Changes
          </button>
          <button class="action-btn copy-btn" onclick="window.enhancedChat.copyMessage(this)">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <rect x="9" y="9" width="13" height="13" rx="2" ry="2"/>
              <path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/>
            </svg>
            Copy
          </button>
        </div>
      </div>
    `;
    
    chatMessages.appendChild(messageDiv);
    this.currentStreamingMessage = messageDiv;
    
    // Scroll to bottom
    chatMessages.scrollTop = chatMessages.scrollHeight;
    
    return messageDiv;
  }

  handleStreamingChunk(chunk, container) {
    if (!container) return;
    
    const textElement = container.querySelector('.message-text');
    const typingIndicator = container.querySelector('.typing-indicator');
    
    if (textElement) {
      // Append chunk and render markdown
      textElement.textContent += chunk;
      
      // Convert markdown to HTML (you can use a library like marked.js)
      textElement.innerHTML = this.renderMarkdown(textElement.textContent);
      
      // Scroll to bottom
      const chatMessages = document.getElementById('chatMessages');
      chatMessages.scrollTop = chatMessages.scrollHeight;
    }
    
    // Hide typing indicator when done
    if (chunk === '') {
      if (typingIndicator) typingIndicator.style.display = 'none';
      container.classList.remove('streaming');
      
      // Show action buttons
      const actions = container.querySelector('.message-actions');
      if (actions) actions.style.display = 'flex';
    }
  }

  renderMarkdown(text) {
    // Basic markdown rendering (consider using marked.js for full support)
    return text
      .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
      .replace(/\*(.*?)\*/g, '<em>$1</em>')
      .replace(/`(.*?)`/g, '<code>$1</code>')
      .replace(/```(\w+)?\n([\s\S]*?)```/g, '<pre><code class="language-$1">$2</code></pre>')
      .replace(/\n/g, '<br>');
  }

  async processAICommands(response) {
    // Parse response for Excel commands
    const commandPattern = /\[EXCEL_COMMAND:(.*?)\]/g;
    const commands = [];
    let match;
    
    while ((match = commandPattern.exec(response)) !== null) {
      try {
        const command = JSON.parse(match[1]);
        commands.push(command);
      } catch (e) {
        console.error('Failed to parse command:', e);
      }
    }
    
    if (commands.length > 0) {
      console.log(`Executing ${commands.length} Excel commands...`);
      
      for (const command of commands) {
        const result = await this.commandExecutor.executeCommand(command);
        
        if (!result.success) {
          this.showNotification(`Command failed: ${result.error}`, 'error');
        } else {
          this.showNotification(`Applied: ${command.description}`, 'success');
        }
      }
    }
  }

  async undo() {
    const result = await this.commandExecutor.undo();
    
    if (result.success) {
      this.showNotification(result.description, 'info');
      this.updateUndoRedoButtons();
    } else {
      this.showNotification(result.error, 'error');
    }
  }

  async redo() {
    const result = await this.commandExecutor.redo();
    
    if (result.success) {
      this.showNotification(result.description, 'info');
      this.updateUndoRedoButtons();
    } else {
      this.showNotification(result.error, 'error');
    }
  }

  undoLastCommand() {
    this.undo();
  }

  updateUndoRedoButtons() {
    const history = this.commandExecutor.getHistory();
    const undoBtn = document.getElementById('undoBtn');
    const redoBtn = document.getElementById('redoBtn');
    
    if (undoBtn) {
      undoBtn.disabled = !history.some(h => h.canUndo && h.isCurrent);
    }
    
    if (redoBtn) {
      redoBtn.disabled = !history.some(h => h.canRedo);
    }
  }

  addMessageToUI(type, content) {
    const chatMessages = document.getElementById('chatMessages');
    const messageDiv = document.createElement('div');
    messageDiv.className = `chat-message ${type}-message`;
    
    let avatar = '';
    if (type === 'user') {
      avatar = '<div class="message-avatar"><img src="assets/user-avatar.svg" alt="You"></div>';
    } else if (type === 'assistant') {
      avatar = `<div class="message-avatar"><img src="assets/ai-avatar.svg" alt="AI"></div>`;
    }
    
    messageDiv.innerHTML = `
      ${avatar}
      <div class="message-content">
        ${type === 'assistant' ? `<div class="message-header"><span class="model-name">${this.modelProvider.providers[this.modelProvider.currentProvider].name}</span></div>` : ''}
        <div class="message-text">${this.renderMarkdown(content)}</div>
      </div>
    `;
    
    chatMessages.appendChild(messageDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }

  showNotification(message, type = 'info') {
    // Create or update notification element
    let notification = document.getElementById('chatNotification');
    
    if (!notification) {
      notification = document.createElement('div');
      notification.id = 'chatNotification';
      notification.className = 'chat-notification';
      document.body.appendChild(notification);
    }
    
    notification.className = `chat-notification ${type}`;
    notification.textContent = message;
    notification.style.display = 'block';
    
    // Auto-hide after 3 seconds
    setTimeout(() => {
      notification.style.display = 'none';
    }, 3000);
  }

  getFormData() {
    // Get current form data if available
    if (window.formHandler) {
      try {
        return window.formHandler.collectAllModelData();
      } catch (e) {
        console.error('Could not collect form data:', e);
      }
    }
    return {};
  }

  handleSelectionChange(detail) {
    // Update UI to show current selection
    const selectionIndicator = document.getElementById('selectionIndicator');
    if (selectionIndicator) {
      selectionIndicator.textContent = `Selected: ${detail.address}`;
    }
  }

  copyMessage(button) {
    const messageText = button.closest('.message-content').querySelector('.message-text').textContent;
    navigator.clipboard.writeText(messageText);
    this.showNotification('Copied to clipboard', 'success');
  }

  async generateExcelPage(specifications) {
    // AI can call this to generate new Excel pages
    const command = {
      type: 'createSheet',
      params: {
        name: specifications.name || 'New Sheet',
        position: specifications.position
      },
      description: `Create sheet: ${specifications.name}`,
      affectedRanges: []
    };
    
    const result = await this.commandExecutor.executeCommand(command);
    
    if (result.success && specifications.content) {
      // Add content to the new sheet
      const contentCommand = {
        type: 'batchUpdate',
        params: {
          updates: specifications.content.map(item => ({
            type: item.type || 'setValue',
            params: {
              worksheet: specifications.name,
              ...item
            }
          }))
        },
        description: `Populate ${specifications.name}`,
        affectedRanges: specifications.content.map(item => ({
          worksheet: specifications.name,
          address: item.range
        }))
      };
      
      await this.commandExecutor.executeCommand(contentCommand);
    }
    
    return result;
  }

  async validateFinancialModel() {
    const context = await this.excelReader.getFullContext();
    
    // Send to AI for validation
    const validationPrompt = `Please validate this financial model and identify any errors or inconsistencies:\n${JSON.stringify(context, null, 2)}`;
    
    this.conversationHistory.push({ role: 'user', content: validationPrompt });
    
    const response = await this.modelProvider.sendMessage(
      this.conversationHistory,
      { excel: context },
      (chunk) => this.handleStreamingChunk(chunk, this.createStreamingMessage())
    );
    
    this.conversationHistory.push({ role: 'assistant', content: response });
    
    return response;
  }

  clearChat() {
    this.conversationHistory = [];
    const chatMessages = document.getElementById('chatMessages');
    if (chatMessages) {
      chatMessages.innerHTML = '';
    }
    this.commandExecutor.clearHistory();
    this.showNotification('Chat cleared', 'info');
  }
}

// Initialize and expose globally
window.EnhancedChatHandler = EnhancedChatHandler;
window.enhancedChat = new EnhancedChatHandler();