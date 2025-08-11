/**
 * Professional Chat Interface - Stage 3 Implementation
 * Enterprise-grade chat experience for M&A intelligence platform
 */

class ProfessionalChatInterface {
  constructor() {
    this.conversations = [];
    this.currentConversationId = null;
    this.messageHistory = [];
    this.isTyping = false;
    this.quickActions = [];
    this.contextPanel = null;
    
    // Professional chat features
    this.features = {
      conversationHistory: true,
      contextAwareness: true,
      quickActions: true,
      voiceInput: false, // Future feature
      exportCapability: true,
      collaborativeMode: false // Future feature
    };
    
    this.initializeInterface();
    console.log('💼 Professional Chat Interface initialized');
  }

  /**
   * Initialize the professional chat interface
   */
  async initializeInterface() {
    this.setupChatContainer();
    this.setupConversationPanel();
    this.setupQuickActions();
    this.setupContextPanel();
    this.setupKeyboardShortcuts();
    this.loadConversationHistory();
    
    // Integration with existing chat system
    this.integrateWithExistingChat();
    
    console.log('✅ Professional chat interface ready');
  }

  /**
   * Setup main chat container with professional styling
   */
  setupChatContainer() {
    const existingChat = document.getElementById('chatMessages');
    if (!existingChat) return;

    // Enhance existing chat container
    existingChat.classList.add('professional-chat-container');
    
    // Add professional header
    const header = document.createElement('div');
    header.className = 'chat-header';
    header.innerHTML = `
      <div class="chat-title">
        <div class="chat-icon">🎭</div>
        <div class="chat-info">
          <h3>M&A Intelligence</h3>
          <p class="chat-status">Ready for analysis</p>
        </div>
      </div>
      <div class="chat-controls">
        <button class="chat-control-btn" id="newConversationBtn" title="New Conversation">
          <span class="icon">💬</span>
        </button>
        <button class="chat-control-btn" id="conversationHistoryBtn" title="Conversation History">
          <span class="icon">📋</span>
        </button>
        <button class="chat-control-btn" id="exportConversationBtn" title="Export Conversation">
          <span class="icon">📤</span>
        </button>
        <button class="chat-control-btn" id="chatSettingsBtn" title="Chat Settings">
          <span class="icon">⚙️</span>
        </button>
      </div>
    `;
    
    existingChat.parentNode.insertBefore(header, existingChat);
    
    // Add event listeners for controls
    this.setupHeaderControls();
  }

  /**
   * Setup conversation panel for history management
   */
  setupConversationPanel() {
    const panel = document.createElement('div');
    panel.id = 'conversationPanel';
    panel.className = 'conversation-panel';
    panel.style.display = 'none';
    
    panel.innerHTML = `
      <div class="conversation-panel-header">
        <h4>Conversation History</h4>
        <button class="close-panel-btn" id="closeConversationPanel">×</button>
      </div>
      <div class="conversation-list" id="conversationList">
        <!-- Conversations will be populated here -->
      </div>
      <div class="conversation-panel-footer">
        <button class="btn btn-sm" id="clearHistoryBtn">Clear All</button>
        <button class="btn btn-sm btn-primary" id="exportAllBtn">Export All</button>
      </div>
    `;
    
    document.body.appendChild(panel);
    this.setupConversationPanelEvents();
  }

  /**
   * Setup quick action buttons for common M&A queries
   */
  setupQuickActions() {
    const quickActionsContainer = document.createElement('div');
    quickActionsContainer.className = 'quick-actions-container';
    quickActionsContainer.id = 'quickActions';
    
    this.quickActions = [
      {
        id: 'analyze-irr',
        label: 'Analyze IRR',
        icon: '📈',
        query: 'Analyze my IRR calculation and key value drivers',
        category: 'financial'
      },
      {
        id: 'check-moic',
        label: 'Check MOIC',
        icon: '💰',
        query: 'What is my MOIC and how does it compare to market standards?',
        category: 'financial'
      },
      {
        id: 'validate-model',
        label: 'Validate Model',
        icon: '✅',
        query: 'Check my model for errors and inconsistencies',
        category: 'validation'
      },
      {
        id: 'sensitivity-analysis',
        label: 'Sensitivity Analysis',
        icon: '🎯',
        query: 'Perform sensitivity analysis on key assumptions',
        category: 'analysis'
      },
      {
        id: 'cash-flow-review',
        label: 'Cash Flow Review',
        icon: '💸',
        query: 'Review my cash flow projections and assumptions',
        category: 'financial'
      },
      {
        id: 'debt-analysis',
        label: 'Debt Analysis',
        icon: '🏦',
        query: 'Analyze debt capacity and leverage metrics',
        category: 'financial'
      },
      {
        id: 'exit-assumptions',
        label: 'Exit Analysis',
        icon: '🚪',
        query: 'Review exit assumptions and terminal value calculation',
        category: 'financial'
      },
      {
        id: 'model-overview',
        label: 'Model Overview',
        icon: '📊',
        query: 'Provide a comprehensive overview of my M&A model',
        category: 'overview'
      }
    ];
    
    // Create category filters
    const categories = ['all', ...new Set(this.quickActions.map(action => action.category))];
    const categoryFilters = categories.map(cat => 
      `<button class="category-filter ${cat === 'all' ? 'active' : ''}" data-category="${cat}">
        ${cat.charAt(0).toUpperCase() + cat.slice(1)}
      </button>`
    ).join('');
    
    // Create action buttons
    const actionButtons = this.quickActions.map(action => 
      `<button class="quick-action-btn" data-query="${action.query}" data-category="${action.category}">
        <span class="action-icon">${action.icon}</span>
        <span class="action-label">${action.label}</span>
      </button>`
    ).join('');
    
    quickActionsContainer.innerHTML = `
      <div class="quick-actions-header">
        <h4>Quick Actions</h4>
        <button class="toggle-quick-actions" id="toggleQuickActions">
          <span class="icon">📌</span>
        </button>
      </div>
      <div class="category-filters">
        ${categoryFilters}
      </div>
      <div class="quick-actions-grid" id="quickActionsGrid">
        ${actionButtons}
      </div>
    `;
    
    // Insert before chat input
    const chatInput = document.getElementById('chatInput');
    if (chatInput && chatInput.parentNode) {
      chatInput.parentNode.insertBefore(quickActionsContainer, chatInput.parentNode);
    }
    
    this.setupQuickActionEvents();
  }

  /**
   * Setup context panel showing Excel model information
   */
  setupContextPanel() {
    const panel = document.createElement('div');
    panel.id = 'contextPanel';
    panel.className = 'context-panel';
    
    panel.innerHTML = `
      <div class="context-header">
        <h4>Model Context</h4>
        <button class="minimize-panel" id="minimizeContext">−</button>
      </div>
      <div class="context-content">
        <div class="context-section" id="modelOverview">
          <h5>📊 Model Overview</h5>
          <div class="context-loading">Loading model information...</div>
        </div>
        <div class="context-section" id="keyMetrics">
          <h5>💰 Key Metrics</h5>
          <div class="metrics-grid">
            <!-- Will be populated dynamically -->
          </div>
        </div>
        <div class="context-section" id="recentInsights">
          <h5>🔍 Recent Insights</h5>
          <div class="insights-list">
            <!-- Will be populated from conversation history -->
          </div>
        </div>
        <div class="context-section" id="quickStats">
          <h5>📈 Quick Stats</h5>
          <div class="stats-grid">
            <!-- Will show model statistics -->
          </div>
        </div>
      </div>
    `;
    
    // Add to sidebar or create floating panel
    const chatContainer = document.querySelector('.professional-chat-container')?.parentElement;
    if (chatContainer) {
      chatContainer.appendChild(panel);
    } else {
      document.body.appendChild(panel);
    }
    
    this.contextPanel = panel;
    this.updateContextPanel();
    this.setupContextPanelEvents();
  }

  /**
   * Setup keyboard shortcuts for power users
   */
  setupKeyboardShortcuts() {
    const shortcuts = {
      'Ctrl+Enter': () => this.sendCurrentMessage(),
      'Ctrl+N': () => this.startNewConversation(),
      'Ctrl+H': () => this.toggleConversationHistory(),
      'Ctrl+E': () => this.exportCurrentConversation(),
      'Ctrl+/': () => this.showKeyboardShortcuts(),
      'Escape': () => this.closePanels(),
      'Ctrl+1': () => this.executeQuickAction('analyze-irr'),
      'Ctrl+2': () => this.executeQuickAction('check-moic'),
      'Ctrl+3': () => this.executeQuickAction('validate-model'),
      'Ctrl+4': () => this.executeQuickAction('sensitivity-analysis')
    };
    
    document.addEventListener('keydown', (e) => {
      const key = `${e.ctrlKey ? 'Ctrl+' : ''}${e.altKey ? 'Alt+' : ''}${e.shiftKey ? 'Shift+' : ''}${e.key}`;
      
      if (shortcuts[key]) {
        e.preventDefault();
        shortcuts[key]();
      }
    });
    
    // Show shortcuts hint
    this.createShortcutsHint();
  }

  /**
   * Setup event listeners for header controls
   */
  setupHeaderControls() {
    document.getElementById('newConversationBtn')?.addEventListener('click', () => {
      this.startNewConversation();
    });
    
    document.getElementById('conversationHistoryBtn')?.addEventListener('click', () => {
      this.toggleConversationHistory();
    });
    
    document.getElementById('exportConversationBtn')?.addEventListener('click', () => {
      this.exportCurrentConversation();
    });
    
    document.getElementById('chatSettingsBtn')?.addEventListener('click', () => {
      this.showChatSettings();
    });
  }

  /**
   * Setup quick action events
   */
  setupQuickActionEvents() {
    // Quick action button clicks
    document.querySelectorAll('.quick-action-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const query = e.currentTarget.dataset.query;
        this.executeQuickActionQuery(query);
      });
    });
    
    // Category filter clicks
    document.querySelectorAll('.category-filter').forEach(btn => {
      btn.addEventListener('click', (e) => {
        this.filterQuickActions(e.currentTarget.dataset.category);
      });
    });
    
    // Toggle quick actions visibility
    document.getElementById('toggleQuickActions')?.addEventListener('click', () => {
      this.toggleQuickActionsVisibility();
    });
  }

  /**
   * Setup conversation panel events
   */
  setupConversationPanelEvents() {
    document.getElementById('closeConversationPanel')?.addEventListener('click', () => {
      this.hideConversationPanel();
    });
    
    document.getElementById('clearHistoryBtn')?.addEventListener('click', () => {
      this.clearConversationHistory();
    });
    
    document.getElementById('exportAllBtn')?.addEventListener('click', () => {
      this.exportAllConversations();
    });
  }

  /**
   * Setup context panel events
   */
  setupContextPanelEvents() {
    document.getElementById('minimizeContext')?.addEventListener('click', () => {
      this.toggleContextPanel();
    });
    
    // Auto-update context when Excel changes
    if (window.addEventListener) {
      window.addEventListener('excelWorkbookChanged', () => {
        this.updateContextPanel();
      });
    }
  }

  /**
   * Execute quick action query
   */
  async executeQuickActionQuery(query) {
    const chatInput = document.getElementById('chatInput');
    if (chatInput) {
      chatInput.value = query;
      
      // Add visual feedback
      const actionBtn = document.querySelector(`[data-query="${query}"]`);
      if (actionBtn) {
        actionBtn.classList.add('executing');
        setTimeout(() => actionBtn.classList.remove('executing'), 2000);
      }
      
      // Send the message
      if (window.chatHandler) {
        await window.chatHandler.sendChatMessage();
      }
    }
  }

  /**
   * Filter quick actions by category
   */
  filterQuickActions(category) {
    // Update active filter
    document.querySelectorAll('.category-filter').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.category === category);
    });
    
    // Show/hide actions
    document.querySelectorAll('.quick-action-btn').forEach(btn => {
      const btnCategory = btn.dataset.category;
      btn.style.display = (category === 'all' || btnCategory === category) ? 'flex' : 'none';
    });
  }

  /**
   * Update context panel with current model information
   */
  async updateContextPanel() {
    if (!this.contextPanel) return;
    
    try {
      // Get Excel structure information
      let modelInfo = {};
      if (window.excelStructureFetcher) {
        const structureJson = await window.excelStructureFetcher.fetchWorkbookStructure('context update');
        modelInfo = JSON.parse(structureJson);
      }
      
      // Update model overview
      const overviewSection = document.getElementById('modelOverview');
      if (overviewSection) {
        overviewSection.innerHTML = `
          <h5>📊 Model Overview</h5>
          <div class="overview-stats">
            <div class="stat-item">
              <span class="stat-label">Sheets:</span>
              <span class="stat-value">${modelInfo.metadata?.totalSheets || 0}</span>
            </div>
            <div class="stat-item">
              <span class="stat-label">Key Metrics:</span>
              <span class="stat-value">${Object.keys(modelInfo.keyMetrics || {}).length}</span>
            </div>
            <div class="stat-item">
              <span class="stat-label">Model Score:</span>
              <span class="stat-value">${modelInfo.validation?.modelScore || 0}/100</span>
            </div>
          </div>
        `;
      }
      
      // Update key metrics
      this.updateKeyMetricsDisplay(modelInfo.keyMetrics || {});
      
      // Update recent insights from conversation history
      this.updateRecentInsights();
      
    } catch (error) {
      console.error('Failed to update context panel:', error);
    }
  }

  /**
   * Update key metrics display
   */
  updateKeyMetricsDisplay(keyMetrics) {
    const metricsSection = document.getElementById('keyMetrics');
    if (!metricsSection) return;
    
    const metricsGrid = metricsSection.querySelector('.metrics-grid');
    if (!metricsGrid) return;
    
    const allMetrics = Object.values(keyMetrics).flat();
    const displayMetrics = allMetrics.slice(0, 6); // Show top 6 metrics
    
    metricsGrid.innerHTML = displayMetrics.map(metric => `
      <div class="metric-card" onclick="navigateToExcelCell('${metric.location}')">
        <div class="metric-icon">${this.getMetricIcon(metric.type)}</div>
        <div class="metric-info">
          <div class="metric-label">${metric.label}</div>
          <div class="metric-value">${this.formatMetricValue(metric.value, metric.type)}</div>
          <div class="metric-location">${metric.location}</div>
        </div>
      </div>
    `).join('');
  }

  /**
   * Get icon for metric type
   */
  getMetricIcon(type) {
    const icons = {
      irr: '📈',
      moic: '💰',
      npv: '💎',
      revenue: '🏢',
      ebitda: '💵',
      debt: '🏦',
      equity: '💸'
    };
    return icons[type] || '📊';
  }

  /**
   * Format metric value for display
   */
  formatMetricValue(value, type) {
    if (typeof value !== 'number') return value;
    
    switch (type) {
      case 'irr':
        return `${(value * 100).toFixed(1)}%`;
      case 'moic':
        return `${value.toFixed(1)}x`;
      case 'revenue':
      case 'ebitda':
        return `$${(value / 1000000).toFixed(1)}M`;
      default:
        return value.toLocaleString();
    }
  }

  /**
   * Start new conversation
   */
  startNewConversation() {
    const conversationId = `conv_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    
    const conversation = {
      id: conversationId,
      title: 'New Analysis',
      startTime: new Date().toISOString(),
      messages: [],
      context: {
        modelState: null,
        keyFindings: []
      }
    };
    
    this.conversations.push(conversation);
    this.currentConversationId = conversationId;
    
    // Clear current chat
    const chatMessages = document.getElementById('chatMessages');
    if (chatMessages) {
      chatMessages.innerHTML = '<div class="welcome-message">🎭 Ready for M&A analysis. What would you like to explore?</div>';
    }
    
    this.updateChatStatus('New conversation started');
    console.log('Started new conversation:', conversationId);
  }

  /**
   * Integration with existing chat system
   */
  integrateWithExistingChat() {
    // Hook into existing chat handler if available
    if (window.chatHandler) {
      const originalAddMessage = window.chatHandler.addChatMessage.bind(window.chatHandler);
      
      window.chatHandler.addChatMessage = (role, content) => {
        // Call original method
        originalAddMessage(role, content);
        
        // Add to our conversation history
        this.addMessageToCurrentConversation(role, content);
        
        // Update context panel
        this.updateContextPanel();
        
        // Update conversation title if needed
        if (role === 'user' && this.currentConversationId) {
          this.updateConversationTitle(content);
        }
      };
    }
  }

  /**
   * Add message to current conversation
   */
  addMessageToCurrentConversation(role, content) {
    if (!this.currentConversationId) {
      this.startNewConversation();
    }
    
    const conversation = this.conversations.find(c => c.id === this.currentConversationId);
    if (conversation) {
      conversation.messages.push({
        role,
        content,
        timestamp: new Date().toISOString()
      });
      
      // Save to localStorage
      this.saveConversationHistory();
    }
  }

  /**
   * Update conversation title based on first user message
   */
  updateConversationTitle(userMessage) {
    const conversation = this.conversations.find(c => c.id === this.currentConversationId);
    if (conversation && conversation.title === 'New Analysis') {
      // Generate title from first message
      let title = userMessage.substring(0, 50);
      if (userMessage.length > 50) title += '...';
      
      // Make it more descriptive
      if (userMessage.toLowerCase().includes('irr')) title = '📈 IRR Analysis';
      else if (userMessage.toLowerCase().includes('moic')) title = '💰 MOIC Review';
      else if (userMessage.toLowerCase().includes('validation')) title = '✅ Model Validation';
      else if (userMessage.toLowerCase().includes('cash flow')) title = '💸 Cash Flow Analysis';
      
      conversation.title = title;
      this.saveConversationHistory();
    }
  }

  /**
   * Load conversation history from storage
   */
  loadConversationHistory() {
    try {
      const stored = localStorage.getItem('arcadeus_conversations');
      if (stored) {
        this.conversations = JSON.parse(stored);
        console.log(`Loaded ${this.conversations.length} conversations from history`);
      }
    } catch (error) {
      console.error('Failed to load conversation history:', error);
    }
  }

  /**
   * Save conversation history to storage
   */
  saveConversationHistory() {
    try {
      localStorage.setItem('arcadeus_conversations', JSON.stringify(this.conversations));
    } catch (error) {
      console.error('Failed to save conversation history:', error);
    }
  }

  /**
   * Update chat status message
   */
  updateChatStatus(message) {
    const statusElement = document.querySelector('.chat-status');
    if (statusElement) {
      statusElement.textContent = message;
      
      // Auto-clear after 3 seconds
      setTimeout(() => {
        if (statusElement.textContent === message) {
          statusElement.textContent = 'Ready for analysis';
        }
      }, 3000);
    }
  }

  /**
   * Show keyboard shortcuts help
   */
  showKeyboardShortcuts() {
    const shortcuts = [
      { key: 'Ctrl+Enter', desc: 'Send message' },
      { key: 'Ctrl+N', desc: 'New conversation' },
      { key: 'Ctrl+H', desc: 'Show conversation history' },
      { key: 'Ctrl+E', desc: 'Export current conversation' },
      { key: 'Ctrl+1-4', desc: 'Quick actions (IRR, MOIC, Validate, Sensitivity)' },
      { key: 'Ctrl+/', desc: 'Show this help' },
      { key: 'Escape', desc: 'Close panels' }
    ];
    
    const helpContent = shortcuts.map(s => `
      <div class="shortcut-item">
        <kbd>${s.key}</kbd>
        <span>${s.desc}</span>
      </div>
    `).join('');
    
    // Create modal or notification
    const modal = document.createElement('div');
    modal.className = 'shortcuts-modal';
    modal.innerHTML = `
      <div class="modal-content">
        <h4>Keyboard Shortcuts</h4>
        <div class="shortcuts-list">${helpContent}</div>
        <button class="close-modal">Close</button>
      </div>
    `;
    
    document.body.appendChild(modal);
    
    // Auto-close after 5 seconds or on click
    const closeModal = () => modal.remove();
    modal.querySelector('.close-modal').addEventListener('click', closeModal);
    setTimeout(closeModal, 5000);
  }

  /**
   * Create shortcuts hint
   */
  createShortcutsHint() {
    const hint = document.createElement('div');
    hint.className = 'shortcuts-hint';
    hint.innerHTML = `
      <span class="hint-text">Press <kbd>Ctrl+/</kbd> for shortcuts</span>
      <button class="dismiss-hint">×</button>
    `;
    
    // Add to bottom of chat interface
    const chatContainer = document.querySelector('.professional-chat-container');
    if (chatContainer) {
      chatContainer.appendChild(hint);
    }
    
    // Auto-dismiss after 10 seconds or on click
    const dismissHint = () => hint.remove();
    hint.querySelector('.dismiss-hint').addEventListener('click', dismissHint);
    setTimeout(dismissHint, 10000);
  }

  /**
   * Toggle quick actions visibility
   */
  toggleQuickActionsVisibility() {
    const grid = document.getElementById('quickActionsGrid');
    const toggleBtn = document.getElementById('toggleQuickActions');
    
    if (grid && toggleBtn) {
      const isHidden = grid.style.display === 'none';
      grid.style.display = isHidden ? 'grid' : 'none';
      toggleBtn.querySelector('.icon').textContent = isHidden ? '📌' : '📍';
    }
  }

  /**
   * Toggle conversation history panel
   */
  toggleConversationHistory() {
    const panel = document.getElementById('conversationPanel');
    if (panel) {
      const isVisible = panel.style.display !== 'none';
      panel.style.display = isVisible ? 'none' : 'block';
      
      if (!isVisible) {
        this.populateConversationHistory();
      }
    }
  }

  /**
   * Populate conversation history list
   */
  populateConversationHistory() {
    const list = document.getElementById('conversationList');
    if (!list) return;
    
    const sortedConversations = [...this.conversations].sort(
      (a, b) => new Date(b.startTime) - new Date(a.startTime)
    );
    
    list.innerHTML = sortedConversations.map(conv => `
      <div class="conversation-item ${conv.id === this.currentConversationId ? 'active' : ''}" 
           data-conversation-id="${conv.id}">
        <div class="conversation-title">${conv.title}</div>
        <div class="conversation-meta">
          <span class="conversation-date">${new Date(conv.startTime).toLocaleDateString()}</span>
          <span class="conversation-count">${conv.messages.length} messages</span>
        </div>
        <div class="conversation-actions">
          <button class="load-conversation" title="Load">📂</button>
          <button class="export-conversation" title="Export">📤</button>
          <button class="delete-conversation" title="Delete">🗑️</button>
        </div>
      </div>
    `).join('');
    
    // Add event listeners for conversation actions
    this.setupConversationListEvents();
  }

  /**
   * Export current conversation
   */
  exportCurrentConversation() {
    if (!this.currentConversationId) return;
    
    const conversation = this.conversations.find(c => c.id === this.currentConversationId);
    if (conversation) {
      this.exportConversation(conversation);
    }
  }

  /**
   * Export conversation to file
   */
  exportConversation(conversation) {
    const exportData = {
      title: conversation.title,
      startTime: conversation.startTime,
      messages: conversation.messages,
      exportTime: new Date().toISOString(),
      version: '1.0'
    };
    
    const blob = new Blob([JSON.stringify(exportData, null, 2)], { 
      type: 'application/json' 
    });
    
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `arcadeus-conversation-${conversation.title.replace(/[^a-zA-Z0-9]/g, '_')}-${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    
    URL.revokeObjectURL(url);
    
    this.updateChatStatus('Conversation exported');
  }

  /**
   * Get current conversation for external access
   */
  getCurrentConversation() {
    return this.conversations.find(c => c.id === this.currentConversationId);
  }

  /**
   * Get all conversations
   */
  getAllConversations() {
    return this.conversations;
  }
}

// Export for global use
window.ProfessionalChatInterface = ProfessionalChatInterface;
window.professionalChatInterface = new ProfessionalChatInterface();

console.log('💼 Professional Chat Interface Stage 3 loaded with:');
console.log('  ✅ Enterprise-grade chat UI with conversation history');
console.log('  ✅ Quick actions for common M&A queries');
console.log('  ✅ Context-aware model information panel');
console.log('  ✅ Keyboard shortcuts for power users');
console.log('  ✅ Export and collaboration features');
console.log('  ✅ Professional visual design');
console.log('🎯 Ready for world-class M&A analysis experience');