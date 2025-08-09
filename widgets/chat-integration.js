// Main integration script for the enhanced chat system
class ChatIntegration {
  constructor() {
    this.initialized = false;
  }

  async initialize() {
    console.log('Initializing Enhanced Chat System...');
    
    try {
      // Wait for Office to be ready
      await this.waitForOffice();
      
      // Initialize all components
      await this.initializeComponents();
      
      // Setup event listeners
      this.setupEventListeners();
      
      // Load saved settings
      this.loadSettings();
      
      this.initialized = true;
      console.log('Enhanced Chat System initialized successfully');
      
    } catch (error) {
      console.error('Failed to initialize chat system:', error);
      this.showError('Failed to initialize chat system. Please refresh the page.');
    }
  }

  async waitForOffice() {
    return new Promise((resolve) => {
      if (typeof Office !== 'undefined') {
        Office.onReady(() => {
          console.log('Office is ready');
          resolve();
        });
      } else {
        // For testing outside of Office
        console.log('Running in standalone mode');
        resolve();
      }
    });
  }

  async initializeComponents() {
    // Initialize the enhanced chat handler
    if (window.enhancedChat) {
      await window.enhancedChat.initialize();
    }
    
    // Initialize Excel context reader if available
    if (typeof Excel !== 'undefined') {
      await this.initializeExcelContext();
    }
    
    // Setup character counter
    this.setupCharacterCounter();
    
    // Setup settings modal
    this.setupSettingsModal();
  }

  async initializeExcelContext() {
    try {
      await Excel.run(async (context) => {
        // Get initial context
        const workbook = context.workbook;
        const worksheet = workbook.worksheets.getActiveWorksheet();
        const selectedRange = workbook.getSelectedRange();
        
        workbook.load('name');
        worksheet.load('name');
        selectedRange.load('address');
        
        await context.sync();
        
        // Update UI with initial context
        this.updateContextDisplay({
          workbook: workbook.name,
          worksheet: worksheet.name,
          selection: selectedRange.address
        });
        
        // Start watching for changes
        if (window.enhancedChat && window.enhancedChat.excelReader) {
          window.enhancedChat.excelReader.watchForChanges((context) => {
            this.updateContextDisplay({
              workbook: context.workbook.name,
              worksheet: context.sheets.find(s => s.isActive)?.name,
              selection: context.selection.address
            });
          }, 2000); // Update every 2 seconds
        }
      });
    } catch (error) {
      console.error('Failed to initialize Excel context:', error);
    }
  }

  updateContextDisplay(context) {
    const workbookElement = document.getElementById('workbookName');
    const sheetElement = document.getElementById('activeSheet');
    const rangeElement = document.getElementById('selectedRange');
    const indicatorElement = document.getElementById('selectionIndicator');
    
    if (workbookElement) workbookElement.textContent = context.workbook || 'N/A';
    if (sheetElement) sheetElement.textContent = context.worksheet || 'N/A';
    if (rangeElement) rangeElement.textContent = context.selection || 'N/A';
    if (indicatorElement) indicatorElement.textContent = `Selected: ${context.selection || 'None'}`;
  }

  setupEventListeners() {
    // Clear chat button
    const clearBtn = document.getElementById('clearChatBtn');
    if (clearBtn) {
      clearBtn.addEventListener('click', () => {
        if (confirm('Are you sure you want to clear the chat history?')) {
          window.enhancedChat.clearChat();
        }
      });
    }
    
    // Settings button
    const settingsBtn = document.getElementById('settingsBtn');
    if (settingsBtn) {
      settingsBtn.addEventListener('click', () => {
        this.showSettingsModal();
      });
    }
    
    // Refresh context button
    const refreshBtn = document.getElementById('refreshContextBtn');
    if (refreshBtn) {
      refreshBtn.addEventListener('click', async () => {
        if (window.enhancedChat && window.enhancedChat.excelReader) {
          const context = await window.enhancedChat.excelReader.getFullContext();
          this.updateContextDisplay({
            workbook: context.workbook?.name,
            worksheet: context.sheets?.find(s => s.isActive)?.name,
            selection: context.selection?.address
          });
          window.enhancedChat.showNotification('Context refreshed', 'success');
        }
      });
    }
  }

  setupCharacterCounter() {
    const chatInput = document.getElementById('chatInput');
    const charCount = document.getElementById('charCount');
    
    if (chatInput && charCount) {
      chatInput.addEventListener('input', () => {
        const count = chatInput.value.length;
        charCount.textContent = count;
        
        if (count > 3800) {
          charCount.style.color = '#dc3545';
        } else if (count > 3000) {
          charCount.style.color = '#ffc107';
        } else {
          charCount.style.color = '#666';
        }
      });
    }
  }

  setupSettingsModal() {
    // Save settings handler
    window.enhancedChat.saveSettings = () => {
      const settings = {
        openaiKey: document.getElementById('openaiKey')?.value,
        anthropicKey: document.getElementById('anthropicKey')?.value,
        googleKey: document.getElementById('googleKey')?.value,
        autoReadExcel: document.getElementById('autoReadExcel')?.checked,
        streamResponses: document.getElementById('streamResponses')?.checked,
        maxHistory: document.getElementById('maxHistory')?.value
      };
      
      // Save to localStorage
      localStorage.setItem('chatSettings', JSON.stringify(settings));
      
      // Update API keys in model provider
      if (window.enhancedChat.modelProvider) {
        window.enhancedChat.modelProvider.apiKeys = {
          'gpt-4': settings.openaiKey,
          'claude-opus': settings.anthropicKey,
          'claude-sonnet': settings.anthropicKey,
          'gemini-pro': settings.googleKey
        };
      }
      
      // Close modal
      document.getElementById('chatSettingsModal').style.display = 'none';
      
      window.enhancedChat.showNotification('Settings saved', 'success');
    };
  }

  loadSettings() {
    const savedSettings = localStorage.getItem('chatSettings');
    if (savedSettings) {
      try {
        const settings = JSON.parse(savedSettings);
        
        // Apply settings to UI
        if (document.getElementById('autoReadExcel')) {
          document.getElementById('autoReadExcel').checked = settings.autoReadExcel !== false;
        }
        if (document.getElementById('streamResponses')) {
          document.getElementById('streamResponses').checked = settings.streamResponses !== false;
        }
        if (document.getElementById('maxHistory')) {
          document.getElementById('maxHistory').value = settings.maxHistory || 20;
        }
        
        // Apply API keys
        if (window.enhancedChat.modelProvider && settings) {
          window.enhancedChat.modelProvider.apiKeys = {
            'gpt-4': settings.openaiKey,
            'claude-opus': settings.anthropicKey,
            'claude-sonnet': settings.anthropicKey,
            'gemini-pro': settings.googleKey
          };
        }
        
      } catch (error) {
        console.error('Failed to load settings:', error);
      }
    }
  }

  showSettingsModal() {
    const modal = document.getElementById('chatSettingsModal');
    if (modal) {
      modal.style.display = 'flex';
      
      // Load current settings
      const savedSettings = localStorage.getItem('chatSettings');
      if (savedSettings) {
        try {
          const settings = JSON.parse(savedSettings);
          
          // Populate fields (don't show actual API keys for security)
          if (settings.openaiKey && document.getElementById('openaiKey')) {
            document.getElementById('openaiKey').placeholder = 'OpenAI API Key (configured)';
          }
          if (settings.anthropicKey && document.getElementById('anthropicKey')) {
            document.getElementById('anthropicKey').placeholder = 'Anthropic API Key (configured)';
          }
          if (settings.googleKey && document.getElementById('googleKey')) {
            document.getElementById('googleKey').placeholder = 'Google API Key (configured)';
          }
        } catch (error) {
          console.error('Failed to load settings for modal:', error);
        }
      }
    }
  }

  showError(message) {
    const notification = document.createElement('div');
    notification.className = 'chat-notification error';
    notification.textContent = message;
    document.body.appendChild(notification);
    
    setTimeout(() => {
      notification.remove();
    }, 5000);
  }
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    window.chatIntegration = new ChatIntegration();
    window.chatIntegration.initialize();
  });
} else {
  window.chatIntegration = new ChatIntegration();
  window.chatIntegration.initialize();
}

// Export for testing
window.ChatIntegration = ChatIntegration;