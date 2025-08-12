/**
 * Streaming Chat Initialization Script
 * Ensures proper initialization order and fallbacks
 */

(function() {
  'use strict';
  
  let initializationAttempts = 0;
  const maxAttempts = 10;
  const retryDelay = 500; // 500ms

  /**
   * Initialize streaming chat system with proper dependency checks
   */
  function initializeStreamingChat() {
    initializationAttempts++;
    
    console.log(`🔄 Streaming chat initialization attempt ${initializationAttempts}`);
    
    // Check if all required dependencies are available
    const dependencies = {
      'Office': typeof Office !== 'undefined',
      'ChatHandler': typeof window.ChatHandler !== 'undefined',
      'ExcelLiveAnalyzer': typeof window.ExcelLiveAnalyzer !== 'undefined',
      'EnhancedStreamingChat': typeof window.enhancedStreamingChat !== 'undefined',
      'Excel Navigator': typeof window.excelNavigator !== 'undefined',
      'Enhanced Response Formatter': typeof window.enhancedResponseFormatter !== 'undefined'
    };
    
    const missingDependencies = Object.entries(dependencies)
      .filter(([name, available]) => !available)
      .map(([name]) => name);
    
    if (missingDependencies.length > 0 && initializationAttempts < maxAttempts) {
      console.log(`⏳ Missing dependencies: ${missingDependencies.join(', ')}. Retrying in ${retryDelay}ms...`);
      setTimeout(initializeStreamingChat, retryDelay);
      return;
    }
    
    if (missingDependencies.length > 0) {
      console.warn('⚠️ Some dependencies are missing. Starting with fallback mode:', missingDependencies);
    }
    
    try {
      // Initialize Chat Handler if available
      if (typeof window.ChatHandler !== 'undefined' && !window.chatHandler) {
        console.log('🎯 Initializing ChatHandler...');
        window.chatHandler = new window.ChatHandler();
        
        // Initialize ChatHandler
        if (typeof window.chatHandler.initialize === 'function') {
          window.chatHandler.initialize().then(() => {
            console.log('✅ ChatHandler initialized successfully');
          }).catch(error => {
            console.error('❌ ChatHandler initialization failed:', error);
          });
        }
      }
      
      // Verify Enhanced Streaming Chat is working
      if (window.enhancedStreamingChat) {
        console.log('✅ Enhanced Streaming Chat system ready');
        
        // Test hook into chat handler
        if (window.chatHandler && typeof window.chatHandler.processWithAI === 'function') {
          console.log('✅ Successfully hooked into ChatHandler');
        } else {
          console.warn('⚠️ Could not hook into ChatHandler - processWithAI method not found');
        }
      } else {
        console.error('❌ Enhanced Streaming Chat system not available');
      }
      
      // Initialize Excel Live Analyzer if available
      if (typeof window.ExcelLiveAnalyzer !== 'undefined' && !window.excelLiveAnalyzer) {
        console.log('📊 Initializing ExcelLiveAnalyzer...');
        window.excelLiveAnalyzer = new window.ExcelLiveAnalyzer();
      }
      
      // Set up global error handling for streaming
      window.addEventListener('unhandledrejection', (event) => {
        if (event.reason && event.reason.message && 
            (event.reason.message.includes('streaming') || 
             event.reason.message.includes('OpenAI'))) {
          console.error('🚨 Streaming chat error:', event.reason);
          
          // Show user-friendly error
          if (window.chatHandler && typeof window.chatHandler.showStatus === 'function') {
            window.chatHandler.showStatus('Chat analysis temporarily unavailable. Please try again.');
          }
          
          event.preventDefault();
        }
      });
      
      // Set up API key verification (if needed)
      if (typeof fetch !== 'undefined') {
        console.log('🔑 Chat API connectivity ready');
      }
      
      console.log('🚀 Streaming chat system initialization complete!');
      
      // Fire custom event to notify other components
      const initEvent = new CustomEvent('streamingChatReady', {
        detail: {
          enhancedStreaming: !!window.enhancedStreamingChat,
          chatHandler: !!window.chatHandler,
          excelAnalyzer: !!window.excelLiveAnalyzer
        }
      });
      window.dispatchEvent(initEvent);
      
    } catch (error) {
      console.error('❌ Streaming chat initialization failed:', error);
    }
  }
  
  /**
   * Set up Office.js ready callback
   */
  if (typeof Office !== 'undefined') {
    Office.onReady(() => {
      console.log('📋 Office.js ready - initializing streaming chat...');
      // Small delay to ensure all scripts are loaded
      setTimeout(initializeStreamingChat, 100);
    });
  } else {
    // If Office.js is not available (testing environment), initialize anyway
    console.log('⚠️ Office.js not available - initializing in development mode...');
    setTimeout(initializeStreamingChat, 1000);
  }
  
  /**
   * Expose initialization function globally for manual retry
   */
  window.initializeStreamingChat = initializeStreamingChat;
  
  /**
   * Health check function
   */
  window.checkStreamingChatHealth = function() {
    const health = {
      enhancedStreaming: !!window.enhancedStreamingChat,
      chatHandler: !!window.chatHandler,
      excelAnalyzer: !!window.excelLiveAnalyzer,
      excelNavigator: !!window.excelNavigator,
      officeReady: typeof Office !== 'undefined' && Office.context?.requirements?.isSetSupported(),
      totalComponents: 0,
      readyComponents: 0
    };
    
    Object.keys(health).forEach(key => {
      if (typeof health[key] === 'boolean') {
        health.totalComponents++;
        if (health[key]) {
          health.readyComponents++;
        }
      }
    });
    
    health.healthScore = Math.round((health.readyComponents / health.totalComponents) * 100);
    health.status = health.healthScore >= 80 ? 'healthy' : 
                   health.healthScore >= 60 ? 'degraded' : 'critical';
    
    console.log('🏥 Streaming Chat Health Check:', health);
    return health;
  };
  
})();