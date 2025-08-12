/**
 * Unified AI Agent
 * Single intelligent agent that handles all user requests using OpenAI + Excel APIs
 */

class UnifiedAiAgent {
  constructor(apiKey) {
    this.apiKey = apiKey || this.getApiKeyFromStorage();
    this.excelApi = new UnifiedExcelApiLibrary();
    this.conversationHistory = [];
    this.excelContext = null;
    this.lastContextUpdate = null;
  }

  /**
   * Get API key from storage or environment
   */
  getApiKeyFromStorage() {
    // Try to get from localStorage first
    let apiKey = localStorage.getItem('openai_api_key');
    
    // If not found, try sessionStorage
    if (!apiKey) {
      apiKey = sessionStorage.getItem('openai_api_key');
    }

    // If still not found, check if it's in a global variable
    if (!apiKey && typeof window.OPENAI_API_KEY !== 'undefined') {
      apiKey = window.OPENAI_API_KEY;
    }

    return apiKey;
  }

  /**
   * Set API key
   */
  setApiKey(apiKey) {
    this.apiKey = apiKey;
    localStorage.setItem('openai_api_key', apiKey);
  }

  /**
   * Main entry point - process any user request
   */
  async processRequest(userInput) {
    console.log('🧠 UnifiedAiAgent processing:', userInput);

    try {
      // 1. Update Excel context if needed
      await this.updateExcelContext();

      // 2. Prepare messages with context
      const messages = await this.prepareMessages(userInput);

      // 3. Call OpenAI with function calling
      const response = await this.callOpenAiWithFunctions(messages);

      // 4. Execute any function calls
      const finalResponse = await this.handleFunctionCalls(response);

      // 5. Update conversation history
      this.updateConversationHistory(userInput, finalResponse);

      return {
        success: true,
        response: finalResponse,
        timestamp: new Date().toISOString()
      };

    } catch (error) {
      console.error('❌ UnifiedAiAgent error:', error);
      return {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  }

  /**
   * Update Excel context for AI awareness
   */
  async updateExcelContext() {
    // Only update context every 30 seconds to avoid spam
    const now = Date.now();
    if (this.lastContextUpdate && (now - this.lastContextUpdate) < 30000) {
      return;
    }

    try {
      console.log('📊 Updating Excel context...');
      const structure = await this.excelApi.executeFunction('getWorkbookStructure', {
        includeData: true,
        maxRows: 10
      });

      if (structure.success) {
        this.excelContext = structure.structure;
        this.lastContextUpdate = now;
        console.log('✅ Excel context updated:', this.excelContext.totalSheets, 'sheets');
      }
    } catch (error) {
      console.error('⚠️ Failed to update Excel context:', error);
    }
  }

  /**
   * Prepare messages with system prompt and context
   */
  async prepareMessages(userInput) {
    const systemPrompt = `You are an intelligent Excel assistant. You can read, analyze, and modify Excel workbooks to help users with any request.

CURRENT EXCEL CONTEXT:
${this.excelContext ? JSON.stringify(this.excelContext, null, 2) : 'No Excel context available'}

CAPABILITIES:
- Read data from any Excel range
- Write/modify Excel data  
- Format cells (colors, fonts, borders, alignment)
- Find cells by content, color, or other criteria
- Analyze data and provide insights
- Calculate formulas and perform computations
- Identify headers and data structures
- Make Excel documents look professional

INSTRUCTIONS:
1. Understand what the user wants (analysis, modification, formatting, etc.)
2. Use Excel functions to gather any needed information
3. Take appropriate actions using available Excel functions
4. Provide clear, helpful responses with specific cell references
5. Be smart about interpreting user requests - use context to understand what they mean
6. If you need to find something like "blue headers" or "revenue numbers", use the search functions first

Available Excel Functions: ${this.excelApi.getOpenAiFunctions().map(f => f.name).join(', ')}

Be helpful, accurate, and contextual in your responses.`;

    const messages = [
      { role: "system", content: systemPrompt },
      ...this.conversationHistory.slice(-4), // Keep last 4 exchanges for context
      { role: "user", content: userInput }
    ];

    return messages;
  }

  /**
   * Call OpenAI with function calling capabilities
   */
  async callOpenAiWithFunctions(messages) {
    if (!this.apiKey) {
      throw new Error('OpenAI API key not configured. Please set your API key.');
    }

    const requestBody = {
      model: "gpt-4",
      messages: messages,
      functions: this.excelApi.getOpenAiFunctions(),
      function_call: "auto",
      temperature: 0.1 // Lower temperature for more consistent Excel operations
    };

    console.log('🤖 Calling OpenAI API...');
    
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${this.apiKey}`
      },
      body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(`OpenAI API error (${response.status}): ${errorData.error?.message || response.statusText}`);
    }

    const data = await response.json();
    return data;
  }

  /**
   * Handle function calls from OpenAI response
   */
  async handleFunctionCalls(openAiResponse) {
    const message = openAiResponse.choices[0].message;

    // If no function call, return the text response
    if (!message.function_call) {
      return message.content;
    }

    console.log('🔧 AI requested function:', message.function_call.name);

    try {
      // Execute the requested function
      const functionName = message.function_call.name;
      const functionArgs = JSON.parse(message.function_call.arguments);
      
      console.log('📝 Function arguments:', functionArgs);

      const functionResult = await this.excelApi.executeFunction(functionName, functionArgs);
      
      console.log('✅ Function result:', functionResult);

      // Send function result back to AI for interpretation
      const followUpMessages = [
        ...openAiResponse.choices[0].message.content ? [{
          role: "assistant",
          content: message.content,
          function_call: message.function_call
        }] : [],
        {
          role: "function",
          name: functionName,
          content: JSON.stringify(functionResult)
        }
      ];

      // Get AI's interpretation of the results
      const interpretationResponse = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${this.apiKey}`
        },
        body: JSON.stringify({
          model: "gpt-4",
          messages: [
            ...this.prepareMessages(""),  // Get system context
            ...followUpMessages
          ],
          temperature: 0.1
        })
      });

      if (!interpretationResponse.ok) {
        throw new Error(`Failed to get AI interpretation: ${interpretationResponse.statusText}`);
      }

      const interpretationData = await interpretationResponse.json();
      return interpretationData.choices[0].message.content;

    } catch (error) {
      console.error('❌ Function execution error:', error);
      return `I encountered an error while trying to execute the Excel operation: ${error.message}. Please check your Excel workbook and try again.`;
    }
  }

  /**
   * Update conversation history
   */
  updateConversationHistory(userInput, aiResponse) {
    this.conversationHistory.push(
      { role: "user", content: userInput },
      { role: "assistant", content: aiResponse }
    );

    // Keep only last 10 messages (5 exchanges)
    if (this.conversationHistory.length > 10) {
      this.conversationHistory = this.conversationHistory.slice(-10);
    }
  }

  /**
   * Check if agent is ready
   */
  isReady() {
    return !!(this.apiKey && this.excelApi);
  }

  /**
   * Get current status
   */
  getStatus() {
    return {
      ready: this.isReady(),
      hasApiKey: !!this.apiKey,
      hasExcelContext: !!this.excelContext,
      conversationLength: this.conversationHistory.length,
      lastContextUpdate: this.lastContextUpdate ? new Date(this.lastContextUpdate).toISOString() : null
    };
  }

  /**
   * Clear conversation history
   */
  clearHistory() {
    this.conversationHistory = [];
    console.log('🧹 Conversation history cleared');
  }

  /**
   * Force refresh Excel context
   */
  async refreshContext() {
    this.lastContextUpdate = null;
    await this.updateExcelContext();
  }
}

// API Key Configuration Helper
class ApiKeyManager {
  static showApiKeyPrompt() {
    const apiKey = prompt(`
🔑 OpenAI API Key Required

To use the intelligent Excel assistant, please enter your OpenAI API key.
You can get one from: https://platform.openai.com/api-keys

Your API key will be stored locally and only used for Excel operations.

Enter your API key:`);

    if (apiKey && apiKey.trim()) {
      return apiKey.trim();
    }
    
    return null;
  }

  static async ensureApiKey() {
    const agent = new UnifiedAiAgent();
    
    if (!agent.apiKey) {
      const apiKey = this.showApiKeyPrompt();
      if (apiKey) {
        agent.setApiKey(apiKey);
        return agent;
      } else {
        throw new Error('OpenAI API key is required for AI assistant functionality');
      }
    }
    
    return agent;
  }
}

// Initialize globally
if (typeof window !== 'undefined') {
  window.UnifiedAiAgent = UnifiedAiAgent;
  window.ApiKeyManager = ApiKeyManager;
  console.log('✅ Unified AI Agent initialized');
}