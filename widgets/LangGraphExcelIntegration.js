/**
 * LangGraph Excel Integration
 * Modern graph-based AI workflow for Excel interactions
 */

// LangGraph-inspired State Management for JavaScript
class ExcelChatState {
  constructor() {
    this.messages = [];
    this.toolResults = {};
    this.userIntent = null;
    this.currentStep = 'start';
    this.excelContext = null;
    this.processingSteps = [];
  }

  addMessage(message) {
    this.messages.push({
      ...message,
      timestamp: new Date().toISOString(),
      id: this.generateId()
    });
  }

  updateToolResults(toolName, result) {
    this.toolResults[toolName] = result;
  }

  addProcessingStep(step) {
    this.processingSteps.push({
      ...step,
      timestamp: new Date().toISOString(),
      stepNumber: this.processingSteps.length + 1
    });
  }

  generateId() {
    return 'msg_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
  }

  getRecentMessages(count = 5) {
    return this.messages.slice(-count);
  }

  serialize() {
    return {
      messages: this.messages,
      toolResults: this.toolResults,
      userIntent: this.userIntent,
      currentStep: this.currentStep,
      excelContext: this.excelContext,
      processingSteps: this.processingSteps
    };
  }

  static deserialize(data) {
    const state = new ExcelChatState();
    Object.assign(state, data);
    return state;
  }
}

class LangGraphExcelWorkflow {
  constructor() {
    this.nodes = new Map();
    this.edges = new Map();
    this.tools = new Map();
    this.currentState = null;
    
    this.setupNodes();
    this.setupEdges();
    this.loadTools();
  }

  setupNodes() {
    // Node 1: Analyze user intent
    this.addNode('analyze_intent', async (state) => {
      console.log('🧠 Node: analyze_intent');
      
      const userMessage = state.messages[state.messages.length - 1];
      const intent = this.analyzeUserIntent(userMessage.content);
      
      state.userIntent = intent;
      state.addProcessingStep({
        node: 'analyze_intent',
        action: 'Intent Analysis',
        result: `Detected intent: ${intent.type} - ${intent.description}`,
        success: true
      });
      
      return state;
    });

    // Node 2: Select appropriate tools
    this.addNode('select_tools', async (state) => {
      console.log('🔧 Node: select_tools');
      
      const toolsToUse = this.selectToolsForIntent(state.userIntent);
      
      state.selectedTools = toolsToUse;
      state.addProcessingStep({
        node: 'select_tools',
        action: 'Tool Selection',
        result: `Selected tools: ${toolsToUse.map(t => t.name).join(', ')}`,
        success: true
      });
      
      return state;
    });

    // Node 3: Execute tools
    this.addNode('execute_tools', async (state) => {
      console.log('⚡ Node: execute_tools');
      
      const toolResults = {};
      
      for (const toolSpec of state.selectedTools) {
        try {
          console.log(`🔧 Executing tool: ${toolSpec.name}`);
          const tool = this.tools.get(toolSpec.name);
          const result = await tool.call(toolSpec.args);
          const parsedResult = JSON.parse(result);
          
          toolResults[toolSpec.name] = parsedResult;
          
          state.addProcessingStep({
            node: 'execute_tools',
            action: `Tool: ${toolSpec.name}`,
            input: toolSpec.args,
            result: parsedResult.success ? 'Success' : 'Failed',
            details: parsedResult,
            success: parsedResult.success
          });
          
        } catch (error) {
          console.error(`❌ Tool ${toolSpec.name} failed:`, error);
          toolResults[toolSpec.name] = { success: false, error: error.message };
          
          state.addProcessingStep({
            node: 'execute_tools',
            action: `Tool: ${toolSpec.name}`,
            result: 'Error',
            error: error.message,
            success: false
          });
        }
      }
      
      state.updateToolResults('execution_batch', toolResults);
      return state;
    });

    // Node 4: Synthesize response
    this.addNode('synthesize_response', async (state) => {
      console.log('📝 Node: synthesize_response');
      
      const response = this.synthesizeResponse(state);
      
      state.addMessage({
        role: 'assistant',
        content: response,
        type: 'final_response',
        toolsUsed: Object.keys(state.toolResults.execution_batch || {}),
        processingSteps: state.processingSteps.length
      });
      
      state.addProcessingStep({
        node: 'synthesize_response',
        action: 'Response Generation',
        result: 'Generated comprehensive response',
        success: true
      });
      
      return state;
    });

    // Node 5: Error handling
    this.addNode('handle_error', async (state) => {
      console.log('⚠️ Node: handle_error');
      
      const errorMessage = this.generateErrorResponse(state);
      
      state.addMessage({
        role: 'assistant',
        content: errorMessage,
        type: 'error_response'
      });
      
      return state;
    });
  }

  setupEdges() {
    // Linear flow with error handling
    this.addEdge('START', 'analyze_intent');
    this.addEdge('analyze_intent', 'select_tools');
    this.addEdge('select_tools', 'execute_tools');
    this.addEdge('execute_tools', 'synthesize_response');
    this.addEdge('synthesize_response', 'END');
    
    // Error handling edges (simplified)
    this.addConditionalEdge('execute_tools', (state) => {
      const hasErrors = Object.values(state.toolResults.execution_batch || {})
        .some(result => !result.success);
      return hasErrors ? 'handle_error' : 'synthesize_response';
    });
  }

  loadTools() {
    // Load existing tools from LangChainProperTools
    if (window.LangChainProperTools) {
      window.LangChainProperTools.excelTools.forEach(tool => {
        this.tools.set(tool.name, tool);
      });
      console.log('✅ Loaded', this.tools.size, 'Excel tools for LangGraph');
    }
  }

  addNode(name, func) {
    this.nodes.set(name, func);
  }

  addEdge(from, to) {
    if (!this.edges.has(from)) {
      this.edges.set(from, []);
    }
    this.edges.get(from).push({ type: 'direct', to });
  }

  addConditionalEdge(from, condition) {
    if (!this.edges.has(from)) {
      this.edges.set(from, []);
    }
    this.edges.get(from).push({ type: 'conditional', condition });
  }

  analyzeUserIntent(userInput) {
    const lowerInput = userInput.toLowerCase();
    
    // Formatting intents
    if (lowerInput.includes('change') || lowerInput.includes('format') || 
        lowerInput.includes('color') || lowerInput.includes('bold')) {
      return {
        type: 'formatting',
        description: 'User wants to format Excel cells',
        confidence: 0.9,
        extractedTerms: this.extractFormattingTerms(userInput)
      };
    }
    
    // Analysis intents
    if (lowerInput.includes('calculate') || lowerInput.includes('irr') || 
        lowerInput.includes('npv') || lowerInput.includes('formula')) {
      return {
        type: 'calculation',
        description: 'User wants to perform financial calculations',
        confidence: 0.8,
        extractedTerms: this.extractCalculationTerms(userInput)
      };
    }
    
    // Search intents
    if (lowerInput.includes('find') || lowerInput.includes('where') || 
        lowerInput.includes('locate') || lowerInput.includes('show')) {
      return {
        type: 'search',
        description: 'User wants to find data in Excel',
        confidence: 0.7,
        extractedTerms: this.extractSearchTerms(userInput)
      };
    }
    
    // Default: general analysis
    return {
      type: 'analysis',
      description: 'General Excel data analysis request',
      confidence: 0.5,
      extractedTerms: {}
    };
  }

  extractFormattingTerms(input) {
    const searchPatterns = [
      /(?:the\s+)?(unlevered\s+irr|levered\s+irr|irr|moic|revenue|ebitda|exit\s+value)/i,
    ];
    
    const colors = ['red', 'green', 'blue', 'yellow', 'orange', 'purple', 'pink'];
    
    let searchTerm = '';
    let formatType = 'color';
    let formatValue = 'green';
    
    for (const pattern of searchPatterns) {
      const match = input.match(pattern);
      if (match) {
        searchTerm = match[1].trim();
        break;
      }
    }
    
    for (const color of colors) {
      if (input.toLowerCase().includes(color)) {
        formatValue = color;
        break;
      }
    }
    
    if (input.toLowerCase().includes('bold')) formatType = 'bold';
    
    return { searchTerm, formatType, formatValue };
  }

  extractCalculationTerms(input) {
    const formulaMatch = input.match(/(irr|npv|pv|fv)\s*\([^)]+\)/i);
    return {
      formula: formulaMatch ? formulaMatch[0] : null,
      metric: input.toLowerCase().includes('irr') ? 'IRR' : 'NPV'
    };
  }

  extractSearchTerms(input) {
    const metrics = ['IRR', 'MOIC', 'Revenue', 'EBITDA', 'NPV', 'Exit Value'];
    const foundMetric = metrics.find(metric => 
      input.toLowerCase().includes(metric.toLowerCase()));
    
    return { metric: foundMetric || 'IRR' };
  }

  selectToolsForIntent(intent) {
    const toolsToUse = [];
    
    switch (intent.type) {
      case 'formatting':
        if (intent.extractedTerms.searchTerm) {
          toolsToUse.push({
            name: 'smart_cell_formatting',
            args: {
              searchTerm: intent.extractedTerms.searchTerm,
              formatType: intent.extractedTerms.formatType,
              formatValue: intent.extractedTerms.formatValue,
              searchAllSheets: true
            }
          });
        }
        break;
        
      case 'calculation':
        if (intent.extractedTerms.formula) {
          toolsToUse.push({
            name: 'evaluate_financial_formula',
            args: { formula: intent.extractedTerms.formula }
          });
        } else {
          toolsToUse.push({
            name: 'find_financial_metric',
            args: { metricName: intent.extractedTerms.metric }
          });
        }
        break;
        
      case 'search':
        toolsToUse.push({
          name: 'find_financial_metric',
          args: { metricName: intent.extractedTerms.metric }
        });
        break;
        
      default:
        // General analysis - try to find financial metrics
        toolsToUse.push({
          name: 'find_financial_metric',
          args: { metricName: 'IRR' }
        });
        break;
    }
    
    return toolsToUse;
  }

  synthesizeResponse(state) {
    const userIntent = state.userIntent;
    const toolResults = state.toolResults.execution_batch || {};
    
    let response = `## ${this.getIntentTitle(userIntent)}\n\n`;
    
    // Add tool results
    if (Object.keys(toolResults).length > 0) {
      response += "### Execution Results:\n";
      
      Object.entries(toolResults).forEach(([toolName, result]) => {
        if (result.success) {
          response += this.formatToolResult(toolName, result);
        } else {
          response += `❌ **${toolName}**: ${result.error}\n`;
        }
      });
    }
    
    // Add insights
    response += "\n### Analysis:\n";
    response += this.generateInsights(userIntent, toolResults);
    
    return response;
  }

  getIntentTitle(intent) {
    const titles = {
      'formatting': '🎨 Cell Formatting Complete',
      'calculation': '📊 Financial Calculation Results',
      'search': '🔍 Excel Data Search Results',
      'analysis': '📈 Excel Data Analysis'
    };
    
    return titles[intent.type] || '📋 Excel Operation Complete';
  }

  formatToolResult(toolName, result) {
    switch (toolName) {
      case 'smart_cell_formatting':
        return `✅ **Formatting Applied**: Found ${result.cellsFound} cells matching "${result.searchTerm}" and successfully formatted ${result.cellsFormatted} cells\n`;
        
      case 'find_financial_metric':
        return `✅ **${result.metric}**: ${result.value} at ${result.location}\n`;
        
      case 'evaluate_financial_formula':
        return `✅ **Formula Result**: ${result.formattedValue} - ${result.interpretation}\n`;
        
      default:
        return `✅ **${toolName}**: Operation completed successfully\n`;
    }
  }

  generateInsights(intent, toolResults) {
    if (intent.type === 'formatting') {
      return "💡 Cells have been formatted based on your specifications. The formatting was applied to cells containing the exact terms you requested.";
    }
    
    return "💡 Operation completed successfully. The results above show the current state of your Excel data.";
  }

  generateErrorResponse(state) {
    return "⚠️ I encountered some issues processing your request. Please check that your Excel workbook contains the expected data and try again.";
  }

  // Main execution method
  async invoke(initialState) {
    console.log('🚀 Starting LangGraph Excel Workflow');
    
    let currentState = initialState;
    let currentNode = 'analyze_intent';
    
    const maxSteps = 10;
    let steps = 0;
    
    while (currentNode !== 'END' && steps < maxSteps) {
      steps++;
      console.log(`📍 Step ${steps}: Executing node '${currentNode}'`);
      
      // Execute current node
      if (this.nodes.has(currentNode)) {
        const nodeFunc = this.nodes.get(currentNode);
        currentState = await nodeFunc(currentState);
        currentState.currentStep = currentNode;
      }
      
      // Determine next node
      const edges = this.edges.get(currentNode) || [];
      let nextNode = 'END';
      
      for (const edge of edges) {
        if (edge.type === 'direct') {
          nextNode = edge.to;
          break;
        } else if (edge.type === 'conditional') {
          nextNode = edge.condition(currentState);
          break;
        }
      }
      
      currentNode = nextNode;
    }
    
    console.log('✅ LangGraph workflow completed');
    return currentState;
  }

  // Stream execution for real-time updates
  async *stream(initialState) {
    console.log('🌊 Starting LangGraph streaming execution');
    
    let currentState = initialState;
    let currentNode = 'analyze_intent';
    
    while (currentNode !== 'END') {
      // Execute node and yield state
      if (this.nodes.has(currentNode)) {
        const nodeFunc = this.nodes.get(currentNode);
        currentState = await nodeFunc(currentState);
        
        yield {
          node: currentNode,
          state: currentState,
          step: currentState.processingSteps[currentState.processingSteps.length - 1]
        };
      }
      
      // Get next node
      const edges = this.edges.get(currentNode) || [];
      currentNode = edges[0]?.to || 'END';
    }
  }
}

// Initialize globally
if (typeof window !== 'undefined') {
  window.LangGraphExcelWorkflow = LangGraphExcelWorkflow;
  window.ExcelChatState = ExcelChatState;
  
  console.log('✅ LangGraph Excel Workflow initialized');
}