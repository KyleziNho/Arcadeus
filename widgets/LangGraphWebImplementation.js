/**
 * Real LangGraph Web Implementation for Excel Add-ins
 * Using @langchain/langgraph/web for browser compatibility
 */

// In a real implementation, these would be imported from:
// import { END, START, StateGraph, Annotation } from "@langchain/langgraph/web";

// For now, we'll create a simplified browser-compatible version
// that mimics the LangGraph web API structure

class LangGraphWebAnnotation {
  static Root(schema) {
    return {
      State: schema,
      create: (initialState) => ({ ...schema, ...initialState })
    };
  }

  static create(config) {
    return {
      reducer: config.reducer || ((x, y) => y),
      default: config.default || null
    };
  }
}

class LangGraphWebStateGraph {
  constructor(stateAnnotation) {
    this.stateAnnotation = stateAnnotation;
    this.nodes = new Map();
    this.edges = new Map();
    this.conditionalEdges = new Map();
    this.entryPoint = null;
  }

  addNode(name, func) {
    this.nodes.set(name, func);
    return this;
  }

  addEdge(from, to) {
    if (!this.edges.has(from)) {
      this.edges.set(from, []);
    }
    this.edges.get(from).push(to);
    return this;
  }

  addConditionalEdges(from, condition, mapping) {
    this.conditionalEdges.set(from, { condition, mapping });
    return this;
  }

  setEntryPoint(node) {
    this.entryPoint = node;
    return this;
  }

  compile(options = {}) {
    return new CompiledGraph(this, options);
  }
}

class CompiledGraph {
  constructor(graph, options) {
    this.graph = graph;
    this.options = options;
  }

  async invoke(initialState, config = {}) {
    let currentState = { ...initialState };
    let currentNode = this.graph.entryPoint || 'START';
    
    const maxSteps = 20;
    let steps = 0;
    
    while (currentNode !== 'END' && steps < maxSteps) {
      steps++;
      console.log(`🌟 LangGraph Step ${steps}: ${currentNode}`);
      
      // Execute current node
      if (this.graph.nodes.has(currentNode)) {
        const nodeFunc = this.graph.nodes.get(currentNode);
        const result = await nodeFunc(currentState);
        
        // Merge result into current state
        if (result) {
          currentState = { ...currentState, ...result };
        }
      }
      
      // Determine next node
      const nextNode = this.getNextNode(currentNode, currentState);
      currentNode = nextNode;
    }
    
    return currentState;
  }

  async *stream(initialState, config = {}) {
    let currentState = { ...initialState };
    let currentNode = this.graph.entryPoint || 'START';
    
    const maxSteps = 20;
    let steps = 0;
    
    while (currentNode !== 'END' && steps < maxSteps) {
      steps++;
      console.log(`🌊 LangGraph Stream Step ${steps}: ${currentNode}`);
      
      // Execute current node
      if (this.graph.nodes.has(currentNode)) {
        const nodeFunc = this.graph.nodes.get(currentNode);
        const result = await nodeFunc(currentState);
        
        // Merge result into current state
        if (result) {
          currentState = { ...currentState, ...result };
        }
        
        // Yield the current state after each node execution
        yield {
          node: currentNode,
          state: currentState,
          step: steps
        };
      }
      
      // Determine next node
      const nextNode = this.getNextNode(currentNode, currentState);
      currentNode = nextNode;
    }
    
    return currentState;
  }

  getNextNode(currentNode, state) {
    // Check conditional edges first
    if (this.graph.conditionalEdges.has(currentNode)) {
      const { condition, mapping } = this.graph.conditionalEdges.get(currentNode);
      const conditionResult = condition(state);
      
      if (mapping[conditionResult]) {
        return mapping[conditionResult];
      }
    }
    
    // Check regular edges
    if (this.graph.edges.has(currentNode)) {
      const edges = this.graph.edges.get(currentNode);
      if (edges.length > 0) {
        return edges[0];
      }
    }
    
    return 'END';
  }
}

// Excel-specific LangGraph implementation
class ExcelLangGraphWorkflow {
  constructor() {
    this.tools = new Map();
    this.initializeTools();
    this.buildWorkflow();
  }

  initializeTools() {
    // Load existing tools from LangChainProperTools
    if (window.LangChainProperTools) {
      window.LangChainProperTools.excelTools.forEach(tool => {
        this.tools.set(tool.name, tool);
      });
      console.log('✅ Loaded', this.tools.size, 'Excel tools for LangGraph Web');
    }
  }

  buildWorkflow() {
    // Define state schema using LangGraph Web pattern
    const ExcelState = LangGraphWebAnnotation.Root({
      messages: LangGraphWebAnnotation.create({
        reducer: (x, y) => x.concat(y),
        default: []
      }),
      userIntent: LangGraphWebAnnotation.create({ default: null }),
      toolResults: LangGraphWebAnnotation.create({ default: {} }),
      processingSteps: LangGraphWebAnnotation.create({
        reducer: (x, y) => x.concat(y),
        default: []
      }),
      needsClarification: LangGraphWebAnnotation.create({ default: false }),
      confidence: LangGraphWebAnnotation.create({ default: 0.0 })
    });

    // Create workflow
    this.workflow = new LangGraphWebStateGraph(ExcelState)
      .addNode('analyze_intent', this.analyzeIntent.bind(this))
      .addNode('select_tools', this.selectTools.bind(this))
      .addNode('execute_tools', this.executeTools.bind(this))
      .addNode('request_clarification', this.requestClarification.bind(this))
      .addNode('synthesize_response', this.synthesizeResponse.bind(this))
      .setEntryPoint('analyze_intent')
      .addConditionalEdges('analyze_intent', 
        (state) => state.confidence < 0.7 ? 'clarification' : 'continue',
        {
          'clarification': 'request_clarification',
          'continue': 'select_tools'
        }
      )
      .addEdge('select_tools', 'execute_tools')
      .addEdge('execute_tools', 'synthesize_response')
      .addEdge('request_clarification', 'END')
      .addEdge('synthesize_response', 'END');

    this.app = this.workflow.compile({
      checkpointer: null, // Could add memory here
      debug: true
    });
  }

  async analyzeIntent(state) {
    console.log('🧠 Node: analyze_intent');
    
    const lastMessage = state.messages[state.messages.length - 1];
    const userMessage = lastMessage?.content || '';
    
    // Enhanced intent analysis with confidence scoring
    const intent = this.analyzeUserIntentWithConfidence(userMessage);
    
    const processingStep = {
      node: 'analyze_intent',
      action: 'Intent Analysis',
      result: `Detected: ${intent.type} (${Math.round(intent.confidence * 100)}% confidence)`,
      success: true,
      timestamp: new Date().toISOString()
    };

    return {
      userIntent: intent,
      confidence: intent.confidence,
      processingSteps: [processingStep]
    };
  }

  async selectTools(state) {
    console.log('🔧 Node: select_tools');
    
    const selectedTools = this.selectToolsForIntent(state.userIntent);
    
    const processingStep = {
      node: 'select_tools',
      action: 'Tool Selection',
      result: `Selected: ${selectedTools.map(t => t.name).join(', ')}`,
      success: true,
      timestamp: new Date().toISOString()
    };

    return {
      selectedTools: selectedTools,
      processingSteps: [processingStep]
    };
  }

  async executeTools(state) {
    console.log('⚡ Node: execute_tools');
    
    const toolResults = {};
    const processingSteps = [];
    
    for (const toolSpec of state.selectedTools) {
      try {
        console.log(`🔧 Executing: ${toolSpec.name}`);
        
        const tool = this.tools.get(toolSpec.name);
        const result = await tool.call(toolSpec.args);
        const parsedResult = JSON.parse(result);
        
        toolResults[toolSpec.name] = parsedResult;
        
        processingSteps.push({
          node: 'execute_tools',
          action: `Tool: ${toolSpec.name}`,
          input: toolSpec.args,
          result: parsedResult.success ? 'Success' : 'Failed',
          details: parsedResult,
          success: parsedResult.success,
          timestamp: new Date().toISOString()
        });
        
      } catch (error) {
        console.error(`❌ Tool ${toolSpec.name} failed:`, error);
        
        toolResults[toolSpec.name] = { success: false, error: error.message };
        
        processingSteps.push({
          node: 'execute_tools',
          action: `Tool: ${toolSpec.name}`,
          result: 'Error',
          error: error.message,
          success: false,
          timestamp: new Date().toISOString()
        });
      }
    }
    
    return {
      toolResults: toolResults,
      processingSteps: processingSteps
    };
  }

  async requestClarification(state) {
    console.log('❓ Node: request_clarification');
    
    const clarificationMessage = {
      role: 'assistant',
      content: this.generateClarificationMessage(state.userIntent),
      type: 'clarification',
      needsUserInput: true,
      timestamp: new Date().toISOString()
    };

    return {
      messages: [clarificationMessage],
      needsClarification: true
    };
  }

  async synthesizeResponse(state) {
    console.log('📝 Node: synthesize_response');
    
    const response = this.synthesizeResponseFromResults(state);
    
    const responseMessage = {
      role: 'assistant',
      content: response,
      type: 'final_response',
      toolsUsed: Object.keys(state.toolResults || {}),
      timestamp: new Date().toISOString()
    };

    const processingStep = {
      node: 'synthesize_response',
      action: 'Response Generation',
      result: 'Generated comprehensive response',
      success: true,
      timestamp: new Date().toISOString()
    };

    return {
      messages: [responseMessage],
      processingSteps: [processingStep]
    };
  }

  analyzeUserIntentWithConfidence(userMessage) {
    const lowerInput = userMessage.toLowerCase();
    
    // Formatting intents with confidence
    if (lowerInput.includes('change') || lowerInput.includes('format') || lowerInput.includes('color')) {
      const searchTerms = this.extractFormattingTerms(userMessage);
      const confidence = searchTerms.searchTerm ? 0.9 : 0.6;
      
      return {
        type: 'formatting',
        confidence: confidence,
        description: 'User wants to format Excel cells',
        extractedTerms: searchTerms
      };
    }
    
    // Calculation intents
    if (lowerInput.includes('calculate') || lowerInput.includes('irr') || lowerInput.includes('npv')) {
      return {
        type: 'calculation',
        confidence: 0.85,
        description: 'User wants to perform calculations',
        extractedTerms: this.extractCalculationTerms(userMessage)
      };
    }
    
    // Search intents
    if (lowerInput.includes('find') || lowerInput.includes('where') || lowerInput.includes('show')) {
      return {
        type: 'search',
        confidence: 0.8,
        description: 'User wants to find data',
        extractedTerms: this.extractSearchTerms(userMessage)
      };
    }
    
    // Low confidence - needs clarification
    return {
      type: 'unclear',
      confidence: 0.3,
      description: 'Intent unclear - needs clarification',
      extractedTerms: {}
    };
  }

  selectToolsForIntent(intent) {
    const tools = [];
    
    if (intent.type === 'formatting' && intent.extractedTerms.searchTerm) {
      tools.push({
        name: 'smart_cell_formatting',
        args: {
          searchTerm: intent.extractedTerms.searchTerm,
          formatType: intent.extractedTerms.formatType || 'color',
          formatValue: intent.extractedTerms.formatValue || 'green',
          searchAllSheets: true
        }
      });
    } else if (intent.type === 'calculation') {
      tools.push({
        name: 'find_financial_metric',
        args: { metricName: intent.extractedTerms.metric || 'IRR' }
      });
    } else if (intent.type === 'search') {
      tools.push({
        name: 'find_financial_metric',
        args: { metricName: intent.extractedTerms.metric || 'All' }
      });
    }
    
    return tools;
  }

  extractFormattingTerms(input) {
    // Extract search term
    const searchPatterns = [
      /(?:the\s+)?(unlevered\s+irr|levered\s+irr|irr|moic|revenue|ebitda)/i,
    ];
    
    let searchTerm = '';
    for (const pattern of searchPatterns) {
      const match = input.match(pattern);
      if (match) {
        searchTerm = match[1].trim();
        break;
      }
    }
    
    // Extract color
    const colors = ['red', 'green', 'blue', 'yellow', 'orange', 'purple'];
    let formatValue = 'green';
    for (const color of colors) {
      if (input.toLowerCase().includes(color)) {
        formatValue = color;
        break;
      }
    }
    
    return { searchTerm, formatType: 'color', formatValue };
  }

  extractCalculationTerms(input) {
    const formulaMatch = input.match(/(irr|npv|pv|fv)\s*\([^)]+\)/i);
    return {
      formula: formulaMatch ? formulaMatch[0] : null,
      metric: input.toLowerCase().includes('irr') ? 'IRR' : 'NPV'
    };
  }

  extractSearchTerms(input) {
    const metrics = ['IRR', 'MOIC', 'Revenue', 'EBITDA'];
    const foundMetric = metrics.find(metric => 
      input.toLowerCase().includes(metric.toLowerCase()));
    return { metric: foundMetric || 'IRR' };
  }

  generateClarificationMessage(intent) {
    return `I'm ${Math.round(intent.confidence * 100)}% confident about your request. Could you please clarify what you'd like me to do? For example:
    
• **Format cells**: "Change the IRR cell to green"
• **Analyze data**: "Calculate the NPV for this model"  
• **Find values**: "Show me where the revenue is located"

What specifically would you like me to help with?`;
  }

  synthesizeResponseFromResults(state) {
    const intent = state.userIntent;
    const toolResults = state.toolResults || {};
    
    let response = `## ${this.getIntentTitle(intent)}\n\n`;
    
    // Add execution results
    if (Object.keys(toolResults).length > 0) {
      response += "### Results:\n";
      
      Object.entries(toolResults).forEach(([toolName, result]) => {
        if (result.success) {
          response += this.formatToolResult(toolName, result);
        } else {
          response += `❌ **${toolName}**: ${result.error}\n`;
        }
      });
    }
    
    return response;
  }

  getIntentTitle(intent) {
    const titles = {
      'formatting': '🎨 Cell Formatting Complete',
      'calculation': '📊 Calculation Results',
      'search': '🔍 Search Results',
      'analysis': '📈 Analysis Complete'
    };
    
    return titles[intent.type] || '📋 Operation Complete';
  }

  formatToolResult(toolName, result) {
    if (toolName === 'smart_cell_formatting') {
      return `✅ **Formatted ${result.cellsFormatted} cells** matching "${result.searchTerm}"\n`;
    } else if (toolName === 'find_financial_metric') {
      return `✅ **${result.metric}**: ${result.value} at ${result.location}\n`;
    }
    return `✅ **${toolName}**: Operation completed\n`;
  }

  // Public API methods
  async invoke(initialState) {
    return await this.app.invoke(initialState);
  }

  async *stream(initialState) {
    yield* this.app.stream(initialState);
  }
}

// Initialize globally for browser use
if (typeof window !== 'undefined') {
  window.LangGraphWebAnnotation = LangGraphWebAnnotation;
  window.LangGraphWebStateGraph = LangGraphWebStateGraph;
  window.ExcelLangGraphWorkflow = ExcelLangGraphWorkflow;
  
  console.log('✅ LangGraph Web Implementation initialized for browsers');
}