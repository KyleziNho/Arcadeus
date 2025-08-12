/**
 * LangChain ReAct Agent Implementation
 * Following expert plan for step-by-step reasoning and memory
 */

class BufferMemory {
  constructor(options = {}) {
    this.chatHistory = [];
    this.maxMessages = options.maxMessages || 20;
    this.returnMessages = options.returnMessages || true;
    this.memoryKey = options.memoryKey || "chat_history";
  }

  async saveContext(inputs, outputs) {
    // Add user input
    if (inputs.input) {
      this.chatHistory.push({
        type: "human",
        content: inputs.input,
        timestamp: new Date().toISOString()
      });
    }

    // Add assistant output
    if (outputs.output) {
      this.chatHistory.push({
        type: "ai", 
        content: outputs.output,
        timestamp: new Date().toISOString()
      });
    }

    // Keep memory within limits
    if (this.chatHistory.length > this.maxMessages) {
      this.chatHistory = this.chatHistory.slice(-this.maxMessages);
    }
  }

  async loadMemoryVariables(values) {
    return {
      [this.memoryKey]: this.returnMessages ? this.chatHistory : this.formatChatHistory()
    };
  }

  formatChatHistory() {
    return this.chatHistory.map(msg => 
      `${msg.type === 'human' ? 'Human' : 'Assistant'}: ${msg.content}`
    ).join('\n');
  }

  clear() {
    this.chatHistory = [];
  }
}

class ReActAgent {
  constructor(options) {
    this.llm = options.llm;
    this.tools = options.tools;
    this.memory = options.memory;
    this.maxIterations = options.maxIterations || 5;
    this.verbose = options.verbose || true;
    
    this.toolMap = new Map();
    this.tools.forEach(tool => {
      this.toolMap.set(tool.name, tool);
    });
  }

  async invoke(input) {
    console.log('🤖 ReAct Agent starting...');
    
    const steps = [];
    let finalAnswer = null;
    let iteration = 0;
    
    // Load memory context
    const memoryVariables = await this.memory.loadMemoryVariables({});
    
    // Initial thought
    let currentThought = `I need to help with: ${input.input}`;
    
    while (iteration < this.maxIterations && !finalAnswer) {
      iteration++;
      console.log(`📝 Iteration ${iteration}: ${currentThought}`);
      
      // Step 1: Think
      const thinkingResult = await this.think(input.input, memoryVariables, steps);
      
      if (thinkingResult.action === 'Final Answer') {
        finalAnswer = thinkingResult.answer;
        break;
      }
      
      // Step 2: Act (use tool)
      if (thinkingResult.action && this.toolMap.has(thinkingResult.action)) {
        const tool = this.toolMap.get(thinkingResult.action);
        const toolInput = thinkingResult.actionInput;
        
        console.log(`🔧 Using tool: ${thinkingResult.action}`, toolInput);
        
        try {
          const observation = await tool.call(toolInput);
          const parsedObservation = JSON.parse(observation);
          
          steps.push({
            action: {
              tool: thinkingResult.action,
              toolInput: toolInput,
              log: `Action: ${thinkingResult.action}\nAction Input: ${JSON.stringify(toolInput)}`
            },
            observation: parsedObservation,
            thought: currentThought
          });
          
          // Update thought for next iteration
          currentThought = `Based on the ${thinkingResult.action} result, I need to analyze further or provide the final answer.`;
          
        } catch (error) {
          console.error('Tool execution error:', error);
          steps.push({
            action: {
              tool: thinkingResult.action,
              toolInput: toolInput
            },
            observation: { error: error.message },
            thought: currentThought
          });
        }
      } else {
        // No valid action found, provide final answer
        finalAnswer = thinkingResult.answer || "I couldn't determine the appropriate action to take.";
      }
    }
    
    // Generate final response if we hit max iterations
    if (!finalAnswer) {
      finalAnswer = await this.generateFinalAnswer(input.input, steps);
    }
    
    const result = {
      output: finalAnswer,
      intermediate_steps: steps,
      iterations: iteration
    };
    
    // Save to memory
    await this.memory.saveContext({ input: input.input }, { output: finalAnswer });
    
    return result;
  }

  async think(userInput, memoryVariables, previousSteps) {
    // Simple ReAct-style thinking logic
    const availableTools = this.tools.map(tool => `${tool.name}: ${tool.description}`).join('\n');
    
    const context = memoryVariables.chat_history ? 
      memoryVariables.chat_history.slice(-4).map(msg => `${msg.type}: ${msg.content}`).join('\n') : '';
    
    const previousActions = previousSteps.map(step => 
      `Action: ${step.action.tool}\nObservation: ${JSON.stringify(step.observation)}`
    ).join('\n\n');
    
    // Determine what action to take based on the input
    const lowerInput = userInput.toLowerCase();
    
    // Financial analysis patterns
    if (lowerInput.includes('irr') || lowerInput.includes('npv') || lowerInput.includes('calculate')) {
      const formulaMatch = userInput.match(/(irr|npv|pv|fv)\s*\([^)]+\)/i);
      if (formulaMatch) {
        return {
          action: 'evaluate_financial_formula',
          actionInput: {
            formula: formulaMatch[0]
          }
        };
      } else {
        return {
          action: 'find_financial_metric',
          actionInput: {
            metricName: lowerInput.includes('irr') ? 'IRR' : 'NPV'
          }
        };
      }
    }
    
    // Reading data patterns
    if (lowerInput.includes('read') || lowerInput.includes('get') || lowerInput.includes('show')) {
      const rangeMatch = userInput.match(/([A-Z]+\d+:[A-Z]+\d+)/i);
      if (rangeMatch) {
        return {
          action: 'read_range',
          actionInput: {
            sheetName: 'Sheet1', // Default, could be smarter
            range: rangeMatch[0]
          }
        };
      }
    }
    
    // Metric finding patterns
    if (lowerInput.includes('find') || lowerInput.includes('locate') || lowerInput.includes('where')) {
      const metrics = ['IRR', 'MOIC', 'Revenue', 'EBITDA', 'NPV', 'Exit Value'];
      const foundMetric = metrics.find(metric => lowerInput.includes(metric.toLowerCase()));
      
      if (foundMetric) {
        return {
          action: 'find_financial_metric',
          actionInput: {
            metricName: foundMetric
          }
        };
      }
    }
    
    // If we have previous steps, analyze and provide final answer
    if (previousSteps.length > 0) {
      const analysisData = this.analyzeSteps(previousSteps);
      return {
        action: 'Final Answer',
        answer: this.formatFinalAnswer(userInput, analysisData)
      };
    }
    
    // Default: try to find relevant financial metrics
    return {
      action: 'find_financial_metric',
      actionInput: {
        metricName: 'All',
        searchAllSheets: true
      }
    };
  }

  analyzeSteps(steps) {
    const analysis = {
      toolsUsed: [],
      dataFound: {},
      errors: []
    };
    
    steps.forEach(step => {
      analysis.toolsUsed.push(step.action.tool);
      
      if (step.observation.success) {
        analysis.dataFound[step.action.tool] = step.observation;
      } else {
        analysis.errors.push(step.observation.error || 'Unknown error');
      }
    });
    
    return analysis;
  }

  formatFinalAnswer(userInput, analysisData) {
    let answer = "## Analysis Results\n\n";
    
    // Add found data
    if (Object.keys(analysisData.dataFound).length > 0) {
      answer += "### Key Findings:\n";
      
      Object.entries(analysisData.dataFound).forEach(([tool, data]) => {
        if (tool === 'find_financial_metric' && data.value) {
          answer += `• **${data.metric}**: ${data.value} at ${data.location}\n`;
        } else if (tool === 'evaluate_financial_formula') {
          answer += `• **Formula Result**: ${data.formattedValue}\n`;
          answer += `  - ${data.interpretation}\n`;
        } else if (tool === 'read_range') {
          answer += `• **Range Data**: Successfully read ${data.data.values.length} rows from ${data.data.address}\n`;
        }
      });
    }
    
    // Add errors if any
    if (analysisData.errors.length > 0) {
      answer += "\n### Issues Encountered:\n";
      analysisData.errors.forEach(error => {
        answer += `⚠️ ${error}\n`;
      });
    }
    
    // Add recommendations
    answer += "\n### Recommendations:\n";
    if (analysisData.dataFound.find_financial_metric) {
      const metric = analysisData.dataFound.find_financial_metric;
      if (metric.metric === 'IRR' && metric.rawValue) {
        const irr = metric.rawValue;
        if (irr > 0.2) {
          answer += "✅ Excellent IRR performance - consider similar investments\n";
        } else if (irr < 0.1) {
          answer += "⚠️ IRR below target - review assumptions and optimization opportunities\n";
        }
      }
    }
    
    return answer;
  }

  async generateFinalAnswer(userInput, steps) {
    if (steps.length === 0) {
      return "I wasn't able to find the specific information you requested. Please ensure your Excel workbook contains the relevant financial data.";
    }
    
    const analysisData = this.analyzeSteps(steps);
    return this.formatFinalAnswer(userInput, analysisData);
  }
}

class LangChainAgentExecutor {
  constructor(options) {
    this.agent = new ReActAgent({
      llm: options.llm,
      tools: options.tools,
      memory: options.memory,
      maxIterations: options.maxIterations || 5,
      verbose: options.verbose || true
    });
    
    this.tools = options.tools;
    this.memory = options.memory;
    this.verbose = options.verbose;
  }

  async invoke(input) {
    return await this.agent.invoke(input);
  }
}

// Initialize globally
if (typeof window !== 'undefined') {
  window.LangChainReActAgent = {
    BufferMemory,
    ReActAgent, 
    LangChainAgentExecutor
  };
  
  console.log('✅ LangChain ReAct Agent initialized');
}