/**
 * Deep Excel Agent
 * Implements Deep Agents architecture for intelligent Excel operations
 * Based on Deep Agents Python package concepts
 */

class DeepExcelAgent {
  constructor(apiKey) {
    this.apiKey = apiKey;
    this.fileSystem = new VirtualFileSystem();
    this.todoList = [];
    this.subAgents = new Map();
    this.conversationHistory = [];
    this.maxDepth = 5;
    this.currentDepth = 0;
    
    this.initializeSubAgents();
    this.initializeTools();
  }

  /**
   * System prompt based on Claude Code approach
   */
  getSystemPrompt() {
    return `You are an expert Excel analyst with deep reasoning capabilities. You work methodically through complex tasks using planning and sub-agents.

## Core Principles

1. **Always Plan Before Acting**: Break down every request into concrete steps using the planning tool
2. **Use Sub-Agents for Focused Work**: Delegate specialized tasks to sub-agents
3. **Maintain Context**: Use the file system to store intermediate results and maintain state
4. **Think Step by Step**: Work through your plan methodically, checking results at each step
5. **Be Thorough**: Don't rush - take time to understand the full context before acting

## Your Capabilities

### Planning Tool (todo_write)
You MUST use this tool at the start of every task to create a plan. Update it as you progress.
- Break complex tasks into 3-10 specific, actionable steps
- Mark steps as 'pending', 'in_progress', or 'completed'
- Re-evaluate and adjust your plan as needed

### File System Tools
Use these to maintain state and store results:
- write_file: Store intermediate results, analysis, or data
- read_file: Retrieve stored information
- edit_file: Modify existing files
- ls: List all files in your workspace

### Excel Tools
Comprehensive Excel operations for data manipulation:
- readExcelRange: Read data from Excel
- writeExcelRange: Write data to Excel
- formatExcelCells: Apply formatting
- findExcelCells: Search for specific cells
- analyzeWorkbook: Understand structure
- executeFormula: Run calculations

### Sub-Agents (create_subagent)
Create specialized agents for focused tasks:
- "research-agent": For gathering and analyzing data
- "formatting-agent": For visual improvements
- "analysis-agent": For complex calculations
- "general-agent": For any task with fresh context

## Workflow Pattern

1. **Receive Request** → Understand what the user wants
2. **Create Plan** → Use todo_write to break down the task
3. **Gather Context** → Read Excel data, understand current state
4. **Execute Plan** → Work through each step, using sub-agents as needed
5. **Store Results** → Save important findings to files
6. **Synthesize** → Combine all results into a coherent response
7. **Verify** → Check that you've actually accomplished the goal

## Important Instructions

- ALWAYS start with a plan using todo_write
- NEVER skip steps or rush through tasks
- USE sub-agents for complex sub-tasks to avoid context pollution
- STORE important information in files for later reference
- UPDATE your todo list as you complete steps
- THINK deeply about each step before executing

Remember: You're not just executing commands - you're solving problems intelligently.`;
  }

  /**
   * Initialize specialized sub-agents
   */
  initializeSubAgents() {
    this.subAgents.set('research-agent', {
      name: 'research-agent',
      description: 'Specialized in researching and analyzing Excel data patterns',
      prompt: `You are a research specialist focused on Excel data analysis. 
      Your job is to thoroughly analyze data, find patterns, and provide insights.
      Be methodical and detailed in your analysis.`,
      tools: ['readExcelRange', 'analyzeWorkbook', 'findExcelCells']
    });

    this.subAgents.set('formatting-agent', {
      name: 'formatting-agent',
      description: 'Expert at making Excel documents look professional',
      prompt: `You are a formatting expert. Your job is to make Excel documents 
      look professional, clean, and easy to read. Focus on visual hierarchy,
      consistent styling, and clear data presentation.`,
      tools: ['formatExcelCells', 'readExcelRange']
    });

    this.subAgents.set('analysis-agent', {
      name: 'analysis-agent',
      description: 'Specialized in complex Excel calculations and formulas',
      prompt: `You are an Excel formula and calculation expert. 
      Your job is to perform complex calculations, create formulas,
      and provide quantitative analysis.`,
      tools: ['executeFormula', 'writeExcelRange', 'readExcelRange']
    });

    this.subAgents.set('general-agent', {
      name: 'general-agent',
      description: 'General purpose agent with fresh context',
      prompt: 'You are a general purpose Excel assistant with all capabilities.',
      tools: null // Has access to all tools
    });
  }

  /**
   * Initialize all available tools
   */
  initializeTools() {
    this.tools = {
      // Planning tool
      todo_write: {
        name: 'todo_write',
        description: 'Create or update your task plan. ALWAYS use this first.',
        execute: this.todoWrite.bind(this)
      },
      
      // File system tools
      write_file: {
        name: 'write_file',
        description: 'Write content to a file for persistence',
        execute: this.writeFile.bind(this)
      },
      
      read_file: {
        name: 'read_file',
        description: 'Read content from a file',
        execute: this.readFile.bind(this)
      },
      
      edit_file: {
        name: 'edit_file',
        description: 'Edit an existing file',
        execute: this.editFile.bind(this)
      },
      
      ls: {
        name: 'ls',
        description: 'List all files in the workspace',
        execute: this.listFiles.bind(this)
      },
      
      // Sub-agent tool
      create_subagent: {
        name: 'create_subagent',
        description: 'Create a sub-agent for a specific task',
        execute: this.createSubAgent.bind(this)
      },
      
      // Excel tools (would integrate your UnifiedExcelApiLibrary here)
      readExcelRange: {
        name: 'readExcelRange',
        description: 'Read data from Excel',
        execute: async (args) => {
          if (window.UnifiedExcelApiLibrary) {
            const lib = new window.UnifiedExcelApiLibrary();
            return await lib.readExcelRange(args);
          }
          return { success: false, error: 'Excel API not available' };
        }
      },
      
      writeExcelRange: {
        name: 'writeExcelRange',
        description: 'Write data to Excel',
        execute: async (args) => {
          if (window.UnifiedExcelApiLibrary) {
            const lib = new window.UnifiedExcelApiLibrary();
            return await lib.writeExcelRange(args);
          }
          return { success: false, error: 'Excel API not available' };
        }
      },
      
      formatExcelCells: {
        name: 'formatExcelCells',
        description: 'Format Excel cells',
        execute: async (args) => {
          if (window.UnifiedExcelApiLibrary) {
            const lib = new window.UnifiedExcelApiLibrary();
            return await lib.formatExcelCells(args);
          }
          return { success: false, error: 'Excel API not available' };
        }
      },
      
      findExcelCells: {
        name: 'findExcelCells',
        description: 'Find cells in Excel',
        execute: async (args) => {
          if (window.UnifiedExcelApiLibrary) {
            const lib = new window.UnifiedExcelApiLibrary();
            return await lib.findExcelCells(args);
          }
          return { success: false, error: 'Excel API not available' };
        }
      },
      
      analyzeWorkbook: {
        name: 'analyzeWorkbook',
        description: 'Get comprehensive workbook structure',
        execute: async (args) => {
          if (window.UnifiedExcelApiLibrary) {
            const lib = new window.UnifiedExcelApiLibrary();
            return await lib.getWorkbookStructure(args || {});
          }
          return { success: false, error: 'Excel API not available' };
        }
      }
    };
  }

  /**
   * Main entry point - process request with deep reasoning
   */
  async processRequest(userInput) {
    console.log('🧠 Deep Excel Agent processing:', userInput);
    this.currentDepth = 0;
    
    try {
      // Build conversation with system prompt
      const messages = [
        { role: 'system', content: this.getSystemPrompt() },
        { role: 'user', content: userInput }
      ];

      // Add file system state to context
      if (this.fileSystem.files.size > 0) {
        messages[0].content += `\n\n## Current File System\n${this.fileSystem.listFiles()}`;
      }

      // Add current todo list to context
      if (this.todoList.length > 0) {
        messages[0].content += `\n\n## Current Todo List\n${JSON.stringify(this.todoList, null, 2)}`;
      }

      // Process with deep reasoning loop
      const response = await this.deepReasoningLoop(messages);
      
      return {
        success: true,
        response: response,
        todoList: this.todoList,
        files: Array.from(this.fileSystem.files.entries()),
        timestamp: new Date().toISOString()
      };

    } catch (error) {
      console.error('❌ Deep Excel Agent error:', error);
      return {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  }

  /**
   * Deep reasoning loop with tool execution
   */
  async deepReasoningLoop(messages) {
    const maxIterations = 15;
    let iterations = 0;
    let finalResponse = '';
    
    while (iterations < maxIterations) {
      iterations++;
      console.log(`🔄 Reasoning iteration ${iterations}`);
      
      // Call AI with current context
      const aiResponse = await this.callAI(messages);
      
      // Check if AI wants to use a tool
      if (aiResponse.function_call) {
        console.log(`🔧 AI requesting tool: ${aiResponse.function_call.name}`);
        
        // Execute the requested tool
        const toolResult = await this.executeTool(
          aiResponse.function_call.name,
          JSON.parse(aiResponse.function_call.arguments)
        );
        
        // Add tool result to conversation
        messages.push({
          role: 'assistant',
          content: aiResponse.content || '',
          function_call: aiResponse.function_call
        });
        
        messages.push({
          role: 'function',
          name: aiResponse.function_call.name,
          content: JSON.stringify(toolResult)
        });
        
        // Continue loop for more reasoning
        continue;
      }
      
      // AI provided final response
      finalResponse = aiResponse.content;
      break;
    }
    
    return finalResponse || 'Completed processing your request.';
  }

  /**
   * Call OpenAI with function calling
   */
  async callAI(messages) {
    const functions = Object.values(this.tools).map(tool => ({
      name: tool.name,
      description: tool.description,
      parameters: {
        type: 'object',
        properties: this.getToolParameters(tool.name)
      }
    }));

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${this.apiKey}`
      },
      body: JSON.stringify({
        model: 'gpt-4',
        messages: messages,
        functions: functions,
        function_call: 'auto',
        temperature: 0.1
      })
    });

    if (!response.ok) {
      throw new Error(`OpenAI API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0].message;
  }

  /**
   * Execute a tool
   */
  async executeTool(toolName, args) {
    const tool = this.tools[toolName];
    if (!tool) {
      return { success: false, error: `Tool ${toolName} not found` };
    }
    
    try {
      return await tool.execute(args);
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  /**
   * Planning tool - Create or update todo list
   */
  async todoWrite(args) {
    const { todos } = args;
    this.todoList = todos;
    
    console.log('📋 Todo list updated:');
    todos.forEach(todo => {
      const status = todo.status === 'completed' ? '✅' : 
                    todo.status === 'in_progress' ? '🔄' : '⏳';
      console.log(`${status} ${todo.id}. ${todo.task}`);
    });
    
    return {
      success: true,
      message: 'Todo list updated',
      todos: this.todoList
    };
  }

  /**
   * File system operations
   */
  async writeFile(args) {
    const { filename, content } = args;
    this.fileSystem.writeFile(filename, content);
    return { success: true, message: `File ${filename} written` };
  }

  async readFile(args) {
    const { filename } = args;
    const content = this.fileSystem.readFile(filename);
    if (content === null) {
      return { success: false, error: `File ${filename} not found` };
    }
    return { success: true, content: content };
  }

  async editFile(args) {
    const { filename, content } = args;
    if (!this.fileSystem.files.has(filename)) {
      return { success: false, error: `File ${filename} not found` };
    }
    this.fileSystem.writeFile(filename, content);
    return { success: true, message: `File ${filename} edited` };
  }

  async listFiles() {
    const files = this.fileSystem.listFiles();
    return { success: true, files: files };
  }

  /**
   * Create a sub-agent for specialized tasks
   */
  async createSubAgent(args) {
    const { agent_name, task, context } = args;
    
    console.log(`🤖 Creating sub-agent: ${agent_name} for task: ${task}`);
    
    const agentConfig = this.subAgents.get(agent_name) || this.subAgents.get('general-agent');
    
    // Create focused context for sub-agent
    const subAgentMessages = [
      { 
        role: 'system', 
        content: agentConfig.prompt + '\n\nContext from parent agent:\n' + (context || '')
      },
      { role: 'user', content: task }
    ];
    
    // Process with sub-agent (limited tools if specified)
    const response = await this.deepReasoningLoop(subAgentMessages);
    
    return {
      success: true,
      agent: agent_name,
      task: task,
      result: response
    };
  }

  /**
   * Get tool parameters based on tool name
   */
  getToolParameters(toolName) {
    const params = {
      todo_write: {
        todos: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              id: { type: 'number' },
              task: { type: 'string' },
              status: { type: 'string', enum: ['pending', 'in_progress', 'completed'] }
            }
          }
        }
      },
      write_file: {
        filename: { type: 'string' },
        content: { type: 'string' }
      },
      read_file: {
        filename: { type: 'string' }
      },
      edit_file: {
        filename: { type: 'string' },
        content: { type: 'string' }
      },
      ls: {},
      create_subagent: {
        agent_name: { type: 'string' },
        task: { type: 'string' },
        context: { type: 'string' }
      },
      readExcelRange: {
        sheetName: { type: 'string' },
        range: { type: 'string' }
      },
      writeExcelRange: {
        sheetName: { type: 'string' },
        range: { type: 'string' },
        values: { type: 'array' }
      },
      formatExcelCells: {
        cells: { type: 'array' },
        formatting: { type: 'object' }
      },
      findExcelCells: {
        searchTerm: { type: 'string' },
        searchType: { type: 'string' }
      },
      analyzeWorkbook: {}
    };
    
    return params[toolName] || {};
  }
}

/**
 * Virtual File System for persistence
 */
class VirtualFileSystem {
  constructor() {
    this.files = new Map();
  }

  writeFile(filename, content) {
    this.files.set(filename, content);
    console.log(`💾 File written: ${filename}`);
  }

  readFile(filename) {
    return this.files.get(filename) || null;
  }

  deleteFile(filename) {
    return this.files.delete(filename);
  }

  listFiles() {
    const fileList = Array.from(this.files.keys());
    return fileList.length > 0 ? fileList.join('\n') : 'No files in workspace';
  }

  clear() {
    this.files.clear();
  }
}

// Initialize globally
if (typeof window !== 'undefined') {
  window.DeepExcelAgent = DeepExcelAgent;
  window.VirtualFileSystem = VirtualFileSystem;
  console.log('✅ Deep Excel Agent initialized with planning, sub-agents, and file system');
}