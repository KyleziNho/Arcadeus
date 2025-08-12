/**
 * Intelligent Excel Agent
 * Uses advanced reasoning to plan and execute Excel operations
 */

class IntelligentExcelAgent {
  constructor() {
    this.tools = new Map();
    this.loadTools();
  }

  loadTools() {
    // Load basic tools
    if (window.LangChainProperTools) {
      window.LangChainProperTools.excelTools.forEach(tool => {
        this.tools.set(tool.name, tool);
      });
    }

    // Load advanced Excel API tools
    if (window.ExcelApiToolkit) {
      window.ExcelApiToolkit.excelApiTools.forEach(tool => {
        this.tools.set(tool.name, tool);
      });
    }

    console.log('🧠 IntelligentExcelAgent loaded', this.tools.size, 'tools');
  }

  /**
   * Main entry point - analyze user request and execute Excel operations
   */
  async processRequest(userInput) {
    console.log('🧠 IntelligentExcelAgent processing:', userInput);

    try {
      // Step 1: Analyze intent and rephrase for Excel operations
      const analysis = await this.analyzeIntent(userInput);
      console.log('📋 Intent Analysis:', analysis);

      // Step 2: Plan the sequence of Excel API operations
      const executionPlan = await this.planExecution(analysis);
      console.log('⚡ Execution Plan:', executionPlan);

      // Step 3: Execute the planned operations
      const results = await this.executeOperations(executionPlan);
      console.log('✅ Execution Results:', results);

      // Step 4: Validate and synthesize response
      const response = await this.synthesizeResponse(analysis, executionPlan, results);
      
      return {
        success: true,
        analysis: analysis,
        executionPlan: executionPlan,
        results: results,
        response: response,
        timestamp: new Date().toISOString()
      };

    } catch (error) {
      console.error('❌ IntelligentExcelAgent error:', error);
      return {
        success: false,
        error: error.message,
        userInput: userInput,
        timestamp: new Date().toISOString()
      };
    }
  }

  /**
   * Analyze user intent and rephrase for Excel API operations
   */
  async analyzeIntent(userInput) {
    const lowerInput = userInput.toLowerCase();

    // Intent categories and their patterns
    const intentPatterns = {
      colorFormatting: {
        patterns: [/change.*color/i, /make.*blue|red|green|yellow/i, /(blue|red|green|yellow).*to.*(blue|red|green|yellow)/i],
        confidence: 0.9,
        description: "User wants to change cell colors"
      },
      headerModification: {
        patterns: [/header/i, /title/i, /column.*name/i],
        confidence: 0.8,
        description: "User wants to modify headers"
      },
      textFormatting: {
        patterns: [/bold/i, /italic/i, /font/i, /size/i],
        confidence: 0.8,
        description: "User wants to change text formatting"
      },
      dataSearch: {
        patterns: [/find/i, /search/i, /locate/i, /where/i],
        confidence: 0.7,
        description: "User wants to find specific data"
      },
      valueModification: {
        patterns: [/change.*value/i, /update.*number/i, /set.*to/i],
        confidence: 0.7,
        description: "User wants to modify cell values"
      }
    };

    let bestMatch = {
      intent: 'general',
      confidence: 0.3,
      description: 'General Excel operation'
    };

    // Find best matching intent
    for (const [intentType, pattern] of Object.entries(intentPatterns)) {
      for (const regex of pattern.patterns) {
        if (regex.test(userInput)) {
          if (pattern.confidence > bestMatch.confidence) {
            bestMatch = {
              intent: intentType,
              confidence: pattern.confidence,
              description: pattern.description
            };
          }
        }
      }
    }

    // Extract specific details from the input
    const details = this.extractDetails(userInput, bestMatch.intent);

    // Rephrase for AI reasoning
    const aiRephrase = this.generateAiRephrase(userInput, bestMatch, details);

    return {
      originalInput: userInput,
      intent: bestMatch.intent,
      confidence: bestMatch.confidence,
      description: bestMatch.description,
      extractedDetails: details,
      aiRephrase: aiRephrase,
      timestamp: new Date().toISOString()
    };
  }

  /**
   * Extract specific details based on intent type
   */
  extractDetails(input, intent) {
    const details = {};

    // Color extraction
    const colors = ['red', 'blue', 'green', 'yellow', 'orange', 'purple', 'black', 'white'];
    const colorMatches = colors.filter(color => 
      new RegExp(`\\b${color}\\b`, 'i').test(input)
    );
    
    if (colorMatches.length > 0) {
      details.colors = colorMatches;
      
      // Try to determine source and target colors
      const colorChangePattern = /(blue|red|green|yellow|orange|purple|black|white).*to.*(blue|red|green|yellow|orange|purple|black|white)/i;
      const match = input.match(colorChangePattern);
      if (match) {
        details.sourceColor = match[1].toLowerCase();
        details.targetColor = match[2].toLowerCase();
      } else if (colorMatches.length >= 2) {
        details.sourceColor = colorMatches[0];
        details.targetColor = colorMatches[1];
      }
    }

    // Target element extraction
    if (/header/i.test(input)) {
      details.targetElements = ['headers'];
    } else if (/cell/i.test(input)) {
      details.targetElements = ['cells'];
    } else if (/row/i.test(input)) {
      details.targetElements = ['rows'];
    } else if (/column/i.test(input)) {
      details.targetElements = ['columns'];
    }

    // Format type extraction
    if (/bold/i.test(input)) {
      details.formatTypes = details.formatTypes || [];
      details.formatTypes.push('bold');
    }
    if (/italic/i.test(input)) {
      details.formatTypes = details.formatTypes || [];
      details.formatTypes.push('italic');
    }

    // Value patterns
    const valuePatterns = input.match(/\b\d+(\.\d+)?\b/g);
    if (valuePatterns) {
      details.numbers = valuePatterns;
    }

    return details;
  }

  /**
   * Generate AI rephrase for reasoning
   */
  generateAiRephrase(originalInput, intent, details) {
    const templates = {
      colorFormatting: `The user wants to change cell formatting in Excel. Specifically: "${originalInput}". I need to find cells with ${details.sourceColor || 'current'} color and change them to ${details.targetColor || 'a different color'}. Target elements: ${details.targetElements ? details.targetElements.join(', ') : 'unspecified cells'}.`,
      
      headerModification: `The user wants to modify headers in Excel. Request: "${originalInput}". I need to identify header cells and apply the requested changes.`,
      
      textFormatting: `The user wants to change text formatting in Excel. Request: "${originalInput}". Format types to apply: ${details.formatTypes ? details.formatTypes.join(', ') : 'general formatting'}.`,
      
      dataSearch: `The user wants to find specific data in Excel. Request: "${originalInput}". I need to search for and locate the specified content.`,
      
      valueModification: `The user wants to modify cell values in Excel. Request: "${originalInput}". I need to find the target cells and update their values.`,
      
      general: `The user has made a general Excel request: "${originalInput}". I need to analyze the Excel workbook and determine the appropriate operations to fulfill this request.`
    };

    return templates[intent.intent] || templates.general;
  }

  /**
   * Plan the sequence of Excel API operations needed
   */
  async planExecution(analysis) {
    const { intent, extractedDetails } = analysis;
    
    const executionPlan = {
      intent: intent,
      steps: [],
      requiredTools: [],
      estimatedTime: 0
    };

    // Plan based on intent type
    switch (intent) {
      case 'colorFormatting':
        await this.planColorFormattingWorkflow(executionPlan, extractedDetails);
        break;
        
      case 'headerModification':
        await this.planHeaderModificationWorkflow(executionPlan, extractedDetails);
        break;
        
      case 'textFormatting':
        await this.planTextFormattingWorkflow(executionPlan, extractedDetails);
        break;
        
      case 'dataSearch':
        await this.planDataSearchWorkflow(executionPlan, extractedDetails);
        break;
        
      default:
        await this.planGeneralWorkflow(executionPlan, extractedDetails);
        break;
    }

    return executionPlan;
  }

  /**
   * Plan workflow for color formatting operations
   */
  async planColorFormattingWorkflow(plan, details) {
    // Step 1: Analyze workbook structure
    plan.steps.push({
      stepNumber: 1,
      tool: 'analyze_workbook_structure',
      purpose: 'Get overview of workbook sheets and structure',
      args: { includeFormatting: true, maxRows: 20 }
    });

    // Step 2: Find cells by source color (if specified)
    if (details.sourceColor) {
      plan.steps.push({
        stepNumber: 2,
        tool: 'find_cells_by_color',
        purpose: `Find all cells with ${details.sourceColor} color`,
        args: { 
          color: details.sourceColor, 
          colorType: 'background'
        }
      });
    }

    // Step 3: Analyze headers if targeting headers specifically
    if (details.targetElements && details.targetElements.includes('headers')) {
      plan.steps.push({
        stepNumber: plan.steps.length + 1,
        tool: 'analyze_headers',
        purpose: 'Identify which cells are headers',
        args: {}
      });
    }

    // Step 4: Apply new formatting
    plan.steps.push({
      stepNumber: plan.steps.length + 1,
      tool: 'format_cells',
      purpose: `Apply ${details.targetColor} color formatting`,
      args: { 
        backgroundColor: details.targetColor
      },
      dependsOn: [2, 3] // Depends on finding cells and identifying headers
    });

    // Step 5: Verify changes
    plan.steps.push({
      stepNumber: plan.steps.length + 1,
      tool: 'read_formatting',
      purpose: 'Verify the formatting changes were applied correctly',
      args: {},
      dependsOn: [4]
    });

    plan.requiredTools = ['analyze_workbook_structure', 'find_cells_by_color', 'analyze_headers', 'format_cells', 'read_formatting'];
    plan.estimatedTime = 5000; // 5 seconds
  }

  /**
   * Plan workflow for header modification
   */
  async planHeaderModificationWorkflow(plan, details) {
    plan.steps.push({
      stepNumber: 1,
      tool: 'analyze_workbook_structure',
      purpose: 'Understand workbook layout',
      args: { includeFormatting: true }
    });

    plan.steps.push({
      stepNumber: 2,
      tool: 'analyze_headers',
      purpose: 'Identify all headers in the workbook',
      args: {}
    });

    // Apply formatting based on details
    const formatArgs = {};
    if (details.targetColor) {
      formatArgs.backgroundColor = details.targetColor;
    }
    if (details.formatTypes) {
      if (details.formatTypes.includes('bold')) formatArgs.bold = true;
      if (details.formatTypes.includes('italic')) formatArgs.italic = true;
    }

    plan.steps.push({
      stepNumber: 3,
      tool: 'format_cells',
      purpose: 'Apply formatting to identified headers',
      args: formatArgs,
      dependsOn: [2]
    });

    plan.requiredTools = ['analyze_workbook_structure', 'analyze_headers', 'format_cells'];
    plan.estimatedTime = 4000;
  }

  /**
   * Plan general workflow for unknown intents
   */
  async planGeneralWorkflow(plan, details) {
    // Always start with workbook analysis
    plan.steps.push({
      stepNumber: 1,
      tool: 'analyze_workbook_structure',
      purpose: 'Understand the Excel workbook structure and content',
      args: { includeFormatting: true, maxRows: 10 }
    });

    plan.requiredTools = ['analyze_workbook_structure'];
    plan.estimatedTime = 2000;
  }

  /**
   * Execute the planned operations
   */
  async executeOperations(executionPlan) {
    const results = {
      steps: [],
      success: true,
      errors: []
    };

    console.log(`⚡ Executing ${executionPlan.steps.length} planned operations`);

    for (const step of executionPlan.steps) {
      console.log(`🔧 Step ${step.stepNumber}: ${step.tool} - ${step.purpose}`);
      
      try {
        const tool = this.tools.get(step.tool);
        if (!tool) {
          throw new Error(`Tool '${step.tool}' not found`);
        }

        // Execute dependent steps' results into current step args if needed
        let resolvedArgs = { ...step.args };
        if (step.dependsOn) {
          resolvedArgs = await this.resolveDependencies(step, results.steps, resolvedArgs);
        }

        const startTime = Date.now();
        const result = await tool.call(resolvedArgs);
        const executionTime = Date.now() - startTime;
        
        const parsedResult = JSON.parse(result);
        
        results.steps.push({
          stepNumber: step.stepNumber,
          tool: step.tool,
          purpose: step.purpose,
          args: resolvedArgs,
          result: parsedResult,
          success: parsedResult.success,
          executionTime: executionTime
        });

        if (!parsedResult.success) {
          results.errors.push(`Step ${step.stepNumber} failed: ${parsedResult.error}`);
          results.success = false;
        }

      } catch (error) {
        console.error(`❌ Step ${step.stepNumber} error:`, error);
        results.steps.push({
          stepNumber: step.stepNumber,
          tool: step.tool,
          purpose: step.purpose,
          args: step.args,
          result: null,
          success: false,
          error: error.message,
          executionTime: 0
        });
        results.errors.push(`Step ${step.stepNumber} failed: ${error.message}`);
        results.success = false;
      }
    }

    return results;
  }

  /**
   * Resolve dependencies between steps
   */
  async resolveDependencies(currentStep, completedSteps, args) {
    // Find completed steps that this step depends on
    const dependencies = currentStep.dependsOn || [];
    
    for (const depStepNumber of dependencies) {
      const depStep = completedSteps.find(s => s.stepNumber === depStepNumber);
      if (depStep && depStep.success) {
        // Inject results from dependency into current step args
        if (depStep.tool === 'find_cells_by_color') {
          args.cells = depStep.result.matches.map(m => m.address);
        } else if (depStep.tool === 'analyze_headers') {
          if (currentStep.tool === 'format_cells') {
            args.cells = depStep.result.headers.map(h => h.address);
          }
        }
      }
    }

    return args;
  }

  /**
   * Synthesize final response based on execution results
   */
  async synthesizeResponse(analysis, executionPlan, results) {
    let response = "";

    // Generate title based on intent
    const intentTitles = {
      colorFormatting: '🎨 Color Formatting Complete',
      headerModification: '📋 Header Modification Complete',
      textFormatting: '✍️ Text Formatting Applied',
      dataSearch: '🔍 Data Search Results',
      general: '📊 Excel Operation Complete'
    };

    response += `## ${intentTitles[analysis.intent] || intentTitles.general}\n\n`;

    // Add AI reasoning explanation
    response += `### 🧠 AI Analysis\n${analysis.aiRephrase}\n\n`;

    // Add execution summary
    response += `### ⚡ Execution Summary\n`;
    response += `- **Steps Executed**: ${results.steps.length}\n`;
    response += `- **Success Rate**: ${results.steps.filter(s => s.success).length}/${results.steps.length}\n`;
    response += `- **Total Time**: ${results.steps.reduce((sum, s) => sum + (s.executionTime || 0), 0)}ms\n\n`;

    // Add detailed results
    response += `### 📋 Detailed Results\n`;
    
    for (const step of results.steps) {
      const statusIcon = step.success ? '✅' : '❌';
      response += `${statusIcon} **Step ${step.stepNumber}**: ${step.purpose}\n`;
      
      if (step.success && step.result) {
        // Summarize key results
        if (step.tool === 'find_cells_by_color') {
          response += `   - Found ${step.result.count} cells with the specified color\n`;
          if (step.result.matches && step.result.matches.length > 0) {
            const addresses = step.result.matches.slice(0, 5).map(m => m.address).join(', ');
            response += `   - Locations: ${addresses}${step.result.matches.length > 5 ? '...' : ''}\n`;
          }
        } else if (step.tool === 'analyze_headers') {
          response += `   - Identified ${step.result.count} headers\n`;
          if (step.result.headers && step.result.headers.length > 0) {
            const headers = step.result.headers.slice(0, 3).map(h => `${h.address}("${h.value}")`).join(', ');
            response += `   - Headers: ${headers}${step.result.headers.length > 3 ? '...' : ''}\n`;
          }
        } else if (step.tool === 'format_cells') {
          response += `   - Formatted ${step.result.count} cells successfully\n`;
        }
      } else if (!step.success) {
        response += `   - Error: ${step.error}\n`;
      }
      
      response += `\n`;
    }

    // Add recommendations if there were issues
    if (!results.success) {
      response += `### ⚠️ Issues Encountered\n`;
      results.errors.forEach(error => {
        response += `- ${error}\n`;
      });
      response += `\n💡 **Recommendation**: Please verify your Excel workbook contains the expected data and try again.\n`;
    }

    return response;
  }
}

// Initialize globally
if (typeof window !== 'undefined') {
  window.IntelligentExcelAgent = IntelligentExcelAgent;
  console.log('✅ IntelligentExcelAgent initialized');
}