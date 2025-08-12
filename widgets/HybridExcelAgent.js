/**
 * Hybrid Excel Agent
 * Combines simple unified agent with deep agent concepts for complex workflows
 */

class HybridExcelAgent {
  constructor(apiKey) {
    this.simpleAgent = new UnifiedAiAgent(apiKey);
    this.planningEnabled = true;
    this.workflowMemory = new Map(); // Simple persistence
  }

  /**
   * Main entry point - decides between simple and complex processing
   */
  async processRequest(userInput) {
    console.log('🤔 Analyzing request complexity...');
    
    const complexity = this.assessComplexity(userInput);
    console.log(`📊 Complexity assessment: ${complexity.level} (score: ${complexity.score})`);

    if (complexity.level === 'simple') {
      // Use fast, direct approach
      return await this.handleSimpleRequest(userInput);
    } else {
      // Use planning + multi-step approach
      return await this.handleComplexRequest(userInput, complexity);
    }
  }

  /**
   * Assess if request needs complex planning
   */
  assessComplexity(input) {
    const lowerInput = input.toLowerCase();
    let score = 0;
    const indicators = [];

    // Multi-step indicators
    const multiStepWords = ['and', 'then', 'also', 'additionally', 'furthermore'];
    multiStepWords.forEach(word => {
      if (lowerInput.includes(word)) {
        score += 15;
        indicators.push(`Multi-step word: ${word}`);
      }
    });

    // Complex operation indicators
    const complexOps = ['analyze', 'comprehensive', 'report', 'summary', 'research', 'compare', 'evaluate'];
    complexOps.forEach(op => {
      if (lowerInput.includes(op)) {
        score += 20;
        indicators.push(`Complex operation: ${op}`);
      }
    });

    // Multiple Excel operations
    const excelOps = ['format', 'calculate', 'find', 'change', 'create', 'add', 'update'];
    const opsFound = excelOps.filter(op => lowerInput.includes(op));
    if (opsFound.length > 2) {
      score += 25;
      indicators.push(`Multiple Excel operations: ${opsFound.join(', ')}`);
    }

    // Length indicator (longer requests often more complex)
    if (input.length > 100) {
      score += 10;
      indicators.push('Long request');
    }

    // Time indicators
    if (lowerInput.includes('step') || lowerInput.includes('first') || lowerInput.includes('second')) {
      score += 15;
      indicators.push('Step-by-step indicators');
    }

    const level = score >= 40 ? 'complex' : 'simple';
    
    return {
      level,
      score,
      indicators,
      threshold: 40
    };
  }

  /**
   * Handle simple requests with direct agent
   */
  async handleSimpleRequest(userInput) {
    console.log('⚡ Using simple agent for fast execution');
    return await this.simpleAgent.processRequest(userInput);
  }

  /**
   * Handle complex requests with planning approach
   */
  async handleComplexRequest(userInput, complexity) {
    console.log('🧠 Using complex workflow with planning');

    try {
      // Step 1: Create a plan
      const plan = await this.createPlan(userInput);
      console.log('📋 Plan created:', plan);

      // Step 2: Execute plan steps
      const executionResults = await this.executePlan(plan, userInput);

      // Step 3: Synthesize final response
      const finalResponse = await this.synthesizePlanResults(plan, executionResults, userInput);

      return {
        success: true,
        response: finalResponse,
        complexity: complexity,
        plan: plan,
        executionResults: executionResults,
        timestamp: new Date().toISOString()
      };

    } catch (error) {
      console.error('❌ Complex workflow failed:', error);
      
      // Fallback to simple agent
      console.log('🔄 Falling back to simple agent');
      return await this.handleSimpleRequest(userInput);
    }
  }

  /**
   * Create execution plan for complex request
   */
  async createPlan(userInput) {
    const planningPrompt = `You are an Excel workflow planner. Break down this request into 3-5 specific, actionable steps:

User Request: "${userInput}"

Create a JSON plan with this structure:
{
  "steps": [
    {
      "stepNumber": 1,
      "action": "Specific action to take",
      "excelOperations": ["list", "of", "excel", "functions", "needed"],
      "expectedOutput": "What this step should accomplish"
    }
  ],
  "overallGoal": "Summary of the complete request"
}

Focus on Excel operations that can be executed programmatically.`;

    const planResponse = await this.simpleAgent.callOpenAi([
      { role: "system", content: planningPrompt },
      { role: "user", content: userInput }
    ]);

    try {
      return JSON.parse(planResponse);
    } catch {
      // If JSON parsing fails, create simple plan
      return {
        steps: [
          {
            stepNumber: 1,
            action: "Analyze request and execute Excel operations",
            excelOperations: ["analyze", "execute"],
            expectedOutput: "Complete the user's request"
          }
        ],
        overallGoal: userInput
      };
    }
  }

  /**
   * Execute each step of the plan
   */
  async executePlan(plan, originalRequest) {
    const results = [];

    for (const step of plan.steps) {
      console.log(`🔧 Executing step ${step.stepNumber}: ${step.action}`);

      try {
        // Create focused request for this step
        const stepRequest = `${step.action}. Context: ${originalRequest}`;
        
        const stepResult = await this.simpleAgent.processRequest(stepRequest);
        
        results.push({
          stepNumber: step.stepNumber,
          action: step.action,
          result: stepResult,
          success: stepResult.success
        });

        // Store intermediate result for next steps
        this.workflowMemory.set(`step_${step.stepNumber}`, stepResult);

      } catch (error) {
        console.error(`❌ Step ${step.stepNumber} failed:`, error);
        results.push({
          stepNumber: step.stepNumber,
          action: step.action,
          result: { success: false, error: error.message },
          success: false
        });
      }
    }

    return results;
  }

  /**
   * Synthesize final response from plan execution
   */
  async synthesizePlanResults(plan, executionResults, originalRequest) {
    const successfulSteps = executionResults.filter(r => r.success);
    const failedSteps = executionResults.filter(r => !r.success);

    let response = `## 🎯 Complex Workflow Completed\n\n`;
    response += `**Original Request**: ${originalRequest}\n\n`;
    response += `### 📋 Execution Summary\n`;
    response += `- **Total Steps**: ${executionResults.length}\n`;
    response += `- **Successful**: ${successfulSteps.length}\n`;
    response += `- **Failed**: ${failedSteps.length}\n\n`;

    response += `### ✅ Completed Steps\n`;
    successfulSteps.forEach(step => {
      response += `**${step.stepNumber}.** ${step.action}\n`;
      if (step.result.response) {
        response += `   - ${step.result.response.substring(0, 150)}...\n`;
      }
      response += `\n`;
    });

    if (failedSteps.length > 0) {
      response += `### ❌ Issues Encountered\n`;
      failedSteps.forEach(step => {
        response += `**${step.stepNumber}.** ${step.action}: ${step.result.error}\n`;
      });
      response += `\n`;
    }

    response += `### 🎉 Overall Result\n`;
    if (successfulSteps.length >= executionResults.length * 0.7) {
      response += `Workflow completed successfully! ${successfulSteps.length} out of ${executionResults.length} steps executed successfully.`;
    } else {
      response += `Workflow partially completed. Please review the issues above and try again if needed.`;
    }

    return response;
  }

  /**
   * Helper method for simple OpenAI calls
   */
  async callOpenAi(messages) {
    // Delegate to simple agent's OpenAI calling method
    const response = await this.simpleAgent.callOpenAiWithFunctions(messages);
    return response.choices[0].message.content;
  }

  /**
   * Clear workflow memory
   */
  clearWorkflowMemory() {
    this.workflowMemory.clear();
    console.log('🧹 Workflow memory cleared');
  }

  /**
   * Get current status
   */
  getStatus() {
    return {
      ...this.simpleAgent.getStatus(),
      planningEnabled: this.planningEnabled,
      workflowMemorySize: this.workflowMemory.size,
      type: 'hybrid'
    };
  }
}

// Initialize globally
if (typeof window !== 'undefined') {
  window.HybridExcelAgent = HybridExcelAgent;
  console.log('✅ Hybrid Excel Agent initialized');
}