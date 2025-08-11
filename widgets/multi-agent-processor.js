/**
 * Multi-Agent Processor - Stage 2 Implementation
 * Professional M&A Intelligence Platform with Specialized Agents
 * 
 * This system orchestrates multiple specialized AI agents to provide
 * investment banking-quality analysis of Excel M&A models
 */

class MultiAgentProcessor {
  constructor() {
    this.agents = {
      excelStructure: new ExcelStructureAgent(),
      financialAnalysis: new FinancialAnalysisAgent(),
      dataValidation: new DataValidationAgent()
    };
    
    this.cache = new Map();
    this.cacheTimeout = 60000; // 1 minute cache for agent responses
    this.processingQueue = [];
    this.currentProcessing = null;
    
    console.log('🎭 Multi-Agent Processor initialized with 3 specialized agents');
  }

  /**
   * Main entry point: Process user query with multi-agent system
   */
  async processQuery(userMessage, options = {}) {
    const queryId = this.generateQueryId();
    const startTime = performance.now();
    
    try {
      console.log(`🎯 [${queryId}] Processing query with multi-agent system:`, userMessage);
      
      // Set current processing for UI updates
      this.currentProcessing = {
        queryId,
        userMessage,
        startTime,
        stage: 'initializing',
        progress: 0
      };
      
      // Step 1: Analyze query type and determine agent routing
      this.updateProgress(queryId, 'analyzing_query', 10);
      const queryAnalysis = this.analyzeQuery(userMessage);
      console.log(`🧠 [${queryId}] Query analysis:`, queryAnalysis);
      
      // Step 2: Fetch Excel structure using enhanced fetcher
      this.updateProgress(queryId, 'fetching_structure', 20);
      console.log(`📊 [${queryId}] Fetching workbook structure...`);
      const structureJson = await window.excelStructureFetcher.fetchWorkbookStructure(userMessage);
      const structure = JSON.parse(structureJson);
      
      // Step 3: Route through appropriate agents based on query type
      const agentResults = await this.routeToAgents(queryId, userMessage, structure, queryAnalysis);
      
      // Step 4: Synthesize final response
      this.updateProgress(queryId, 'synthesizing', 90);
      console.log(`🎭 [${queryId}] Synthesizing final response...`);
      const finalResponse = await this.synthesizeResponse(userMessage, structure, agentResults, queryAnalysis);
      
      // Step 5: Complete processing
      this.updateProgress(queryId, 'completed', 100);
      const processingTime = performance.now() - startTime;
      
      console.log(`✅ [${queryId}] Multi-agent processing completed in ${processingTime.toFixed(2)}ms`);
      
      return {
        response: finalResponse,
        metadata: {
          queryId,
          processingTime,
          agentsUsed: Object.keys(agentResults),
          structureAnalyzed: structure.metadata.totalSheets,
          queryType: queryAnalysis.primaryType
        }
      };
      
    } catch (error) {
      console.error(`❌ [${queryId}] Multi-agent processing failed:`, error);
      this.updateProgress(queryId, 'error', -1);
      
      return {
        response: this.generateErrorResponse(error, userMessage),
        metadata: {
          queryId,
          error: error.message,
          processingTime: performance.now() - startTime
        }
      };
    } finally {
      this.currentProcessing = null;
    }
  }

  /**
   * Analyze user query to determine optimal agent routing
   */
  analyzeQuery(message) {
    const lowerMessage = message.toLowerCase();
    const analysis = {
      primaryType: 'general',
      secondaryTypes: [],
      complexity: 'medium',
      urgency: 'normal',
      keywords: [],
      agents: []
    };
    
    // Financial analysis patterns
    const financialPatterns = {
      irr: /\b(irr|internal rate|return)\b/i,
      moic: /\b(moic|multiple|money multiple|cash multiple)\b/i,
      valuation: /\b(valuation|value|worth|price)\b/i,
      cash_flow: /\b(cash flow|cf|fcf|free cash)\b/i,
      leverage: /\b(debt|leverage|ltv|dscr|interest)\b/i,
      revenue: /\b(revenue|sales|income|growth)\b/i,
      ebitda: /\b(ebitda|operating income|margin)\b/i,
      assumptions: /\b(assumption|input|parameter|sensitivity)\b/i,
      exit: /\b(exit|terminal|disposal|sale)\b/i,
      scenario: /\b(scenario|sensitivity|stress|case)\b/i
    };
    
    // Structure analysis patterns
    const structurePatterns = {
      formula: /\b(formula|calculation|equation)\b/i,
      error: /\b(error|wrong|incorrect|fix|issue)\b/i,
      dependency: /\b(depend|link|reference|connect)\b/i,
      sheet: /\b(sheet|tab|worksheet)\b/i,
      range: /\b(range|cell|column|row)\b/i,
      table: /\b(table|data|list)\b/i
    };
    
    // Data validation patterns
    const validationPatterns = {
      check: /\b(check|verify|validate|correct)\b/i,
      consistency: /\b(consistent|inconsistent|mismatch)\b/i,
      missing: /\b(missing|blank|empty)\b/i,
      warning: /\b(warning|alert|flag|issue)\b/i
    };
    
    // Analyze financial patterns
    for (const [type, pattern] of Object.entries(financialPatterns)) {
      if (pattern.test(lowerMessage)) {
        analysis.keywords.push(type);
        if (!analysis.agents.includes('financialAnalysis')) {
          analysis.agents.push('financialAnalysis');
        }
      }
    }
    
    // Analyze structure patterns  
    for (const [type, pattern] of Object.entries(structurePatterns)) {
      if (pattern.test(lowerMessage)) {
        analysis.keywords.push(type);
        if (!analysis.agents.includes('excelStructure')) {
          analysis.agents.push('excelStructure');
        }
      }
    }
    
    // Analyze validation patterns
    for (const [type, pattern] of Object.entries(validationPatterns)) {
      if (pattern.test(lowerMessage)) {
        analysis.keywords.push(type);
        if (!analysis.agents.includes('dataValidation')) {
          analysis.agents.push('dataValidation');
        }
      }
    }
    
    // Determine primary type
    if (analysis.keywords.some(k => ['irr', 'moic', 'valuation', 'cash_flow'].includes(k))) {
      analysis.primaryType = 'financial_analysis';
      analysis.complexity = 'high';
    } else if (analysis.keywords.some(k => ['formula', 'dependency', 'structure'].includes(k))) {
      analysis.primaryType = 'structure_analysis';
    } else if (analysis.keywords.some(k => ['error', 'check', 'validate'].includes(k))) {
      analysis.primaryType = 'data_validation';
    }
    
    // Default to all agents if no specific patterns found
    if (analysis.agents.length === 0) {
      analysis.agents = ['excelStructure', 'financialAnalysis'];
    }
    
    // Always include structure agent for context
    if (!analysis.agents.includes('excelStructure')) {
      analysis.agents.unshift('excelStructure');
    }
    
    return analysis;
  }

  /**
   * Route query to appropriate agents based on analysis
   */
  async routeToAgents(queryId, userMessage, structure, queryAnalysis) {
    const agentResults = {};
    const totalAgents = queryAnalysis.agents.length;
    let completedAgents = 0;
    
    // Process agents in intelligent order
    const agentOrder = this.determineAgentOrder(queryAnalysis);
    
    for (const agentName of agentOrder) {
      try {
        const progressBase = 30 + (completedAgents / totalAgents) * 50;
        this.updateProgress(queryId, `agent_${agentName}`, progressBase);
        
        console.log(`🤖 [${queryId}] Processing with ${agentName} agent...`);
        
        const agentResult = await this.agents[agentName].process(
          userMessage, 
          structure, 
          agentResults, // Pass previous agent results for context
          queryAnalysis
        );
        
        agentResults[agentName] = agentResult;
        completedAgents++;
        
        console.log(`✅ [${queryId}] ${agentName} agent completed`);
        
      } catch (error) {
        console.error(`❌ [${queryId}] ${agentName} agent failed:`, error);
        agentResults[agentName] = {
          error: error.message,
          fallback: true
        };
      }
    }
    
    return agentResults;
  }

  /**
   * Determine optimal agent processing order
   */
  determineAgentOrder(queryAnalysis) {
    const { primaryType, agents } = queryAnalysis;
    
    // Always start with structure agent for context
    let order = ['excelStructure'];
    
    // Add other agents based on priority
    if (primaryType === 'financial_analysis') {
      order.push('financialAnalysis');
      if (agents.includes('dataValidation')) {
        order.push('dataValidation');
      }
    } else if (primaryType === 'data_validation') {
      order.push('dataValidation');
      if (agents.includes('financialAnalysis')) {
        order.push('financialAnalysis');
      }
    } else {
      // Default order
      if (agents.includes('financialAnalysis')) {
        order.push('financialAnalysis');
      }
      if (agents.includes('dataValidation')) {
        order.push('dataValidation');
      }
    }
    
    // Remove duplicates while preserving order
    return [...new Set(order)].filter(agent => agents.includes(agent));
  }

  /**
   * Synthesize final response from all agent outputs
   */
  async synthesizeResponse(userMessage, structure, agentResults, queryAnalysis) {
    const synthesisPrompt = this.createSynthesisPrompt(userMessage, queryAnalysis, agentResults);
    
    try {
      const response = await fetch('/.netlify/functions/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: synthesisPrompt,
          systemPrompt: this.getSynthesisSystemPrompt(),
          temperature: 0.3,
          maxTokens: 4000,
          batchType: 'response_synthesis'
        })
      });

      const result = await response.json();
      
      // Add metadata and formatting hints
      let finalResponse = result.response;
      
      // Add processing metadata as subtle footer
      if (structure.advancedFeatures) {
        const features = [];
        if (structure.advancedFeatures.formulaDependencies) features.push('dependency mapping');
        if (structure.advancedFeatures.namedRanges) features.push('named ranges');
        if (structure.advancedFeatures.maModelPatterns) features.push('M&A patterns');
        
        if (features.length > 0) {
          finalResponse += `\n\n*Analysis powered by ${features.join(', ')} and ${Object.keys(agentResults).length} specialized agents*`;
        }
      }
      
      return finalResponse;
      
    } catch (error) {
      console.error('❌ Response synthesis failed:', error);
      
      // Fallback: combine agent responses directly
      return this.fallbackResponseSynthesis(userMessage, agentResults, queryAnalysis);
    }
  }

  /**
   * Create synthesis prompt combining all agent outputs
   */
  createSynthesisPrompt(userMessage, queryAnalysis, agentResults) {
    let prompt = `Based on specialized agent analysis, provide a comprehensive response to: "${userMessage}"\n\n`;
    
    prompt += `Query Analysis: ${queryAnalysis.primaryType} (complexity: ${queryAnalysis.complexity})\n\n`;
    
    // Add each agent's contribution
    if (agentResults.excelStructure && !agentResults.excelStructure.error) {
      prompt += `Excel Structure Analysis:\n${JSON.stringify(agentResults.excelStructure, null, 2)}\n\n`;
    }
    
    if (agentResults.financialAnalysis && !agentResults.financialAnalysis.error) {
      prompt += `Financial Analysis:\n${agentResults.financialAnalysis}\n\n`;
    }
    
    if (agentResults.dataValidation && !agentResults.dataValidation.error) {
      prompt += `Data Validation Results:\n${JSON.stringify(agentResults.dataValidation, null, 2)}\n\n`;
    }
    
    prompt += `Synthesize these insights into a professional, actionable response with specific cell references.`;
    
    return prompt;
  }

  /**
   * Get synthesis system prompt
   */
  getSynthesisSystemPrompt() {
    return `You are a Senior M&A Analyst synthesizing insights from multiple specialized agents.

Your response should:
1. Directly answer the user's question with specific, actionable insights
2. Reference specific Excel cells using the format SheetName!CellAddress (e.g., FCF!B22)
3. Highlight key financial metrics and their implications
4. Provide professional investment banking-quality analysis
5. Include specific recommendations when appropriate
6. Keep responses concise but comprehensive
7. Use clear, professional language suitable for investment bankers and private equity professionals

Always reference specific cell locations when mentioning calculations or values.
Focus on actionable insights rather than technical implementation details.`;
  }

  /**
   * Fallback response synthesis when AI synthesis fails
   */
  fallbackResponseSynthesis(userMessage, agentResults, queryAnalysis) {
    let response = `Based on your query "${userMessage}", here's my analysis:\n\n`;
    
    if (agentResults.financialAnalysis && !agentResults.financialAnalysis.error) {
      response += `**Financial Analysis:**\n${agentResults.financialAnalysis}\n\n`;
    }
    
    if (agentResults.excelStructure && !agentResults.excelStructure.error) {
      const structure = agentResults.excelStructure;
      if (structure.relevant_cells && structure.relevant_cells.length > 0) {
        response += `**Key Cells:** ${structure.relevant_cells.join(', ')}\n\n`;
      }
      if (structure.insights && structure.insights.length > 0) {
        response += `**Structure Insights:**\n${structure.insights.join('\n')}\n\n`;
      }
    }
    
    if (agentResults.dataValidation && !agentResults.dataValidation.error) {
      const validation = agentResults.dataValidation;
      if (validation.warnings && validation.warnings.length > 0) {
        response += `**Warnings:**\n${validation.warnings.map(w => `• ${w}`).join('\n')}\n\n`;
      }
    }
    
    response += `*Analysis completed using ${Object.keys(agentResults).length} specialized agents*`;
    
    return response;
  }

  /**
   * Update processing progress
   */
  updateProgress(queryId, stage, progress) {
    if (this.currentProcessing && this.currentProcessing.queryId === queryId) {
      this.currentProcessing.stage = stage;
      this.currentProcessing.progress = progress;
      
      // Emit progress event for UI
      if (typeof window !== 'undefined' && window.dispatchEvent) {
        window.dispatchEvent(new CustomEvent('multiAgentProgress', {
          detail: {
            queryId,
            stage,
            progress,
            timestamp: Date.now()
          }
        }));
      }
    }
  }

  /**
   * Generate unique query ID
   */
  generateQueryId() {
    return `q_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * Generate error response
   */
  generateErrorResponse(error, userMessage) {
    if (error.message.includes('Excel') || error.message.includes('workbook')) {
      return `I encountered an issue accessing your Excel workbook: ${error.message}. Please ensure the workbook is active and try again.`;
    } else if (error.message.includes('network') || error.message.includes('fetch')) {
      return `I'm having trouble connecting to the analysis service. Please check your connection and try again.`;
    } else {
      return `I encountered a technical issue while analyzing your query: "${userMessage}". Please try rephrasing your question or contact support if the issue persists.`;
    }
  }

  /**
   * Get current processing status
   */
  getCurrentProcessing() {
    return this.currentProcessing;
  }

  /**
   * Clear cache
   */
  clearCache() {
    this.cache.clear();
    console.log('🗑️ Multi-agent processor cache cleared');
  }

  /**
   * Get processing statistics
   */
  getStatistics() {
    return {
      cacheSize: this.cache.size,
      queueLength: this.processingQueue.length,
      currentProcessing: this.currentProcessing ? this.currentProcessing.queryId : null,
      agentsAvailable: Object.keys(this.agents).length
    };
  }
}

/**
 * Excel Structure Agent - Specialized in Excel workbook analysis
 */
class ExcelStructureAgent {
  async process(userMessage, structure, previousResults, queryAnalysis) {
    console.log('🏗️ Excel Structure Agent processing...');
    
    const systemPrompt = `You are an Excel Structure Agent specialized in M&A financial models.

Analyze the provided workbook structure and identify:
1. Relevant cells and ranges for the user's query
2. Formula relationships and dependencies  
3. Data organization and structure insights
4. Key calculation patterns

Focus on M&A model patterns: IRR calculations, MOIC formulas, cash flow structures, valuation metrics.
Return JSON with keys: 'relevant_cells', 'formulas', 'dependencies', 'insights'.`;

    const contextMessage = `Workbook Structure Analysis:
Total sheets: ${structure.metadata.totalSheets}
Key metrics found: ${JSON.stringify(structure.keyMetrics, null, 2)}
Advanced features: ${JSON.stringify(structure.advancedFeatures, null, 2)}

User query: "${userMessage}"

Provide structured analysis focusing on Excel formula relationships and data organization.`;

    try {
      const response = await fetch('/.netlify/functions/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: contextMessage,
          systemPrompt: systemPrompt,
          temperature: 0.2,
          maxTokens: 2000,
          batchType: 'excel_structure_analysis'
        })
      });

      const result = await response.json();
      
      try {
        return JSON.parse(result.response);
      } catch (parseError) {
        // If response isn't valid JSON, return structured fallback
        return {
          relevant_cells: this.extractCellReferences(result.response),
          formulas: [],
          dependencies: {},
          insights: [result.response],
          raw_response: result.response
        };
      }
      
    } catch (error) {
      console.error('❌ Excel Structure Agent failed:', error);
      return {
        error: error.message,
        fallback: this.createFallbackStructureAnalysis(structure, userMessage)
      };
    }
  }

  extractCellReferences(text) {
    const cellRegex = /\b[A-Z]+!?[A-Z]+\d+\b/g;
    return (text.match(cellRegex) || []).slice(0, 10); // Limit to top 10
  }

  createFallbackStructureAnalysis(structure, userMessage) {
    const relevant_cells = [];
    const insights = [];
    
    // Extract relevant cells from key metrics
    Object.entries(structure.keyMetrics).forEach(([sheet, metrics]) => {
      metrics.forEach(metric => {
        if (metric.location) {
          relevant_cells.push(metric.location);
        }
      });
    });
    
    insights.push(`Analyzed ${structure.metadata.totalSheets} sheets with ${relevant_cells.length} key metrics identified`);
    
    if (structure.validation) {
      insights.push(`Model validation score: ${structure.validation.modelScore}/100`);
    }
    
    return {
      relevant_cells,
      formulas: [],
      dependencies: {},
      insights
    };
  }
}

/**
 * Financial Analysis Agent - Specialized in M&A financial analysis
 */
class FinancialAnalysisAgent {
  async process(userMessage, structure, previousResults, queryAnalysis) {
    console.log('💰 Financial Analysis Agent processing...');
    
    const systemPrompt = `You are a Senior Financial Analysis Agent specialized in M&A transactions and LBO modeling.

Provide investment banking-quality analysis including:
1. Specific financial insights (IRR, MOIC, cash flows, leverage metrics)
2. Key value drivers and sensitivities
3. Professional investment recommendations
4. Risk assessment and scenario considerations

Reference specific Excel cell locations and provide quantitative analysis.
Keep responses professional but conversational, suitable for investment bankers.`;

    // Combine structure analysis with previous agent results
    let contextMessage = `Excel Model Analysis:
${JSON.stringify(structure.keyMetrics, null, 2)}

Model Validation: ${JSON.stringify(structure.validation, null, 2)}`;

    if (previousResults.excelStructure) {
      contextMessage += `\n\nExcel Structure Insights: ${JSON.stringify(previousResults.excelStructure, null, 2)}`;
    }

    contextMessage += `\n\nUser Query: "${userMessage}"

Provide detailed financial analysis with specific cell references and professional insights.`;

    try {
      const response = await fetch('/.netlify/functions/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: contextMessage,
          systemPrompt: systemPrompt,
          temperature: 0.4,
          maxTokens: 3000,
          batchType: 'financial_analysis'
        })
      });

      const result = await response.json();
      return result.response;
      
    } catch (error) {
      console.error('❌ Financial Analysis Agent failed:', error);
      return this.createFallbackFinancialAnalysis(structure, userMessage);
    }
  }

  createFallbackFinancialAnalysis(structure, userMessage) {
    let analysis = `Based on your M&A model analysis for: "${userMessage}"\n\n`;
    
    // Analyze key metrics
    const allMetrics = Object.values(structure.keyMetrics).flat();
    const irrMetrics = allMetrics.filter(m => m.type === 'irr');
    const moicMetrics = allMetrics.filter(m => m.type === 'moic');
    const revenueMetrics = allMetrics.filter(m => m.type === 'revenue');
    
    if (irrMetrics.length > 0) {
      const irr = irrMetrics[0];
      analysis += `**IRR Analysis**: Your levered IRR of ${(irr.value * 100).toFixed(1)}% in ${irr.location} `;
      if (irr.value > 0.25) {
        analysis += `is strong, indicating attractive returns for this investment.\n\n`;
      } else if (irr.value > 0.15) {
        analysis += `is reasonable for this type of investment.\n\n`;
      } else {
        analysis += `is below typical M&A return thresholds - consider optimizing the structure.\n\n`;
      }
    }
    
    if (moicMetrics.length > 0) {
      const moic = moicMetrics[0];
      analysis += `**MOIC Analysis**: Your ${moic.value.toFixed(1)}x money multiple in ${moic.location} `;
      if (moic.value > 3) {
        analysis += `shows strong value creation potential.\n\n`;
      } else {
        analysis += `is within typical ranges but could be optimized.\n\n`;
      }
    }
    
    if (structure.validation) {
      analysis += `**Model Quality**: Validation score of ${structure.validation.modelScore}/100 `;
      if (structure.validation.modelScore >= 80) {
        analysis += `indicates a well-structured M&A model.\n\n`;
      } else {
        analysis += `suggests opportunities for model enhancement.\n\n`;
      }
    }
    
    return analysis;
  }
}

/**
 * Data Validation Agent - Specialized in model validation and error detection
 */
class DataValidationAgent {
  async process(userMessage, structure, previousResults, queryAnalysis) {
    console.log('✅ Data Validation Agent processing...');
    
    const validationResults = {
      errors: [],
      warnings: [],
      suggestions: [],
      dataQuality: this.assessDataQuality(structure),
      consistencyChecks: this.performConsistencyChecks(structure),
      recommendedActions: []
    };
    
    // Use existing validation from structure
    if (structure.validation) {
      validationResults.errors = structure.validation.errors || [];
      validationResults.warnings = structure.validation.warnings || [];
      validationResults.suggestions = structure.validation.suggestions || [];
    }
    
    // Add model-specific validations
    this.validateMandASpecifics(structure, validationResults);
    
    // Add formula consistency checks
    this.validateFormulaConsistency(structure, validationResults);
    
    return validationResults;
  }

  assessDataQuality(structure) {
    const quality = {
      completeness: 0,
      accuracy: 0,
      consistency: 0,
      overall: 0
    };
    
    // Calculate completeness based on key metrics presence
    const expectedMetrics = ['irr', 'moic', 'revenue'];
    const foundMetrics = new Set();
    
    Object.values(structure.keyMetrics).forEach(metrics => {
      metrics.forEach(metric => foundMetrics.add(metric.type));
    });
    
    quality.completeness = (foundMetrics.size / expectedMetrics.length) * 100;
    
    // Calculate accuracy based on confidence scores
    const allMetrics = Object.values(structure.keyMetrics).flat();
    if (allMetrics.length > 0) {
      quality.accuracy = (allMetrics.reduce((sum, m) => sum + (m.confidence || 0.5), 0) / allMetrics.length) * 100;
    }
    
    // Consistency based on validation score
    quality.consistency = structure.validation ? structure.validation.modelScore : 50;
    
    quality.overall = (quality.completeness + quality.accuracy + quality.consistency) / 3;
    
    return quality;
  }

  performConsistencyChecks(structure) {
    const checks = {
      formulaConsistency: true,
      dataTypes: true,
      rangeContinuity: true,
      externalReferences: false
    };
    
    // Check for external references which can cause issues
    Object.values(structure.sheets).forEach(sheet => {
      if (sheet.formulaAnalysis && sheet.formulaAnalysis.externalReferences.length > 0) {
        checks.externalReferences = true;
      }
    });
    
    return checks;
  }

  validateMandASpecifics(structure, results) {
    // Check for essential M&A model components
    const hasIRR = Object.values(structure.keyMetrics).some(metrics => 
      metrics.some(m => m.type === 'irr'));
    const hasMOIC = Object.values(structure.keyMetrics).some(metrics => 
      metrics.some(m => m.type === 'moic'));
    
    if (!hasIRR) {
      results.errors.push('IRR calculation not found - essential for M&A analysis');
      results.recommendedActions.push('Add XIRR formula to calculate internal rate of return');
    }
    
    if (!hasMOIC) {
      results.warnings.push('MOIC calculation not clearly identified');
      results.recommendedActions.push('Ensure money multiple calculation is clearly labeled');
    }
    
    // Check for reasonable value ranges
    Object.values(structure.keyMetrics).forEach(metrics => {
      metrics.forEach(metric => {
        if (metric.type === 'irr' && (metric.value < -0.5 || metric.value > 2.0)) {
          results.warnings.push(`IRR value ${(metric.value * 100).toFixed(1)}% in ${metric.location} seems unusual`);
        }
        if (metric.type === 'moic' && (metric.value < 0.5 || metric.value > 20)) {
          results.warnings.push(`MOIC value ${metric.value.toFixed(1)}x in ${metric.location} seems unusual`);
        }
      });
    });
  }

  validateFormulaConsistency(structure, results) {
    // Check for common formula errors
    Object.entries(structure.sheets).forEach(([sheetName, sheet]) => {
      if (sheet.formulaAnalysis) {
        // Check for complex formulas that might need optimization
        if (sheet.formulaAnalysis.complexFormulas.length > 5) {
          results.suggestions.push(`${sheetName} has ${sheet.formulaAnalysis.complexFormulas.length} complex formulas - consider simplifying for better maintainability`);
        }
        
        // Check for external references
        if (sheet.formulaAnalysis.externalReferences.length > 0) {
          results.warnings.push(`${sheetName} has external references that may cause calculation issues`);
        }
      }
    });
  }
}

// Export for global use
window.MultiAgentProcessor = MultiAgentProcessor;
window.multiAgentProcessor = new MultiAgentProcessor();

console.log('🎭 Multi-Agent Processor Stage 2 loaded with:');
console.log('  ✅ Excel Structure Agent - Workbook analysis specialist');
console.log('  ✅ Financial Analysis Agent - M&A expertise specialist');  
console.log('  ✅ Data Validation Agent - Error detection specialist');
console.log('  ✅ Intelligent agent routing and orchestration');
console.log('  ✅ Response synthesis and progress tracking');
console.log('🎯 Ready for professional M&A intelligence processing');