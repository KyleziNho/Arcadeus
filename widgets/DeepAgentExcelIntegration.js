/**
 * Deep Agent Excel Integration
 * Integrates MCP-style Excel tools with Deep Agent architecture
 * Combines best of both approaches
 */

class DeepAgentExcelIntegration {
  constructor(apiKey) {
    this.apiKey = apiKey;
    this.excelTools = new ExcelToolRegistry();
    this.fileSystem = new VirtualFileSystem();
    this.todoList = [];
    
    this.initializeEnhancedTools();
  }

  /**
   * Initialize enhanced tools that combine Deep Agent patterns with MCP tool structure
   */
  initializeEnhancedTools() {
    this.tools = {
      // Planning tool from Deep Agent
      todo_write: {
        name: 'todo_write',
        description: 'Create or update your task plan. ALWAYS use this first.',
        execute: this.todoWrite.bind(this)
      },
      
      // File system tools from Deep Agent
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
      
      // MCP-style Excel tools (enhanced)
      excel_read_data: {
        name: 'excel_read_data',
        description: 'Read data from Excel range with banking intelligence',
        execute: this.enhancedReadData.bind(this)
      },
      
      excel_write_data: {
        name: 'excel_write_data',
        description: 'Write data to Excel with validation',
        execute: this.enhancedWriteData.bind(this)
      },
      
      excel_format_range: {
        name: 'excel_format_range',
        description: 'Format Excel cells with banking standards',
        execute: this.enhancedFormatRange.bind(this)
      },
      
      excel_apply_formula: {
        name: 'excel_apply_formula',
        description: 'Apply formula with safety validation',
        execute: this.enhancedApplyFormula.bind(this)
      },
      
      // Investment banking specific tools
      banking_calculate_irr: {
        name: 'banking_calculate_irr',
        description: 'Calculate IRR with banking context',
        execute: this.bankingCalculateIRR.bind(this)
      },
      
      banking_validate_model: {
        name: 'banking_validate_model',
        description: 'Validate financial model for common errors',
        execute: this.bankingValidateModel.bind(this)
      },
      
      banking_find_metrics: {
        name: 'banking_find_metrics',
        description: 'Find key financial metrics in the model',
        execute: this.bankingFindMetrics.bind(this)
      }
    };
  }

  /**
   * Enhanced Excel tools that add banking intelligence to MCP patterns
   */
  
  async enhancedReadData(args) {
    const { sheetName, startCell, endCell, analyzeForBanking = true } = args;
    
    try {
      // Use MCP tool for basic read
      const tool = this.excelTools.getTool('read_data');
      const result = await tool.execute({ sheetName, startCell, endCell });
      
      if (!result.success) {
        return result;
      }
      
      // Add banking intelligence
      if (analyzeForBanking) {
        const analysis = await this.analyzeBankingData(result.data);
        result.bankingAnalysis = analysis;
        
        // Store in file system for later reference
        await this.writeFile({
          filename: `data_analysis_${sheetName}_${Date.now()}.json`,
          content: JSON.stringify({
            rawData: result.data,
            analysis: analysis,
            timestamp: new Date().toISOString()
          })
        });
      }
      
      return result;
      
    } catch (error) {
      return {
        success: false,
        error: `Enhanced read data failed: ${error.message}`,
        tool: 'enhanced_read_data'
      };
    }
  }
  
  async enhancedWriteData(args) {
    const { sheetName, startCell, data, validateBeforeWrite = true } = args;
    
    try {
      // Validate data for banking standards
      if (validateBeforeWrite) {
        const validation = await this.validateBankingData(data);
        if (!validation.valid) {
          return {
            success: false,
            error: `Data validation failed: ${validation.errors.join(', ')}`,
            tool: 'enhanced_write_data'
          };
        }
      }
      
      // Use MCP tool for actual write
      const tool = this.excelTools.getTool('write_data');
      const result = await tool.execute({ sheetName, startCell, data });
      
      if (result.success) {
        // Log the operation
        await this.writeFile({
          filename: `write_log_${Date.now()}.json`,
          content: JSON.stringify({
            operation: 'write_data',
            sheet: sheetName,
            range: result.range,
            rowsWritten: result.rowsWritten,
            timestamp: new Date().toISOString()
          })
        });
      }
      
      return result;
      
    } catch (error) {
      return {
        success: false,
        error: `Enhanced write data failed: ${error.message}`,
        tool: 'enhanced_write_data'
      };
    }
  }
  
  async enhancedFormatRange(args) {
    const { sheetName, startCell, endCell, ...formatOptions } = args;
    
    try {
      // Apply banking standard formatting if not specified
      const bankingFormats = this.getBankingStandardFormats(formatOptions);
      
      const tool = this.excelTools.getTool('format_range');
      const result = await tool.execute({
        sheetName,
        startCell,
        endCell,
        ...bankingFormats
      });
      
      if (result.success) {
        console.log(`Applied banking-standard formatting to ${result.range}`);
      }
      
      return result;
      
    } catch (error) {
      return {
        success: false,
        error: `Enhanced format range failed: ${error.message}`,
        tool: 'enhanced_format_range'
      };
    }
  }
  
  async enhancedApplyFormula(args) {
    const { sheetName, cell, formula, validateSafety = true } = args;
    
    try {
      // Additional safety validation beyond MCP server
      if (validateSafety) {
        const safetyCheck = await this.validateFormulaSafety(formula);
        if (!safetyCheck.safe) {
          return {
            success: false,
            error: `Formula safety check failed: ${safetyCheck.reason}`,
            tool: 'enhanced_apply_formula'
          };
        }
      }
      
      const tool = this.excelTools.getTool('apply_formula');
      const result = await tool.execute({ sheetName, cell, formula });
      
      if (result.success) {
        // Store formula for dependency tracking
        await this.writeFile({
          filename: `formulas_${sheetName}.json`,
          content: JSON.stringify({
            cell: cell,
            formula: formula,
            result: result.result,
            timestamp: new Date().toISOString()
          })
        });
      }
      
      return result;
      
    } catch (error) {
      return {
        success: false,
        error: `Enhanced apply formula failed: ${error.message}`,
        tool: 'enhanced_apply_formula'
      };
    }
  }

  /**
   * Investment Banking Specific Enhanced Tools
   */
  
  async bankingCalculateIRR(args) {
    const { sheetName, cashFlowRange, datesRange, analysis = true } = args;
    
    try {
      // Use the banking-specific IRR tool
      const tool = this.excelTools.getTool('calculate_irr');
      const result = await tool.execute({ sheetName, cashFlowRange, datesRange });
      
      if (result.success && analysis) {
        // Add banking context analysis
        const irrAnalysis = this.analyzeIRRResult(result.irr);
        result.bankingContext = irrAnalysis;
        
        // Store analysis
        await this.writeFile({
          filename: `irr_analysis_${Date.now()}.json`,
          content: JSON.stringify({
            irr: result.irr,
            formatted: result.formatted,
            analysis: irrAnalysis,
            ranges: { cashFlows: cashFlowRange, dates: datesRange },
            timestamp: new Date().toISOString()
          })
        });
      }
      
      return result;
      
    } catch (error) {
      return {
        success: false,
        error: `Banking IRR calculation failed: ${error.message}`,
        tool: 'banking_calculate_irr'
      };
    }
  }
  
  async bankingValidateModel(args) {
    const { sheetName, validationLevel = 'comprehensive' } = args;
    
    try {
      const validations = [];
      
      // Run multiple validation checks
      const checks = [
        this.checkCircularReferences(sheetName),
        this.checkFormulaConsistency(sheetName),
        this.validateFinancialAssumptions(sheetName),
        this.checkBankingStandards(sheetName)
      ];
      
      const results = await Promise.all(checks);
      
      const summary = {
        success: true,
        validationLevel: validationLevel,
        totalChecks: results.length,
        passed: results.filter(r => r.passed).length,
        failed: results.filter(r => !r.passed).length,
        details: results,
        recommendations: this.generateValidationRecommendations(results)
      };
      
      // Store validation report
      await this.writeFile({
        filename: `validation_report_${sheetName}_${Date.now()}.json`,
        content: JSON.stringify(summary, null, 2)
      });
      
      return summary;
      
    } catch (error) {
      return {
        success: false,
        error: `Model validation failed: ${error.message}`,
        tool: 'banking_validate_model'
      };
    }
  }
  
  async bankingFindMetrics(args) {
    const { sheetName, metricTypes = ['IRR', 'MOIC', 'EBITDA', 'Revenue'] } = args;
    
    try {
      const foundMetrics = {};
      
      for (const metric of metricTypes) {
        const locations = await this.findMetricInSheet(sheetName, metric);
        if (locations.length > 0) {
          foundMetrics[metric] = locations;
        }
      }
      
      const summary = {
        success: true,
        sheet: sheetName,
        metricsFound: Object.keys(foundMetrics).length,
        totalLocations: Object.values(foundMetrics).reduce((sum, locs) => sum + locs.length, 0),
        metrics: foundMetrics,
        suggestions: this.generateMetricSuggestions(foundMetrics)
      };
      
      // Store metrics map
      await this.writeFile({
        filename: `metrics_map_${sheetName}.json`,
        content: JSON.stringify(summary, null, 2)
      });
      
      return summary;
      
    } catch (error) {
      return {
        success: false,
        error: `Metric finding failed: ${error.message}`,
        tool: 'banking_find_metrics'
      };
    }
  }

  /**
   * Banking Intelligence Helper Methods
   */
  
  async analyzeBankingData(data) {
    // Analyze data for banking patterns
    const patterns = {
      hasFinancialData: false,
      likelyMetrics: [],
      dataQuality: 'unknown',
      recommendations: []
    };
    
    // Look for financial patterns
    for (const row of data) {
      for (const [cell, value] of Object.entries(row)) {
        if (typeof value === 'number') {
          patterns.hasFinancialData = true;
          
          // Check for common financial metric ranges
          if (value > 0.1 && value < 0.5) {
            patterns.likelyMetrics.push('Possible IRR/Return');
          }
          if (value > 1 && value < 10) {
            patterns.likelyMetrics.push('Possible Multiple');
          }
        }
      }
    }
    
    return patterns;
  }
  
  async validateBankingData(data) {
    const errors = [];
    
    // Check for common banking data issues
    for (let i = 0; i < data.length; i++) {
      for (let j = 0; j < data[i].length; j++) {
        const cell = data[i][j];
        
        // Check for obvious errors
        if (typeof cell === 'string' && cell.includes('#ERROR')) {
          errors.push(`Error formula in row ${i+1}, col ${j+1}`);
        }
        
        // Check for unrealistic values
        if (typeof cell === 'number' && Math.abs(cell) > 1000000) {
          errors.push(`Unusually large value in row ${i+1}, col ${j+1}: ${cell}`);
        }
      }
    }
    
    return {
      valid: errors.length === 0,
      errors: errors
    };
  }
  
  getBankingStandardFormats(userFormats) {
    // Apply Goldman Sachs style formatting defaults
    const defaults = {
      fontFamily: 'Arial',
      fontSize: 10,
      backgroundColor: userFormats.backgroundColor || '#FFFFFF',
      fontColor: userFormats.fontColor || '#000000',
      borderStyle: 'Thin'
    };
    
    return { ...defaults, ...userFormats };
  }
  
  async validateFormulaSafety(formula) {
    // Enhanced safety beyond MCP server validation
    const risks = [];
    
    // Check for potentially dangerous patterns
    if (formula.includes('INDIRECT')) {
      risks.push('INDIRECT function can cause performance issues');
    }
    
    if (formula.includes('VOLATILE')) {
      risks.push('Volatile functions cause unnecessary recalculation');
    }
    
    // Check complexity
    const functionCount = (formula.match(/[A-Z]+\(/g) || []).length;
    if (functionCount > 10) {
      risks.push('Formula is very complex and may be hard to debug');
    }
    
    return {
      safe: risks.length === 0,
      reason: risks.join('; ')
    };
  }
  
  analyzeIRRResult(irrValue) {
    if (typeof irrValue !== 'number') {
      return { interpretation: 'Error in calculation', risk: 'high' };
    }
    
    const annualIRR = irrValue;
    
    if (annualIRR < 0.05) {
      return { interpretation: 'Below market returns', risk: 'high', recommendation: 'Review assumptions' };
    } else if (annualIRR < 0.15) {
      return { interpretation: 'Market returns', risk: 'medium', recommendation: 'Acceptable for low-risk investments' };
    } else if (annualIRR < 0.30) {
      return { interpretation: 'Strong returns', risk: 'low', recommendation: 'Good investment opportunity' };
    } else {
      return { interpretation: 'Exceptionally high returns', risk: 'medium', recommendation: 'Verify assumptions - may be too optimistic' };
    }
  }

  /**
   * Integration with existing Deep Agent methods
   */
  
  async todoWrite(args) {
    const { todos } = args;
    this.todoList = todos;
    
    console.log('📋 Enhanced todo list updated:');
    todos.forEach(todo => {
      const status = todo.status === 'completed' ? '✅' : 
                    todo.status === 'in_progress' ? '🔄' : '⏳';
      console.log(`${status} ${todo.id}. ${todo.task}`);
    });
    
    return {
      success: true,
      message: 'Todo list updated with banking context',
      todos: this.todoList
    };
  }
  
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

  /**
   * Mock validation methods (implement with real logic)
   */
  
  async checkCircularReferences(sheetName) {
    return { passed: true, check: 'circular_references', message: 'No circular references found' };
  }
  
  async checkFormulaConsistency(sheetName) {
    return { passed: true, check: 'formula_consistency', message: 'Formulas are consistent' };
  }
  
  async validateFinancialAssumptions(sheetName) {
    return { passed: true, check: 'financial_assumptions', message: 'Assumptions are reasonable' };
  }
  
  async checkBankingStandards(sheetName) {
    return { passed: true, check: 'banking_standards', message: 'Meets banking standards' };
  }
  
  generateValidationRecommendations(results) {
    const failed = results.filter(r => !r.passed);
    if (failed.length === 0) {
      return ['Model validation passed all checks'];
    }
    
    return failed.map(f => `Fix ${f.check}: ${f.message}`);
  }
  
  async findMetricInSheet(sheetName, metric) {
    // Mock implementation - would search for metric patterns
    return [
      { cell: 'D25', value: 0.18, confidence: 0.9 }
    ];
  }
  
  generateMetricSuggestions(foundMetrics) {
    const suggestions = [];
    
    if (!foundMetrics.IRR) {
      suggestions.push('Consider adding IRR calculation for investment analysis');
    }
    if (!foundMetrics.MOIC) {
      suggestions.push('Add MOIC calculation for multiple analysis');
    }
    
    return suggestions;
  }

  /**
   * Main processing method that combines everything
   */
  async processRequest(userInput) {
    console.log('🧠 Deep Agent Excel Integration processing:', userInput);
    
    try {
      // Enhanced system prompt that understands both Deep Agent and MCP patterns
      const systemPrompt = `You are an expert Excel analyst with comprehensive banking knowledge and tool access.

AVAILABLE TOOLS:
- Planning: todo_write (ALWAYS use first)
- File System: write_file, read_file  
- Excel Data: excel_read_data, excel_write_data
- Excel Formatting: excel_format_range
- Excel Formulas: excel_apply_formula  
- Banking Analysis: banking_calculate_irr, banking_validate_model, banking_find_metrics

APPROACH:
1. Create a plan using todo_write
2. Use Excel tools to gather data and context
3. Apply banking intelligence and validation
4. Store results in file system
5. Provide comprehensive analysis

Focus on investment banking best practices and Goldman Sachs standards.`;

      const messages = [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userInput }
      ];

      // Add context from file system
      if (this.fileSystem.files.size > 0) {
        const fileList = Array.from(this.fileSystem.files.keys()).slice(0, 5).join(', ');
        messages[0].content += `\n\nAvailable files: ${fileList}`;
      }

      // Process with enhanced reasoning
      const response = await this.deepReasoningLoop(messages);
      
      return {
        success: true,
        response: response,
        todoList: this.todoList,
        files: Array.from(this.fileSystem.files.entries()),
        toolsUsed: Object.keys(this.tools),
        timestamp: new Date().toISOString()
      };

    } catch (error) {
      console.error('❌ Deep Agent Excel Integration error:', error);
      return {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  }

  /**
   * Execute tool with enhanced error handling
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

  // Placeholder for deep reasoning loop (implement based on your existing Deep Agent)
  async deepReasoningLoop(messages) {
    // This would implement the same pattern as your existing Deep Agent
    // but with access to the enhanced Excel tools
    return "Enhanced Deep Agent processing complete with MCP-style Excel tools.";
  }
}

// Export for use
window.DeepAgentExcelIntegration = DeepAgentExcelIntegration;

console.log('✅ Deep Agent Excel Integration ready with MCP tool patterns');