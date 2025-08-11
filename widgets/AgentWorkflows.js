/**
 * Agent Workflows - Structured task execution patterns
 * Implements OpenAI's best practices for agent orchestration
 */

/**
 * Base Workflow Class
 */
class BaseWorkflow {
    constructor(name) {
        this.name = name;
        this.steps = [];
        this.requiresConfirmation = false;
    }

    /**
     * Execute workflow with proper error handling and logging
     */
    async execute(analysis, context, tools) {
        const workflowId = this.generateWorkflowId();
        console.log(`🔄 Starting workflow: ${this.name} (${workflowId})`);

        try {
            // Pre-execution validation
            const validation = await this.validate(analysis, context);
            if (!validation.valid) {
                return this.createErrorResponse(validation.reason);
            }

            // Request confirmation if needed
            if (this.requiresConfirmation || analysis.requiresConfirmation) {
                const confirmation = await this.requestConfirmation(analysis, context);
                if (!confirmation.confirmed) {
                    return this.createCancelledResponse(confirmation.reason);
                }
            }

            // Execute workflow steps
            const result = await this.executeSteps(analysis, context, tools);
            
            console.log(`✅ Workflow completed: ${this.name}`);
            return result;

        } catch (error) {
            console.error(`❌ Workflow failed: ${this.name}`, error);
            return this.createErrorResponse(error.message);
        }
    }

    /**
     * Validate workflow can be executed
     */
    async validate(analysis, context) {
        return { valid: true };
    }

    /**
     * Request user confirmation for risky operations
     */
    async requestConfirmation(analysis, context) {
        // In a real implementation, this would show a confirmation dialog
        // For now, we'll assume confirmation based on risk level
        if (analysis.riskLevel === 'high') {
            return {
                confirmed: false,
                reason: 'High-risk operation requires explicit confirmation'
            };
        }
        return { confirmed: true };
    }

    /**
     * Execute the workflow steps
     */
    async executeSteps(analysis, context, tools) {
        throw new Error('executeSteps must be implemented by subclass');
    }

    createErrorResponse(message) {
        return {
            success: false,
            type: 'workflow_error',
            message: message,
            workflowName: this.name
        };
    }

    createCancelledResponse(reason) {
        return {
            success: false,
            type: 'workflow_cancelled',
            message: `Operation cancelled: ${reason}`,
            workflowName: this.name
        };
    }

    generateWorkflowId() {
        return `${this.name}_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`;
    }
}

/**
 * Format Cells Workflow
 */
class FormatCellsWorkflow extends BaseWorkflow {
    constructor() {
        super('format_cells');
        this.requiresConfirmation = false;
    }

    async executeSteps(analysis, context, tools) {
        const formatter = tools.getTool('cellFormatter');
        if (!formatter) {
            throw new Error('Cell formatter tool not available');
        }

        return await Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                
                // Determine target range
                const targetRange = await this.determineTargetRange(analysis, worksheet, context);
                
                // Apply formatting
                const formatResult = await this.applyFormatting(analysis, targetRange, context);
                
                await context.sync();
                
                return {
                    success: true,
                    type: 'format_cells_success',
                    message: `✅ Successfully formatted ${formatResult.cellsChanged} cells`,
                    actionsTaken: formatResult.actions,
                    workflowName: this.name
                };

            } catch (error) {
                throw new Error(`Format cells operation failed: ${error.message}`);
            }
        });
    }

    async determineTargetRange(analysis, worksheet, context) {
        // If specific ranges mentioned, use those
        if (analysis.extractedEntities.ranges && analysis.extractedEntities.ranges.length > 0) {
            return worksheet.getRange(analysis.extractedEntities.ranges[0]);
        }

        // If no range specified, use selected range or find headers
        try {
            const selection = context.workbook.getSelectedRange();
            selection.load(['address']);
            await context.sync();
            
            if (selection.address !== 'A1') {
                return selection;
            }
        } catch (e) {
            // Fall back to finding headers
        }

        // Default: search for header-like cells
        return await this.findHeaderCells(worksheet, context, analysis);
    }

    async findHeaderCells(worksheet, context, analysis) {
        const usedRange = worksheet.getUsedRange();
        usedRange.load(['values', 'rowCount', 'columnCount']);
        await context.sync();

        // Look for cells in first few rows that might be headers
        const headerCandidates = [];
        const maxRowsToCheck = Math.min(5, usedRange.rowCount);

        for (let row = 0; row < maxRowsToCheck; row++) {
            for (let col = 0; col < usedRange.columnCount; col++) {
                const cellValue = usedRange.values[row][col];
                if (cellValue && typeof cellValue === 'string' && cellValue.length > 0) {
                    const cellAddress = this.getCellAddress(row, col);
                    headerCandidates.push(cellAddress);
                }
            }
        }

        // Return range covering header candidates
        if (headerCandidates.length > 0) {
            const firstCell = headerCandidates[0];
            const lastCell = headerCandidates[headerCandidates.length - 1];
            return worksheet.getRange(`${firstCell}:${lastCell}`);
        }

        // Fallback to A1
        return worksheet.getRange('A1');
    }

    async applyFormatting(analysis, range, context) {
        const actions = [];
        let cellsChanged = 0;

        range.load(['address', 'cellCount']);
        await context.sync();
        cellsChanged = range.cellCount;

        // Apply color changes
        if (analysis.extractedEntities.colors && analysis.extractedEntities.colors.length > 0) {
            const targetColor = this.mapColorToHex(analysis.extractedEntities.colors[0]);
            range.format.fill.color = targetColor;
            range.format.font.color = this.getContrastingTextColor(targetColor);
            actions.push(`Changed background color to ${analysis.extractedEntities.colors[0]}`);
        }

        // Apply text formatting
        if (analysis.message.toLowerCase().includes('bold')) {
            range.format.font.bold = true;
            actions.push('Applied bold formatting');
        }

        if (analysis.message.toLowerCase().includes('italic')) {
            range.format.font.italic = true;
            actions.push('Applied italic formatting');
        }

        return { actions, cellsChanged };
    }

    mapColorToHex(colorName) {
        const colorMap = {
            'red': '#dc2626',
            'green': '#16a34a',
            'blue': '#2563eb',
            'yellow': '#eab308',
            'orange': '#ea580c',
            'purple': '#7c3aed',
            'pink': '#db2777',
            'gray': '#6b7280',
            'black': '#000000',
            'white': '#ffffff'
        };
        return colorMap[colorName.toLowerCase()] || '#16a34a';
    }

    getContrastingTextColor(backgroundColor) {
        const darkColors = ['#dc2626', '#16a34a', '#2563eb', '#7c3aed', '#db2777', '#000000'];
        return darkColors.includes(backgroundColor) ? '#ffffff' : '#000000';
    }

    getCellAddress(row, col) {
        return String.fromCharCode(65 + col) + (row + 1);
    }
}

/**
 * Analyze Data Workflow - Safe read-only operations
 */
class AnalyzeDataWorkflow extends BaseWorkflow {
    constructor() {
        super('analyze_data');
        this.requiresConfirmation = false;
    }

    async executeSteps(analysis, context, tools) {
        const analyzer = tools.getTool('dataAnalyzer');
        if (!analyzer) {
            throw new Error('Data analyzer tool not available');
        }

        return await Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const usedRange = worksheet.getUsedRange();
                
                usedRange.load(['values', 'formulas', 'address', 'rowCount', 'columnCount']);
                await context.sync();

                const analysisResult = await this.analyzeSpreadsheetData(usedRange, analysis);
                
                return {
                    success: true,
                    type: 'analyze_data_success',
                    message: analysisResult.summary,
                    data: analysisResult.data,
                    insights: analysisResult.insights,
                    workflowName: this.name
                };

            } catch (error) {
                throw new Error(`Data analysis failed: ${error.message}`);
            }
        });
    }

    async analyzeSpreadsheetData(range, analysis) {
        const data = {
            totalCells: range.rowCount * range.columnCount,
            nonEmptyCells: 0,
            numericCells: 0,
            formulaCells: 0,
            textCells: 0
        };

        const insights = [];
        let sampleValues = [];

        // Analyze cell contents
        for (let row = 0; row < range.rowCount; row++) {
            for (let col = 0; col < range.columnCount; col++) {
                const value = range.values[row][col];
                const formula = range.formulas[row][col];

                if (value !== null && value !== undefined && value !== '') {
                    data.nonEmptyCells++;
                    
                    if (formula && formula.startsWith('=')) {
                        data.formulaCells++;
                    } else if (typeof value === 'number') {
                        data.numericCells++;
                        if (sampleValues.length < 10) {
                            sampleValues.push(value);
                        }
                    } else {
                        data.textCells++;
                    }
                }
            }
        }

        // Generate insights
        const dataUtilization = (data.nonEmptyCells / data.totalCells) * 100;
        insights.push(`Data utilization: ${dataUtilization.toFixed(1)}% (${data.nonEmptyCells}/${data.totalCells} cells)`);

        if (data.formulaCells > 0) {
            insights.push(`Contains ${data.formulaCells} formula cells - this appears to be a calculation sheet`);
        }

        if (sampleValues.length > 0) {
            const avg = sampleValues.reduce((a, b) => a + b, 0) / sampleValues.length;
            insights.push(`Numeric data average: ${avg.toFixed(2)}`);
        }

        return {
            summary: `📊 Analyzed ${range.address} containing ${data.nonEmptyCells} data cells`,
            data: data,
            insights: insights
        };
    }
}

/**
 * Modify Data Workflow - Requires confirmation for data changes
 */
class ModifyDataWorkflow extends BaseWorkflow {
    constructor() {
        super('modify_data');
        this.requiresConfirmation = true;
    }

    async validate(analysis, context) {
        // Ensure we have specific targets
        if (!analysis.extractedEntities.ranges && !context.selectedRange) {
            return {
                valid: false,
                reason: 'No specific range selected for data modification'
            };
        }
        return { valid: true };
    }

    async executeSteps(analysis, context, tools) {
        return {
            success: false,
            type: 'requires_confirmation',
            message: '⚠️ Data modification requires user confirmation. Please confirm you want to proceed with this change.',
            workflowName: this.name
        };
    }
}

/**
 * Delete Data Workflow - High-risk operations
 */
class DeleteDataWorkflow extends BaseWorkflow {
    constructor() {
        super('delete_data');
        this.requiresConfirmation = true;
    }

    async validate(analysis, context) {
        return {
            valid: false,
            reason: 'Delete operations are currently disabled for safety'
        };
    }

    async executeSteps(analysis, context, tools) {
        return {
            success: false,
            type: 'operation_blocked',
            message: '🚫 Delete operations are blocked for safety. Please use Excel directly for data deletion.',
            workflowName: this.name
        };
    }
}

/**
 * Add Formula Workflow
 */
class AddFormulaWorkflow extends BaseWorkflow {
    constructor() {
        super('add_formula');
        this.requiresConfirmation = false;
    }

    async executeSteps(analysis, context, tools) {
        return {
            success: false,
            type: 'not_implemented',
            message: '🔧 Formula building is in development. Currently only formatting and analysis are available.',
            workflowName: this.name
        };
    }
}

// Export workflows
window.FormatCellsWorkflow = FormatCellsWorkflow;
window.AnalyzeDataWorkflow = AnalyzeDataWorkflow;
window.ModifyDataWorkflow = ModifyDataWorkflow;
window.DeleteDataWorkflow = DeleteDataWorkflow;
window.AddFormulaWorkflow = AddFormulaWorkflow;

console.log('🔄 Agent Workflows loaded');