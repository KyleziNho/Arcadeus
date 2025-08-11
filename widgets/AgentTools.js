/**
 * Agent Tools - Standardized tool implementations
 * Following OpenAI's structured tool approach
 */

/**
 * Base Tool Class
 */
class BaseTool {
    constructor(name, description) {
        this.name = name;
        this.description = description;
        this.schema = null;
        this.isInitialized = false;
    }

    async initialize() {
        this.isInitialized = true;
        console.log(`🔧 Tool initialized: ${this.name}`);
    }

    async execute(parameters, context) {
        if (!this.isInitialized) {
            throw new Error(`Tool ${this.name} not initialized`);
        }
        
        // Validate parameters against schema
        const validation = this.validateParameters(parameters);
        if (!validation.valid) {
            throw new Error(`Invalid parameters: ${validation.errors.join(', ')}`);
        }

        return await this.run(parameters, context);
    }

    validateParameters(parameters) {
        // Basic validation - can be enhanced with JSON Schema
        return { valid: true, errors: [] };
    }

    async run(parameters, context) {
        throw new Error('run method must be implemented by subclass');
    }

    createSuccessResult(data, message) {
        return {
            success: true,
            tool: this.name,
            data: data,
            message: message,
            timestamp: new Date().toISOString()
        };
    }

    createErrorResult(error) {
        return {
            success: false,
            tool: this.name,
            error: error,
            timestamp: new Date().toISOString()
        };
    }
}

/**
 * Cell Formatter Tool
 */
class CellFormatterTool extends BaseTool {
    constructor() {
        super('cellFormatter', 'Formats Excel cells with colors, fonts, and styles');
        this.schema = {
            type: 'object',
            properties: {
                range: { type: 'string', description: 'Target cell range (e.g., A1:C3)' },
                backgroundColor: { type: 'string', description: 'Background color hex code' },
                fontColor: { type: 'string', description: 'Font color hex code' },
                bold: { type: 'boolean', description: 'Apply bold formatting' },
                italic: { type: 'boolean', description: 'Apply italic formatting' }
            }
        };
    }

    async run(parameters, context) {
        try {
            return await Excel.run(async (excelContext) => {
                const worksheet = excelContext.workbook.worksheets.getActiveWorksheet();
                const range = worksheet.getRange(parameters.range || 'A1');
                
                range.load(['address', 'cellCount']);
                
                // Apply formatting
                if (parameters.backgroundColor) {
                    range.format.fill.color = parameters.backgroundColor;
                }
                
                if (parameters.fontColor) {
                    range.format.font.color = parameters.fontColor;
                }
                
                if (parameters.bold) {
                    range.format.font.bold = true;
                }
                
                if (parameters.italic) {
                    range.format.font.italic = true;
                }
                
                await excelContext.sync();
                
                return this.createSuccessResult({
                    range: range.address,
                    cellCount: range.cellCount,
                    formattingApplied: parameters
                }, `Formatted ${range.cellCount} cells in range ${range.address}`);
            });
        } catch (error) {
            return this.createErrorResult(error.message);
        }
    }
}

/**
 * Data Analyzer Tool
 */
class DataAnalyzerTool extends BaseTool {
    constructor() {
        super('dataAnalyzer', 'Analyzes Excel data and provides insights');
        this.schema = {
            type: 'object',
            properties: {
                range: { type: 'string', description: 'Range to analyze (optional, defaults to used range)' },
                includeFormulas: { type: 'boolean', description: 'Include formula analysis' },
                includeStatistics: { type: 'boolean', description: 'Calculate statistics for numeric data' }
            }
        };
    }

    async run(parameters, context) {
        try {
            return await Excel.run(async (excelContext) => {
                const worksheet = excelContext.workbook.worksheets.getActiveWorksheet();
                const range = parameters.range ? 
                    worksheet.getRange(parameters.range) : 
                    worksheet.getUsedRange();
                
                range.load(['values', 'formulas', 'address', 'rowCount', 'columnCount']);
                await excelContext.sync();
                
                const analysis = this.performDataAnalysis(range, parameters);
                
                return this.createSuccessResult(analysis, 
                    `Analyzed ${analysis.summary.totalCells} cells in range ${range.address}`);
            });
        } catch (error) {
            return this.createErrorResult(error.message);
        }
    }

    performDataAnalysis(range, parameters) {
        const analysis = {
            range: range.address,
            summary: {
                totalCells: range.rowCount * range.columnCount,
                nonEmptyCells: 0,
                numericCells: 0,
                textCells: 0,
                formulaCells: 0
            },
            data: {
                values: [],
                formulas: parameters.includeFormulas ? [] : null
            },
            statistics: parameters.includeStatistics ? {} : null,
            insights: []
        };

        const numericValues = [];

        // Analyze each cell
        for (let row = 0; row < range.rowCount; row++) {
            for (let col = 0; col < range.columnCount; col++) {
                const value = range.values[row][col];
                const formula = range.formulas[row][col];

                if (value !== null && value !== undefined && value !== '') {
                    analysis.summary.nonEmptyCells++;
                    
                    if (formula && formula.startsWith('=')) {
                        analysis.summary.formulaCells++;
                        if (parameters.includeFormulas) {
                            analysis.data.formulas.push({
                                cell: this.getCellAddress(row, col),
                                formula: formula,
                                value: value
                            });
                        }
                    } else if (typeof value === 'number') {
                        analysis.summary.numericCells++;
                        numericValues.push(value);
                    } else {
                        analysis.summary.textCells++;
                    }

                    // Sample first 20 non-empty values
                    if (analysis.data.values.length < 20) {
                        analysis.data.values.push({
                            cell: this.getCellAddress(row, col),
                            value: value,
                            type: typeof value
                        });
                    }
                }
            }
        }

        // Calculate statistics for numeric data
        if (parameters.includeStatistics && numericValues.length > 0) {
            analysis.statistics = this.calculateStatistics(numericValues);
        }

        // Generate insights
        analysis.insights = this.generateInsights(analysis.summary, numericValues);

        return analysis;
    }

    calculateStatistics(values) {
        const sorted = [...values].sort((a, b) => a - b);
        const sum = values.reduce((a, b) => a + b, 0);
        const mean = sum / values.length;
        
        return {
            count: values.length,
            sum: sum,
            mean: mean,
            median: sorted[Math.floor(sorted.length / 2)],
            min: Math.min(...values),
            max: Math.max(...values),
            range: Math.max(...values) - Math.min(...values)
        };
    }

    generateInsights(summary, numericValues) {
        const insights = [];
        
        const utilizationRate = (summary.nonEmptyCells / summary.totalCells) * 100;
        insights.push({
            type: 'utilization',
            message: `Data utilization: ${utilizationRate.toFixed(1)}%`,
            value: utilizationRate
        });

        if (summary.formulaCells > 0) {
            insights.push({
                type: 'complexity',
                message: `Contains ${summary.formulaCells} calculated fields`,
                value: summary.formulaCells
            });
        }

        if (numericValues.length > 0) {
            const avg = numericValues.reduce((a, b) => a + b, 0) / numericValues.length;
            insights.push({
                type: 'numeric_summary',
                message: `Average of numeric values: ${avg.toFixed(2)}`,
                value: avg
            });
        }

        return insights;
    }

    getCellAddress(row, col) {
        return String.fromCharCode(65 + col) + (row + 1);
    }
}

/**
 * Data Modifier Tool - For safe data updates
 */
class DataModifierTool extends BaseTool {
    constructor() {
        super('dataModifier', 'Safely modifies Excel data with validation');
        this.schema = {
            type: 'object',
            properties: {
                range: { type: 'string', description: 'Target range for modification' },
                operation: { type: 'string', enum: ['set', 'clear', 'increment'], description: 'Type of modification' },
                value: { description: 'New value to set (if operation is set)' },
                incrementBy: { type: 'number', description: 'Amount to increment by (if operation is increment)' }
            },
            required: ['range', 'operation']
        };
    }

    validateParameters(parameters) {
        const errors = [];
        
        if (!parameters.range) {
            errors.push('range is required');
        }
        
        if (!parameters.operation || !['set', 'clear', 'increment'].includes(parameters.operation)) {
            errors.push('operation must be one of: set, clear, increment');
        }
        
        if (parameters.operation === 'set' && parameters.value === undefined) {
            errors.push('value is required when operation is set');
        }
        
        if (parameters.operation === 'increment' && typeof parameters.incrementBy !== 'number') {
            errors.push('incrementBy must be a number when operation is increment');
        }

        return { valid: errors.length === 0, errors };
    }

    async run(parameters, context) {
        try {
            return await Excel.run(async (excelContext) => {
                const worksheet = excelContext.workbook.worksheets.getActiveWorksheet();
                const range = worksheet.getRange(parameters.range);
                
                range.load(['address', 'values', 'cellCount']);
                await excelContext.sync();

                // Store original values for rollback
                const originalValues = range.values;
                const result = { 
                    range: range.address,
                    cellsModified: range.cellCount,
                    operation: parameters.operation,
                    originalValues: originalValues
                };

                // Perform modification
                switch (parameters.operation) {
                    case 'set':
                        range.values = [[parameters.value]];
                        result.newValue = parameters.value;
                        break;
                        
                    case 'clear':
                        range.clear();
                        result.newValue = null;
                        break;
                        
                    case 'increment':
                        // Only increment numeric cells
                        const newValues = originalValues.map(row =>
                            row.map(cell => 
                                typeof cell === 'number' ? cell + parameters.incrementBy : cell
                            )
                        );
                        range.values = newValues;
                        result.incrementBy = parameters.incrementBy;
                        break;
                }

                await excelContext.sync();
                
                return this.createSuccessResult(result, 
                    `${parameters.operation} operation completed on ${result.cellsModified} cells`);
            });
        } catch (error) {
            return this.createErrorResult(error.message);
        }
    }
}

/**
 * Formula Builder Tool - For creating Excel formulas
 */
class FormulaBuilderTool extends BaseTool {
    constructor() {
        super('formulaBuilder', 'Creates and inserts Excel formulas');
        this.schema = {
            type: 'object',
            properties: {
                targetCell: { type: 'string', description: 'Cell where formula will be placed' },
                formulaType: { type: 'string', enum: ['sum', 'average', 'count', 'custom'], description: 'Type of formula' },
                sourceRange: { type: 'string', description: 'Range to calculate from' },
                customFormula: { type: 'string', description: 'Custom formula (if formulaType is custom)' }
            },
            required: ['targetCell', 'formulaType']
        };
    }

    async run(parameters, context) {
        try {
            return await Excel.run(async (excelContext) => {
                const worksheet = excelContext.workbook.worksheets.getActiveWorksheet();
                const targetCell = worksheet.getRange(parameters.targetCell);
                
                let formula;
                switch (parameters.formulaType) {
                    case 'sum':
                        formula = `=SUM(${parameters.sourceRange})`;
                        break;
                    case 'average':
                        formula = `=AVERAGE(${parameters.sourceRange})`;
                        break;
                    case 'count':
                        formula = `=COUNT(${parameters.sourceRange})`;
                        break;
                    case 'custom':
                        formula = parameters.customFormula;
                        break;
                    default:
                        throw new Error('Invalid formula type');
                }

                targetCell.formulas = [[formula]];
                targetCell.load(['address', 'values']);
                await excelContext.sync();

                return this.createSuccessResult({
                    targetCell: targetCell.address,
                    formula: formula,
                    result: targetCell.values[0][0]
                }, `Formula ${formula} added to cell ${targetCell.address}`);
            });
        } catch (error) {
            return this.createErrorResult(error.message);
        }
    }
}

// Export tools
window.CellFormatterTool = CellFormatterTool;
window.DataAnalyzerTool = DataAnalyzerTool;
window.DataModifierTool = DataModifierTool;
window.FormulaBuilderTool = FormulaBuilderTool;

console.log('🛠️ Agent Tools loaded');