/**
 * Direct Excel Actions - Simple, reliable Excel manipulation
 * Bypasses complex agent system for immediate Excel operations
 */

class DirectExcelActions {
    constructor() {
        this.isReady = false;
        this.initialize();
    }

    async initialize() {
        try {
            if (typeof Office !== 'undefined' && Office.context) {
                this.isReady = true;
                console.log('✅ Direct Excel Actions ready');
            } else {
                console.log('⏳ Waiting for Office.js...');
                setTimeout(() => this.initialize(), 500);
            }
        } catch (error) {
            console.error('❌ Direct Excel Actions initialization failed:', error);
        }
    }

    /**
     * Detect if message is requesting Excel modifications
     */
    isExcelActionRequest(message) {
        const actionPatterns = [
            // Color changes
            /change.*color/i,
            /change.*background/i,
            /make.*green/i,
            /make.*red/i,
            /make.*blue/i,
            /make.*yellow/i,
            /color.*header/i,
            /highlight/i,
            /set.*background/i,
            
            // Text formatting  
            /make.*bold/i,
            /bold.*header/i,
            /make.*italic/i,
            /change.*font/i,
            
            // General formatting
            /format.*cell/i,
            /format.*header/i,
            /format.*range/i,
            /apply.*format/i,
            
            // Specific Excel actions
            /change.*the.*header/i,
            /format.*the.*header/i,
            /color.*the.*header/i,
            /make.*header.*bold/i,
            /change.*header.*color/i,
            
            // Conditional formatting hints
            /conditional.*format/i,
            /highlight.*if/i,
            /color.*based.*on/i
        ];

        const lowerMessage = message.toLowerCase();
        
        // Check patterns
        const hasActionPattern = actionPatterns.some(pattern => pattern.test(message));
        
        // Additional keyword checks
        const hasActionKeywords = (
            (lowerMessage.includes('change') && (lowerMessage.includes('color') || lowerMessage.includes('format'))) ||
            (lowerMessage.includes('make') && (lowerMessage.includes('bold') || lowerMessage.includes('green') || lowerMessage.includes('red'))) ||
            (lowerMessage.includes('header') && (lowerMessage.includes('color') || lowerMessage.includes('bold') || lowerMessage.includes('format'))) ||
            lowerMessage.includes('highlight') ||
            lowerMessage.includes('format')
        );
        
        const isActionRequest = hasActionPattern || hasActionKeywords;
        
        if (isActionRequest) {
            console.log(`🎯 Detected Excel action request: "${message}"`);
        }
        
        return isActionRequest;
    }

    /**
     * Execute Excel action based on user message
     */
    async executeAction(message) {
        if (!this.isReady) {
            throw new Error('Direct Excel Actions not ready');
        }

        console.log('🎯 Executing direct Excel action:', message);

        try {
            // Determine action type based on message content
            if (message.toLowerCase().includes('conditional') || message.toLowerCase().includes('highlight if') || message.toLowerCase().includes('based on')) {
                return await this.createConditionalFormat(message);
            } else if (message.toLowerCase().includes('color') || message.toLowerCase().includes('background') || message.toLowerCase().includes('highlight')) {
                return await this.changeColor(message);
            } else if (message.toLowerCase().includes('bold')) {
                return await this.makeBold(message);
            } else if (message.toLowerCase().includes('format')) {
                return await this.formatCells(message);
            } else {
                return await this.genericFormat(message);
            }
        } catch (error) {
            console.error('❌ Excel action failed:', error);
            return {
                success: false,
                message: `Failed to execute Excel action: ${error.message}`,
                error: error
            };
        }
    }

    /**
     * Change cell colors using proper Excel Add-in API
     */
    async changeColor(message) {
        return await Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                
                // Determine target range - prioritize selected range, fallback to header area
                let targetRange;
                let rangeDescription;
                
                try {
                    // Try to get selected range first
                    const selectedRange = context.workbook.getSelectedRange();
                    selectedRange.load(['address']);
                    await context.sync();
                    
                    if (selectedRange.address !== '$A$1') { // If something is actually selected
                        targetRange = selectedRange;
                        rangeDescription = `selected range ${selectedRange.address}`;
                    } else {
                        throw new Error('No selection, use default');
                    }
                } catch (e) {
                    // If no meaningful selection, target likely header areas
                    if (message.toLowerCase().includes('header')) {
                        targetRange = worksheet.getRange('A1:Z3'); // First 3 rows (typical headers)
                        rangeDescription = 'header area A1:Z3';
                    } else {
                        targetRange = worksheet.getRange('A1:E5'); // Small default range
                        rangeDescription = 'default range A1:E5';
                    }
                }

                // Load range properties to check what we're working with
                targetRange.load(['address', 'values']);
                await context.sync();

                // Extract target color from message
                const targetColor = this.extractColor(message);
                console.log(`🎨 Applying ${this.getColorName(targetColor)} to ${rangeDescription}`);

                // Apply formatting using Excel Add-in API
                targetRange.format.fill.color = targetColor;
                targetRange.format.font.color = this.getContrastingColor(targetColor);
                
                // If this is a header formatting request, also make it bold
                if (message.toLowerCase().includes('header') || message.toLowerCase().includes('bold')) {
                    targetRange.format.font.bold = true;
                }
                
                await context.sync();

                // Count non-empty cells that were affected
                let affectedCells = 0;
                const values = targetRange.values;
                for (let i = 0; i < values.length; i++) {
                    for (let j = 0; j < values[i].length; j++) {
                        if (values[i][j] !== null && values[i][j] !== undefined && values[i][j] !== '') {
                            affectedCells++;
                        }
                    }
                }

                return {
                    success: true,
                    message: `✅ Applied ${this.getColorName(targetColor)} formatting to ${rangeDescription}`,
                    action: 'color_change',
                    range: targetRange.address,
                    color: targetColor,
                    affectedCells: affectedCells,
                    details: `Changed background to ${this.getColorName(targetColor)} and adjusted text color for visibility`
                };

            } catch (error) {
                console.error('Color change error:', error);
                return {
                    success: false,
                    message: `❌ Failed to change colors: ${error.message}. Make sure you have an Excel file open with data.`
                };
            }
        });
    }

    /**
     * Make text bold
     */
    async makeBold(message) {
        return await Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                
                let targetRange;
                try {
                    targetRange = context.workbook.getSelectedRange();
                    await context.sync();
                } catch (e) {
                    targetRange = worksheet.getRange('A1:Z3');
                }

                targetRange.load(['address']);
                await context.sync();

                // Apply bold formatting
                targetRange.format.font.bold = true;
                await context.sync();

                return {
                    success: true,
                    message: `✅ Made text bold in range ${targetRange.address}`,
                    action: 'bold_format',
                    range: targetRange.address
                };

            } catch (error) {
                return {
                    success: false,
                    message: `❌ Failed to make text bold: ${error.message}`
                };
            }
        });
    }

    /**
     * Generic cell formatting
     */
    async formatCells(message) {
        return await Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                
                let targetRange;
                try {
                    targetRange = context.workbook.getSelectedRange();
                    await context.sync();
                } catch (e) {
                    targetRange = worksheet.getRange('A1:Z5');
                }

                targetRange.load(['address']);
                await context.sync();

                const actions = [];

                // Apply color if mentioned
                if (message.toLowerCase().includes('color') || message.toLowerCase().includes('background')) {
                    const color = this.extractColor(message);
                    targetRange.format.fill.color = color;
                    targetRange.format.font.color = this.getContrastingColor(color);
                    actions.push(`Changed background to ${this.getColorName(color)}`);
                }

                // Apply bold if mentioned
                if (message.toLowerCase().includes('bold')) {
                    targetRange.format.font.bold = true;
                    actions.push('Applied bold formatting');
                }

                // Apply italic if mentioned
                if (message.toLowerCase().includes('italic')) {
                    targetRange.format.font.italic = true;
                    actions.push('Applied italic formatting');
                }

                await context.sync();

                return {
                    success: true,
                    message: `✅ Applied formatting to ${targetRange.address}: ${actions.join(', ')}`,
                    action: 'format_cells',
                    range: targetRange.address,
                    changes: actions
                };

            } catch (error) {
                return {
                    success: false,
                    message: `❌ Failed to format cells: ${error.message}`
                };
            }
        });
    }

    /**
     * Create conditional formatting rules using Excel Add-in API
     */
    async createConditionalFormat(message) {
        return await Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                
                // Try to get selected range, fallback to data area
                let targetRange;
                try {
                    const selectedRange = context.workbook.getSelectedRange();
                    selectedRange.load(['address']);
                    await context.sync();
                    targetRange = selectedRange;
                } catch (e) {
                    // Default to a reasonable data range
                    targetRange = worksheet.getRange('A1:Z50');
                }

                targetRange.load(['address']);
                await context.sync();

                const lowerMessage = message.toLowerCase();
                let conditionalFormat;
                let formatDescription;

                // Create different conditional formats based on message content
                if (lowerMessage.includes('negative') || lowerMessage.includes('less than') || lowerMessage.includes('< 0')) {
                    // Highlight negative numbers in red (following Excel Add-in API example)
                    conditionalFormat = targetRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                    conditionalFormat.cellValue.format.font.color = "red";
                    conditionalFormat.cellValue.rule = { formula1: "0", operator: "LessThan" };
                    formatDescription = "negative values in red";
                    
                } else if (lowerMessage.includes('greater than') || lowerMessage.includes('> 0')) {
                    // Highlight positive numbers in green
                    conditionalFormat = targetRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                    conditionalFormat.cellValue.format.font.color = "green";
                    conditionalFormat.cellValue.rule = { formula1: "0", operator: "GreaterThan" };
                    formatDescription = "positive values in green";
                    
                } else if (lowerMessage.includes('color scale') || lowerMessage.includes('gradient')) {
                    // Create color scale (following Excel Add-in API example)
                    conditionalFormat = targetRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
                    const criteria = {
                        minimum: {
                            formula: null,
                            type: Excel.ConditionalFormatColorCriterionType.lowestValue,
                            color: "blue"
                        },
                        midpoint: {
                            formula: "50",
                            type: Excel.ConditionalFormatColorCriterionType.percent,
                            color: "yellow"
                        },
                        maximum: {
                            formula: null,
                            type: Excel.ConditionalFormatColorCriterionType.highestValue,
                            color: "red"
                        }
                    };
                    conditionalFormat.colorScale.criteria = criteria;
                    formatDescription = "color scale from blue (low) to red (high)";
                    
                } else {
                    // Default: highlight non-empty cells
                    conditionalFormat = targetRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                    conditionalFormat.custom.rule.formula = '=LEN(TRIM(A1))>0';
                    conditionalFormat.custom.format.fill.color = this.extractColor(message);
                    formatDescription = `non-empty cells in ${this.getColorName(this.extractColor(message))}`;
                }

                await context.sync();

                return {
                    success: true,
                    message: `✅ Applied conditional formatting to ${targetRange.address}`,
                    action: 'conditional_format',
                    range: targetRange.address,
                    details: `Highlighted ${formatDescription}`
                };

            } catch (error) {
                console.error('Conditional formatting error:', error);
                return {
                    success: false,
                    message: `❌ Failed to create conditional formatting: ${error.message}`
                };
            }
        });
    }

    /**
     * Generic formatting action
     */
    async genericFormat(message) {
        // Default to color change if we can't determine specific action
        return await this.changeColor(message);
    }

    /**
     * Extract color from user message
     */
    extractColor(message) {
        const colorMap = {
            'green': '#22c55e',
            'red': '#ef4444',
            'blue': '#3b82f6',
            'yellow': '#fbbf24',
            'orange': '#f97316',
            'purple': '#8b5cf6',
            'pink': '#ec4899',
            'gray': '#6b7280',
            'grey': '#6b7280',
            'black': '#000000',
            'white': '#ffffff',
            'light green': '#86efac',
            'dark green': '#15803d',
            'light blue': '#93c5fd',
            'dark blue': '#1e40af'
        };

        const lowerMessage = message.toLowerCase();
        
        // Check for specific color mentions
        for (const [colorName, hexCode] of Object.entries(colorMap)) {
            if (lowerMessage.includes(colorName)) {
                return hexCode;
            }
        }

        // Default to green
        return '#22c55e';
    }

    /**
     * Get contrasting text color
     */
    getContrastingColor(backgroundColor) {
        const darkColors = ['#22c55e', '#ef4444', '#3b82f6', '#8b5cf6', '#ec4899', '#000000', '#15803d', '#1e40af'];
        return darkColors.includes(backgroundColor) ? '#ffffff' : '#000000';
    }

    /**
     * Get human-readable color name
     */
    getColorName(hexCode) {
        const colorNames = {
            '#22c55e': 'green',
            '#ef4444': 'red',
            '#3b82f6': 'blue',
            '#fbbf24': 'yellow',
            '#f97316': 'orange',
            '#8b5cf6': 'purple',
            '#ec4899': 'pink',
            '#6b7280': 'gray',
            '#000000': 'black',
            '#ffffff': 'white'
        };
        return colorNames[hexCode] || 'color';
    }

    /**
     * Create formatted response HTML
     */
    formatResponse(result) {
        if (!result.success) {
            return `<div class="action-error">${result.message}</div>
                   <div class="action-tip">💡 Try selecting specific cells first, then ask me to format them. For example: "Change the selected cells to green background"</div>`;
        }

        let html = `<div class="action-success">${result.message}</div>`;
        
        // Show operation details
        if (result.details) {
            html += `<div class="action-details" style="margin-top: 8px; padding: 8px; background: #f0f9ff; border-radius: 4px; font-size: 13px;">`;
            html += `<strong>🎨 Operation:</strong> ${result.details}`;
            html += `</div>`;
        }
        
        // Show affected range and cell count
        if (result.range) {
            html += `<div class="action-details" style="margin-top: 8px;">`;
            html += `<strong>📍 Range affected:</strong> <span class="cell-address">${result.range}</span>`;
            if (result.affectedCells !== undefined) {
                html += ` <span style="color: #6b7280;">(${result.affectedCells} cells with data)</span>`;
            }
            html += `</div>`;
        }

        // Show changes applied
        if (result.changes && result.changes.length > 0) {
            html += `<div class="action-details" style="margin-top: 8px;">`;
            html += `<h4 style="margin: 0 0 4px 0; font-size: 13px;">📋 Changes Applied:</h4>`;
            html += `<ul style="margin: 0; padding-left: 16px; font-size: 12px;">`;
            result.changes.forEach(change => {
                html += `<li>${change}</li>`;
            });
            html += `</ul>`;
            html += `</div>`;
        }

        // Add helpful tip for future actions
        if (result.action === 'color_change') {
            html += `<div class="action-tip" style="margin-top: 8px; padding: 6px; background: #fffbeb; border-left: 3px solid #f59e0b; border-radius: 4px; font-size: 12px; color: #92400e;">`;
            html += `💡 <strong>Pro tip:</strong> Select specific cells or ranges before asking for formatting to target exactly what you want to change.`;
            html += `</div>`;
        }

        return html;
    }
}

// Initialize and make globally available
window.directExcelActions = new DirectExcelActions();

console.log('🚀 Direct Excel Actions loaded');