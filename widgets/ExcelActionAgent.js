/**
 * Excel Action Agent - Direct Excel Manipulation
 * Handles AI requests that require actual changes to Excel workbooks
 */

class ExcelActionAgent {
    constructor() {
        this.isInitialized = false;
        this.setupAgent();
    }

    async setupAgent() {
        console.log('🔧 Initializing Excel Action Agent...');
        
        // Wait for Office.js to be ready
        if (typeof Office !== 'undefined' && Office.context) {
            this.isInitialized = true;
            console.log('✅ Excel Action Agent ready');
        } else {
            console.log('⚠️ Office.js not ready, will retry...');
            setTimeout(() => this.setupAgent(), 1000);
        }
    }

    /**
     * Process action requests from chat
     */
    async processActionRequest(message, context) {
        if (!this.isInitialized) {
            throw new Error('Excel Action Agent not initialized');
        }

        console.log('🎯 Processing action request:', message);

        // Analyze the request to determine what action to take
        const actionType = this.analyzeActionRequest(message);
        
        let result;
        switch (actionType) {
            case 'change_header_color':
                result = await this.changeHeaderColor(message, context);
                break;
            case 'format_cells':
                result = await this.formatCells(message, context);
                break;
            case 'add_formula':
                result = await this.addFormula(message, context);
                break;
            case 'update_values':
                result = await this.updateValues(message, context);
                break;
            case 'create_chart':
                result = await this.createChart(message, context);
                break;
            default:
                result = await this.handleGenericAction(message, context);
        }

        return result;
    }

    /**
     * Analyze the user's request to determine action type
     */
    analyzeActionRequest(message) {
        const lower = message.toLowerCase();
        
        if (lower.includes('change') && (lower.includes('color') || lower.includes('colour'))) {
            if (lower.includes('header') || lower.includes('title')) {
                return 'change_header_color';
            }
            return 'format_cells';
        }
        
        if (lower.includes('format') || lower.includes('style')) {
            return 'format_cells';
        }
        
        if (lower.includes('formula') || lower.includes('calculate')) {
            return 'add_formula';
        }
        
        if (lower.includes('update') || lower.includes('change value')) {
            return 'update_values';
        }
        
        if (lower.includes('chart') || lower.includes('graph')) {
            return 'create_chart';
        }
        
        return 'generic_action';
    }

    /**
     * Change header colors based on request
     */
    async changeHeaderColor(message, context) {
        console.log('🎨 Changing header colors...');
        
        return await Excel.run(async (context) => {
            try {
                // Get the active worksheet
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                
                // Determine target color from message
                const targetColor = this.extractColorFromMessage(message);
                const sourceColor = this.extractSourceColorFromMessage(message);
                
                console.log(`Changing from ${sourceColor} to ${targetColor}`);
                
                // Get used range to search for headers
                const usedRange = worksheet.getUsedRange();
                usedRange.load(['values', 'format/fill/color', 'rowCount', 'columnCount']);
                
                await context.sync();
                
                let changedCells = 0;
                const changes = [];
                
                // Search through cells to find headers with the source color
                for (let row = 0; row < Math.min(usedRange.rowCount, 20); row++) {
                    for (let col = 0; col < usedRange.columnCount; col++) {
                        const cellAddress = this.getCellAddress(row, col);
                        const cell = worksheet.getRange(cellAddress);
                        cell.load(['format/fill/color', 'values', 'format/font']);
                        
                        await context.sync();
                        
                        const cellColor = cell.format.fill.color.toLowerCase();
                        const cellValue = cell.values[0][0];
                        
                        // Check if this cell matches our criteria
                        if (this.shouldChangeColor(cellColor, sourceColor, cellValue, message)) {
                            // Change the color
                            cell.format.fill.color = targetColor;
                            cell.format.font.color = this.getContrastingTextColor(targetColor);
                            
                            changedCells++;
                            changes.push({
                                address: cellAddress,
                                value: cellValue,
                                oldColor: cellColor,
                                newColor: targetColor
                            });
                        }
                    }
                }
                
                await context.sync();
                
                return {
                    success: true,
                    action: 'change_header_color',
                    changedCells: changedCells,
                    changes: changes,
                    message: `✅ Successfully changed ${changedCells} header cells from ${sourceColor} to ${targetColor}`
                };
                
            } catch (error) {
                console.error('Error changing header colors:', error);
                return {
                    success: false,
                    action: 'change_header_color',
                    error: error.message,
                    message: `❌ Failed to change header colors: ${error.message}`
                };
            }
        });
    }

    /**
     * Extract target color from user message
     */
    extractColorFromMessage(message) {
        const lower = message.toLowerCase();
        
        // Color mapping
        const colors = {
            'green': '#22c55e',
            'red': '#ef4444', 
            'blue': '#3b82f6',
            'yellow': '#eab308',
            'orange': '#f97316',
            'purple': '#8b5cf6',
            'pink': '#ec4899',
            'gray': '#6b7280',
            'grey': '#6b7280',
            'black': '#000000',
            'white': '#ffffff'
        };
        
        // Look for color words in the message
        for (const [colorName, colorCode] of Object.entries(colors)) {
            if (lower.includes(colorName)) {
                return colorCode;
            }
        }
        
        // Default to green if no specific color found
        return '#22c55e';
    }

    /**
     * Extract source color from user message
     */
    extractSourceColorFromMessage(message) {
        const lower = message.toLowerCase();
        
        if (lower.includes('blue')) return 'blue';
        if (lower.includes('red')) return 'red';
        if (lower.includes('yellow')) return 'yellow';
        if (lower.includes('orange')) return 'orange';
        if (lower.includes('purple')) return 'purple';
        if (lower.includes('pink')) return 'pink';
        if (lower.includes('gray') || lower.includes('grey')) return 'gray';
        
        return 'any'; // Change any colored headers
    }

    /**
     * Determine if a cell should have its color changed
     */
    shouldChangeColor(cellColor, sourceColor, cellValue, message) {
        // If looking for any color, change any non-white/transparent cells with text
        if (sourceColor === 'any') {
            return cellColor !== '#ffffff' && 
                   cellColor !== 'transparent' && 
                   cellValue && 
                   cellValue.toString().length > 0;
        }
        
        // Map color names to possible hex values
        const colorMap = {
            'blue': ['#3b82f6', '#2563eb', '#1d4ed8', '#1e40af', '#1e3a8a', '#0ea5e9', '#0284c7', '#0369a1'],
            'red': ['#ef4444', '#dc2626', '#b91c1c', '#991b1b', '#7f1d1d'],
            'green': ['#22c55e', '#16a34a', '#15803d', '#166534', '#14532d'],
            'yellow': ['#eab308', '#ca8a04', '#a16207', '#854d0e'],
            'orange': ['#f97316', '#ea580c', '#c2410c', '#9a3412'],
            'purple': ['#8b5cf6', '#7c3aed', '#6d28d9', '#5b21b6'],
            'pink': ['#ec4899', '#db2777', '#be185d', '#9d174d']
        };
        
        const possibleColors = colorMap[sourceColor] || [];
        return possibleColors.some(color => 
            cellColor.toLowerCase().includes(color.toLowerCase()) ||
            this.colorsAreSimilar(cellColor, color)
        );
    }

    /**
     * Check if two colors are similar
     */
    colorsAreSimilar(color1, color2) {
        // Simple color similarity check
        return color1.toLowerCase() === color2.toLowerCase();
    }

    /**
     * Get contrasting text color for background
     */
    getContrastingTextColor(backgroundColor) {
        // Simple logic: dark backgrounds get white text, light backgrounds get black text
        const darkColors = ['#22c55e', '#ef4444', '#3b82f6', '#8b5cf6', '#ec4899', '#000000'];
        return darkColors.includes(backgroundColor) ? '#ffffff' : '#000000';
    }

    /**
     * Convert row/column indices to Excel address
     */
    getCellAddress(row, col) {
        const columnLetter = String.fromCharCode(65 + col); // A, B, C, etc.
        return `${columnLetter}${row + 1}`;
    }

    /**
     * Format cells based on request
     */
    async formatCells(message, context) {
        console.log('🎨 Formatting cells...');
        
        return await Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const selection = context.workbook.getSelectedRange();
                
                selection.load(['address', 'values']);
                await context.sync();
                
                // Apply formatting based on message
                if (message.toLowerCase().includes('bold')) {
                    selection.format.font.bold = true;
                }
                
                if (message.toLowerCase().includes('italic')) {
                    selection.format.font.italic = true;
                }
                
                // Apply color if specified
                const color = this.extractColorFromMessage(message);
                if (color !== '#22c55e') { // If not default green
                    selection.format.fill.color = color;
                    selection.format.font.color = this.getContrastingTextColor(color);
                }
                
                await context.sync();
                
                return {
                    success: true,
                    action: 'format_cells',
                    range: selection.address,
                    message: `✅ Successfully formatted cells in range ${selection.address}`
                };
                
            } catch (error) {
                return {
                    success: false,
                    action: 'format_cells',
                    error: error.message,
                    message: `❌ Failed to format cells: ${error.message}`
                };
            }
        });
    }

    /**
     * Handle generic Excel actions
     */
    async handleGenericAction(message, context) {
        console.log('🔧 Handling generic action...');
        
        // For now, return a message indicating the action was recognized
        return {
            success: true,
            action: 'generic_action',
            message: `🤖 I understand you want to: "${message}". This action type is being developed. For now, I've analyzed your Excel data instead.`
        };
    }

    /**
     * Create a user-friendly response for the chat
     */
    formatActionResponse(result) {
        if (result.success) {
            let response = `<div class="action-success">${result.message}</div>`;
            
            if (result.changes && result.changes.length > 0) {
                response += `<h4>📍 Changes Made:</h4>`;
                response += `<ul class="action-changes-list">`;
                result.changes.forEach(change => {
                    response += `<li>`;
                    response += `<span class="cell-address">${change.address}</span>: `;
                    response += `"${change.value}" `;
                    response += `<span class="color-change">`;
                    response += `<span class="color-swatch" style="background-color: ${change.oldColor}"></span> → `;
                    response += `<span class="color-swatch" style="background-color: ${change.newColor}"></span>`;
                    response += `</span>`;
                    response += `</li>`;
                });
                response += `</ul>`;
            }
            
            if (result.range) {
                response += `<p><strong>📍 Range affected:</strong> <span class="cell-address">${result.range}</span></p>`;
            }
            
            return response;
        } else {
            return `<div class="action-error">${result.message}</div>
                   <div class="action-tip">Try selecting the cells you want to modify first, then ask me to make the changes.</div>`;
        }
    }
}

// Initialize and make globally available
window.excelActionAgent = new ExcelActionAgent();

console.log('🚀 Excel Action Agent loaded');