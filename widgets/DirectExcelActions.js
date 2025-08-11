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
            /change.*color/i,
            /make.*bold/i,
            /format.*cell/i,
            /highlight/i,
            /set.*background/i,
            /change.*font/i,
            /make.*green/i,
            /make.*red/i,
            /make.*blue/i,
            /color.*header/i,
            /bold.*header/i,
            /format.*header/i
        ];

        return actionPatterns.some(pattern => pattern.test(message));
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
            // Determine action type
            if (message.toLowerCase().includes('color') || message.toLowerCase().includes('background')) {
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
     * Change cell colors
     */
    async changeColor(message) {
        return await Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                
                // Get selected range or use a default range
                let targetRange;
                try {
                    targetRange = context.workbook.getSelectedRange();
                    await context.sync();
                } catch (e) {
                    // If no selection, use first few header rows
                    targetRange = worksheet.getRange('A1:Z5');
                }

                // Load range properties safely
                targetRange.load(['address', 'values']);
                await context.sync();

                // Extract color from message
                const targetColor = this.extractColor(message);
                console.log('🎨 Applying color:', targetColor);

                // Apply formatting
                targetRange.format.fill.color = targetColor;
                targetRange.format.font.color = this.getContrastingColor(targetColor);
                
                await context.sync();

                return {
                    success: true,
                    message: `✅ Changed cell colors to ${this.getColorName(targetColor)} in range ${targetRange.address}`,
                    action: 'color_change',
                    range: targetRange.address,
                    color: targetColor
                };

            } catch (error) {
                console.error('Color change error:', error);
                return {
                    success: false,
                    message: `❌ Failed to change colors: ${error.message}`
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
            return `<div class="action-error">${result.message}</div>`;
        }

        let html = `<div class="action-success">${result.message}</div>`;
        
        if (result.changes && result.changes.length > 0) {
            html += `<div class="action-details">`;
            html += `<h4>📍 Changes Applied:</h4>`;
            html += `<ul>`;
            result.changes.forEach(change => {
                html += `<li>${change}</li>`;
            });
            html += `</ul>`;
            html += `</div>`;
        }

        if (result.range) {
            html += `<div class="action-details">`;
            html += `<strong>📍 Range affected:</strong> <span class="cell-address">${result.range}</span>`;
            html += `</div>`;
        }

        return html;
    }
}

// Initialize and make globally available
window.directExcelActions = new DirectExcelActions();

console.log('🚀 Direct Excel Actions loaded');