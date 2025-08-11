/**
 * Safe Excel Context Reader - Handles Excel API calls with proper error handling
 * Fixes PropertyNotLoaded errors and provides reliable context data
 */

class SafeExcelContext {
    constructor() {
        this.isAvailable = false;
        this.lastGoodContext = null;
        this.initialize();
    }

    async initialize() {
        try {
            if (typeof Excel !== 'undefined' && Office && Office.context) {
                this.isAvailable = true;
                console.log('✅ Safe Excel Context ready');
            } else {
                console.log('⏳ Waiting for Excel API...');
                setTimeout(() => this.initialize(), 500);
            }
        } catch (error) {
            console.error('❌ Safe Excel Context initialization failed:', error);
        }
    }

    /**
     * Get basic context safely with minimal API calls
     */
    async getBasicContext() {
        if (!this.isAvailable) {
            return { 
                error: 'Excel API not available',
                lastGoodContext: this.lastGoodContext 
            };
        }

        try {
            return await Excel.run(async (context) => {
                const workbook = context.workbook;
                const activeWorksheet = workbook.worksheets.getActiveWorksheet();
                
                // Load only essential properties safely
                activeWorksheet.load(['name']);
                await context.sync();

                const worksheetName = activeWorksheet.name;
                
                // Try to get used range safely
                let usedRangeInfo = null;
                try {
                    const usedRange = activeWorksheet.getUsedRangeOrNullObject();
                    usedRange.load(['address', 'rowCount', 'columnCount']);
                    await context.sync();
                    
                    if (!usedRange.isNullObject) {
                        usedRangeInfo = {
                            address: usedRange.address,
                            rowCount: usedRange.rowCount,
                            columnCount: usedRange.columnCount
                        };
                    }
                } catch (rangeError) {
                    console.warn('Could not get used range:', rangeError);
                }

                // Try to get selected range safely
                let selectedRangeInfo = null;
                try {
                    const selectedRange = workbook.getSelectedRange();
                    selectedRange.load(['address']);
                    await context.sync();
                    
                    selectedRangeInfo = {
                        address: selectedRange.address
                    };
                } catch (selectionError) {
                    console.warn('Could not get selected range:', selectionError);
                }

                const basicContext = {
                    worksheetName: worksheetName,
                    usedRange: usedRangeInfo,
                    selectedRange: selectedRangeInfo,
                    timestamp: new Date().toISOString(),
                    hasData: usedRangeInfo !== null
                };

                this.lastGoodContext = basicContext;
                return basicContext;
            });
        } catch (error) {
            console.error('Error getting basic Excel context:', error);
            return { 
                error: error.message,
                lastGoodContext: this.lastGoodContext 
            };
        }
    }

    /**
     * Get sample data from the active sheet (first 5x5 cells)
     */
    async getSampleData() {
        if (!this.isAvailable) {
            return { error: 'Excel API not available' };
        }

        try {
            return await Excel.run(async (context) => {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                
                // Get a small sample range (A1:E5)
                const sampleRange = worksheet.getRange('A1:E5');
                sampleRange.load(['values']);
                await context.sync();

                const sampleData = [];
                for (let row = 0; row < sampleRange.values.length; row++) {
                    const rowData = [];
                    for (let col = 0; col < sampleRange.values[row].length; col++) {
                        const value = sampleRange.values[row][col];
                        if (value !== null && value !== undefined && value !== '') {
                            rowData.push({
                                address: this.getCellAddress(row, col),
                                value: value,
                                type: typeof value
                            });
                        }
                    }
                    if (rowData.length > 0) {
                        sampleData.push(rowData);
                    }
                }

                return {
                    sampleData: sampleData,
                    totalNonEmptyCells: sampleData.reduce((sum, row) => sum + row.length, 0),
                    timestamp: new Date().toISOString()
                };
            });
        } catch (error) {
            console.error('Error getting sample data:', error);
            return { error: error.message };
        }
    }

    /**
     * Check if a specific range has data
     */
    async checkRangeHasData(rangeAddress) {
        if (!this.isAvailable) {
            return false;
        }

        try {
            return await Excel.run(async (context) => {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const range = worksheet.getRange(rangeAddress);
                
                range.load(['values']);
                await context.sync();

                // Check if any cell in the range has data
                for (let row = 0; row < range.values.length; row++) {
                    for (let col = 0; col < range.values[row].length; col++) {
                        const value = range.values[row][col];
                        if (value !== null && value !== undefined && value !== '') {
                            return true;
                        }
                    }
                }
                return false;
            });
        } catch (error) {
            console.error('Error checking range data:', error);
            return false;
        }
    }

    /**
     * Get comprehensive context with fallbacks
     */
    async getComprehensiveContext() {
        const basicContext = await this.getBasicContext();
        if (basicContext.error) {
            return basicContext;
        }

        const sampleData = await this.getSampleData();
        
        return {
            ...basicContext,
            sampleData: sampleData.sampleData || [],
            totalNonEmptyCells: sampleData.totalNonEmptyCells || 0,
            hasActualData: (sampleData.totalNonEmptyCells || 0) > 0,
            contextType: 'comprehensive'
        };
    }

    /**
     * Convert row/col indices to Excel address (A1, B2, etc.)
     */
    getCellAddress(row, col) {
        let columnName = '';
        let tempCol = col;
        
        while (tempCol >= 0) {
            columnName = String.fromCharCode(65 + (tempCol % 26)) + columnName;
            tempCol = Math.floor(tempCol / 26) - 1;
        }
        
        return columnName + (row + 1);
    }

    /**
     * Generate summary for AI consumption
     */
    generateContextSummary(context) {
        if (context.error) {
            return `Excel context unavailable: ${context.error}`;
        }

        let summary = `Active worksheet: "${context.worksheetName}"`;
        
        if (context.usedRange) {
            summary += `\nData range: ${context.usedRange.address} (${context.usedRange.rowCount} rows × ${context.usedRange.columnCount} columns)`;
        }

        if (context.selectedRange) {
            summary += `\nSelected: ${context.selectedRange.address}`;
        }

        if (context.hasActualData && context.sampleData) {
            summary += `\nSample data preview:`;
            context.sampleData.slice(0, 3).forEach((row, rowIndex) => {
                const rowValues = row.map(cell => `${cell.address}: ${cell.value}`).join(', ');
                summary += `\n  Row ${rowIndex + 1}: ${rowValues}`;
            });
            
            if (context.totalNonEmptyCells > 0) {
                summary += `\nTotal data cells found: ${context.totalNonEmptyCells}`;
            }
        } else {
            summary += '\nNo data detected in preview area';
        }

        return summary;
    }

    /**
     * Get context optimized for chat responses
     */
    async getChatContext() {
        const context = await this.getComprehensiveContext();
        const summary = this.generateContextSummary(context);
        
        return {
            ...context,
            summary: summary,
            isUsable: !context.error && (context.hasActualData || context.hasData)
        };
    }

    /**
     * Reset and refresh context
     */
    async refresh() {
        this.lastGoodContext = null;
        return await this.getChatContext();
    }
}

// Initialize and make globally available
window.safeExcelContext = new SafeExcelContext();

console.log('🛡️ Safe Excel Context loaded');