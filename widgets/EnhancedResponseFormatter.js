/**
 * Enhanced Response Formatter - Makes IRR references and values clickable
 * Implements Excel green highlighting and cell navigation
 */

class EnhancedResponseFormatter {
    constructor() {
        this.cellNavigator = window.excelNavigator;
    }

    /**
     * Format response text to make IRR references and values clickable
     */
    formatResponse(text) {
        if (!text || typeof text !== 'string') return text;

        let formattedText = text;

        // Step 1: Make IRR references clickable
        formattedText = this.makeIRRReferencesClickable(formattedText);

        // Step 2: Make percentage values with cell references clickable
        formattedText = this.makePercentageValuesClickable(formattedText);

        // Step 3: Make cell references clickable
        formattedText = this.makeCellReferencesClickable(formattedText);

        // Step 4: Make other metric references clickable
        formattedText = this.makeMetricReferencesClickable(formattedText);

        // Step 5: Apply non-clickable highlights for emphasis
        formattedText = this.applyNonClickableHighlights(formattedText);

        return formattedText;
    }

    /**
     * Make IRR references clickable (e.g., "Unlevered IRR (B21)")
     */
    makeIRRReferencesClickable(text) {
        // Pattern: "Unlevered IRR (B21)" or "IRR (C15)" etc.
        const irrPattern = /([A-Za-z\s]*IRR[A-Za-z\s]*)\s*\(([A-Z]+\d+)\)/gi;
        
        return text.replace(irrPattern, (match, irrType, cellRef) => {
            const cleanIRRType = irrType.trim();
            return `<span class="irr-reference-clickable" onclick="navigateToCell('${cellRef}')" title="Click to navigate to ${cellRef}">${cleanIRRType} (${cellRef})</span>`;
        });
    }

    /**
     * Make percentage values with implicit cell references clickable
     */
    makePercentageValuesClickable(text) {
        // Pattern: "20.1%" or similar percentages that might reference cells
        const percentagePattern = /(\d+\.?\d*)%/g;
        
        return text.replace(percentagePattern, (match, number) => {
            // Try to find associated cell reference in context
            const cellRef = this.findAssociatedCellReference(text, match);
            
            if (cellRef) {
                return `<span class="value-highlight" onclick="navigateToCell('${cellRef}')" title="Click to navigate to ${cellRef}">${match}</span>`;
            } else {
                // If no cell reference found, make it a non-clickable highlight
                return `<span class="non-clickable-highlight">${match}</span>`;
            }
        });
    }

    /**
     * Make standalone cell references clickable
     */
    makeCellReferencesClickable(text) {
        // Pattern: B21, C15, etc. but avoid those already in spans
        const cellPattern = /(?<![\w\-"])([A-Z]+\d+)(?![\w\-"])/g;
        
        return text.replace(cellPattern, (match, cellRef) => {
            // Skip if already inside a span or other HTML element
            if (this.isInsideHTMLElement(text, match)) {
                return match;
            }
            
            return `<span class="cell-reference-clickable" onclick="navigateToCell('${cellRef}')" title="Navigate to ${cellRef}">${cellRef}</span>`;
        });
    }

    /**
     * Make other metric references clickable (MOIC, NPV, etc.)
     */
    makeMetricReferencesClickable(text) {
        const metricsPattern = /(MOIC|NPV|EBITDA|Revenue|Multiple)\s*[:\s]*([A-Z]+\d+)/gi;
        
        return text.replace(metricsPattern, (match, metricName, cellRef) => {
            return `<span class="metric-reference-clickable" onclick="navigateToCell('${cellRef}')" title="Click to navigate to ${cellRef}">${metricName}: ${cellRef}</span>`;
        });
    }

    /**
     * Apply non-clickable highlights for emphasis
     */
    applyNonClickableHighlights(text) {
        // Highlight important financial terms that don't have cell references
        const importantTerms = /\b(EBITDA|Revenue|Net Income|Cash Flow|Assets|Liabilities)\b(?!\s*\([A-Z]+\d+\))/gi;
        
        return text.replace(importantTerms, (match) => {
            // Skip if already inside a span
            if (this.isInsideHTMLElement(text, match)) {
                return match;
            }
            return `<span class="non-clickable-highlight">${match}</span>`;
        });
    }

    /**
     * Try to find associated cell reference for a value in the text context
     */
    findAssociatedCellReference(text, value) {
        // Look for patterns like "20.1% (B21)" or "20.1% in B21"
        const contextPatterns = [
            new RegExp(`${this.escapeRegex(value)}\\s*\\(([A-Z]+\\d+)\\)`, 'i'),
            new RegExp(`${this.escapeRegex(value)}\\s+(?:in|at|from)\\s+([A-Z]+\\d+)`, 'i'),
            new RegExp(`([A-Z]+\\d+)[:\\s]+${this.escapeRegex(value)}`, 'i')
        ];

        for (const pattern of contextPatterns) {
            const match = text.match(pattern);
            if (match) {
                return match[1];
            }
        }

        return null;
    }

    /**
     * Check if a text match is already inside an HTML element
     */
    isInsideHTMLElement(fullText, match) {
        const index = fullText.indexOf(match);
        if (index === -1) return false;

        const beforeMatch = fullText.substring(0, index);

        // Check if we're inside any HTML element
        const openTagCount = (beforeMatch.match(/<[^/][^>]*>/g) || []).length;
        const closeTagCount = (beforeMatch.match(/<\/[^>]*>/g) || []).length;

        // Also specifically check for highlight classes to prevent nested highlighting
        const highlightClassPattern = /<span[^>]*class="[^"]*(?:highlight|clickable)[^"]*"/gi;
        let lastHighlightSpan = -1;
        let match2;
        
        while ((match2 = highlightClassPattern.exec(beforeMatch)) !== null) {
            lastHighlightSpan = match2.index;
        }
        
        if (lastHighlightSpan >= 0) {
            // Check if there's a closing span after the last highlight span
            const afterHighlightSpan = beforeMatch.substring(lastHighlightSpan);
            const closingSpans = (afterHighlightSpan.match(/<\/span>/g) || []).length;
            const openingSpans = (afterHighlightSpan.match(/<span[^>]*>/g) || []).length;
            
            if (openingSpans > closingSpans) {
                return true; // We're inside a highlight span
            }
        }

        return openTagCount > closeTagCount;
    }

    /**
     * Escape special regex characters
     */
    escapeRegex(string) {
        return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    /**
     * Format a complete analysis response with proper highlighting
     */
    formatAnalysisResponse(analysisText) {
        // First apply the standard formatting
        let formatted = this.formatResponse(analysisText);

        // Add some structure for better readability
        formatted = formatted.replace(/\n\n/g, '</p><p>');
        formatted = formatted.replace(/\n/g, '<br>');
        formatted = `<p>${formatted}</p>`;

        // Clean up any empty paragraphs
        formatted = formatted.replace(/<p><\/p>/g, '');
        formatted = formatted.replace(/<p><br><\/p>/g, '');

        return formatted;
    }

    /**
     * Format specific metric callouts
     */
    formatMetricCallout(metricName, value, cellRef) {
        return `<div class="metric-callout">
            <span class="metric-name">${metricName}:</span>
            <span class="value-highlight" onclick="navigateToCell('${cellRef}')" title="Navigate to ${cellRef}">${value}</span>
            <span class="cell-location">(${cellRef})</span>
        </div>`;
    }
}

/**
 * Global function to navigate to Excel cells
 */
window.navigateToCell = function(cellAddress) {
    console.log(`🎯 Navigating to cell: ${cellAddress}`);
    
    if (window.excelNavigator) {
        // Use the existing excel navigator
        window.excelNavigator.navigateToCell(cellAddress);
    } else {
        console.warn('Excel navigator not available, using fallback...');
        // Fallback: try basic Excel navigation
        if (typeof Excel !== 'undefined') {
            Excel.run(async (context) => {
                try {
                    const worksheet = context.workbook.worksheets.getActiveWorksheet();
                    const range = worksheet.getRange(cellAddress);
                    range.select();
                    await context.sync();
                    console.log(`✅ Navigated to ${cellAddress} using fallback`);
                    
                    // Navigation successful - no popup needed
                    console.log(`✅ Successfully navigated to ${cellAddress}`);
                } catch (error) {
                    console.error(`❌ Failed to navigate to ${cellAddress}:`, error);
                    
                    // Navigation failed - log only, no popup
                    console.warn(`❌ Could not navigate to ${cellAddress}`);
                }
            });
        } else {
            console.error('Excel API not available for navigation');
        }
    }
};

// Initialize and make globally available
window.enhancedResponseFormatter = new EnhancedResponseFormatter();

console.log('🎨 Enhanced Response Formatter loaded');