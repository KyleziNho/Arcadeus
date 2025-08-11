/**
 * Excel Navigator - Navigate to specific cells from chat responses
 * Professional M&A tool functionality
 */

class ExcelNavigator {
  constructor() {
    this.isNavigating = false;
  }

  /**
   * Navigate to a specific Excel cell or range
   */
  async navigateToCell(cellReference) {
    if (this.isNavigating) {
      console.log('Navigation already in progress...');
      return;
    }

    console.log(`🎯 Navigating to Excel cell: ${cellReference}`);
    this.isNavigating = true;

    try {
      if (typeof Excel === 'undefined') {
        throw new Error('Excel API not available');
      }

      return await Excel.run(async (context) => {
        // Parse cell reference (e.g., "FCF!B18" or "B18")
        const { worksheetName, cellAddress } = this.parseCellReference(cellReference);
        
        let worksheet;
        
        if (worksheetName) {
          // Navigate to specific worksheet
          try {
            worksheet = context.workbook.worksheets.getItem(worksheetName);
          } catch (error) {
            // If exact name doesn't exist, try to find similar worksheet
            const allWorksheets = context.workbook.worksheets;
            allWorksheets.load('items/name');
            await context.sync();
            
            const matchingSheet = this.findMatchingWorksheet(worksheetName, allWorksheets.items);
            if (matchingSheet) {
              worksheet = context.workbook.worksheets.getItem(matchingSheet.name);
              console.log(`📄 Using similar worksheet: ${matchingSheet.name}`);
            } else {
              throw new Error(`Worksheet "${worksheetName}" not found`);
            }
          }
        } else {
          // Use active worksheet
          worksheet = context.workbook.worksheets.getActiveWorksheet();
        }

        // Get the target range
        const targetRange = worksheet.getRange(cellAddress);
        
        // Select and navigate to the range
        targetRange.select();
        
        // Load range properties for feedback
        targetRange.load(['address', 'values', 'formulas']);
        worksheet.load('name');
        
        await context.sync();
        
        // Provide user feedback
        const actualValue = targetRange.values[0][0];
        const actualFormula = targetRange.formulas[0][0];
        
        this.showNavigationFeedback({
          cellReference,
          actualAddress: targetRange.address,
          worksheetName: worksheet.name,
          value: actualValue,
          formula: actualFormula
        });
        
        console.log(`✅ Successfully navigated to ${targetRange.address} on ${worksheet.name}`);
        
        return {
          success: true,
          address: targetRange.address,
          worksheet: worksheet.name,
          value: actualValue,
          formula: actualFormula
        };
      });
      
    } catch (error) {
      console.error('❌ Navigation failed:', error);
      this.showNavigationError(cellReference, error.message);
      
      return {
        success: false,
        error: error.message
      };
    } finally {
      this.isNavigating = false;
    }
  }

  /**
   * Parse cell reference into worksheet and cell parts
   */
  parseCellReference(reference) {
    // Handle formats like "FCF!B18", "Dashboard!A1", or just "B18"
    const match = reference.match(/^(?:([^!]+)!)?([A-Z]+\d+(?::[A-Z]+\d+)?)$/i);
    
    if (!match) {
      throw new Error(`Invalid cell reference format: ${reference}`);
    }
    
    return {
      worksheetName: match[1] || null,
      cellAddress: match[2]
    };
  }

  /**
   * Find a worksheet with similar name (fuzzy matching)
   */
  findMatchingWorksheet(targetName, worksheets) {
    const target = targetName.toLowerCase();
    
    // Try exact match first
    let match = worksheets.find(ws => ws.name.toLowerCase() === target);
    if (match) return match;
    
    // Try partial match
    match = worksheets.find(ws => ws.name.toLowerCase().includes(target) || target.includes(ws.name.toLowerCase()));
    if (match) return match;
    
    // Try common abbreviations
    const abbreviations = {
      'fcf': ['free cash flow', 'cash flow', 'cf'],
      'pl': ['profit loss', 'p&l', 'income'],
      'bs': ['balance sheet', 'balance'],
      'assumptions': ['inputs', 'params', 'parameters'],
      'dashboard': ['summary', 'overview']
    };
    
    for (const [abbrev, fullNames] of Object.entries(abbreviations)) {
      if (target === abbrev) {
        match = worksheets.find(ws => 
          fullNames.some(name => ws.name.toLowerCase().includes(name))
        );
        if (match) return match;
      }
    }
    
    return null;
  }

  /**
   * Show navigation feedback to user
   */
  showNavigationFeedback(info) {
    // Create a temporary notification
    const notification = document.createElement('div');
    notification.className = 'excel-navigation-feedback';
    notification.innerHTML = `
      <div class="nav-feedback-content">
        <div class="nav-feedback-header">
          <span class="nav-feedback-icon">🎯</span>
          <span class="nav-feedback-title">Navigated to Excel</span>
        </div>
        <div class="nav-feedback-details">
          <div><strong>Cell:</strong> ${info.actualAddress}</div>
          <div><strong>Sheet:</strong> ${info.worksheetName}</div>
          ${info.value !== undefined ? `<div><strong>Value:</strong> ${info.value}</div>` : ''}
          ${info.formula && info.formula !== info.value ? `<div><strong>Formula:</strong> ${info.formula}</div>` : ''}
        </div>
      </div>
    `;

    // Add to chat area
    const chatContainer = document.getElementById('chatMessages') || document.body;
    chatContainer.appendChild(notification);

    // Auto-remove after 5 seconds
    setTimeout(() => {
      notification.remove();
    }, 5000);
  }

  /**
   * Show navigation error to user
   */
  showNavigationError(cellReference, errorMessage) {
    const notification = document.createElement('div');
    notification.className = 'excel-navigation-error';
    notification.innerHTML = `
      <div class="nav-error-content">
        <div class="nav-error-header">
          <span class="nav-error-icon">❌</span>
          <span class="nav-error-title">Navigation Failed</span>
        </div>
        <div class="nav-error-details">
          <div><strong>Target:</strong> ${cellReference}</div>
          <div><strong>Error:</strong> ${errorMessage}</div>
        </div>
      </div>
    `;

    const chatContainer = document.getElementById('chatMessages') || document.body;
    chatContainer.appendChild(notification);

    setTimeout(() => {
      notification.remove();
    }, 5000);
  }

  /**
   * Get cell value for tooltip preview
   */
  async getCellPreview(cellReference) {
    try {
      if (typeof Excel === 'undefined') {
        return { error: 'Excel not available' };
      }

      return await Excel.run(async (context) => {
        const { worksheetName, cellAddress } = this.parseCellReference(cellReference);
        
        let worksheet;
        if (worksheetName) {
          worksheet = context.workbook.worksheets.getItem(worksheetName);
        } else {
          worksheet = context.workbook.worksheets.getActiveWorksheet();
        }

        const range = worksheet.getRange(cellAddress);
        range.load(['values', 'formulas', 'numberFormat']);
        worksheet.load('name');
        
        await context.sync();

        return {
          worksheet: worksheet.name,
          address: cellReference,
          value: range.values[0][0],
          formula: range.formulas[0][0],
          format: range.numberFormat[0][0]
        };
      });
    } catch (error) {
      return { error: error.message };
    }
  }

  /**
   * Inject navigation styles
   */
  injectStyles() {
    if (document.getElementById('excel-navigation-styles')) return;

    const style = document.createElement('style');
    style.id = 'excel-navigation-styles';
    style.textContent = `
      /* Clickable cell references */
      .cell-highlight {
        cursor: pointer !important;
        transition: all 0.2s ease !important;
        position: relative !important;
        text-decoration: none !important;
      }

      .cell-highlight:hover {
        background: #065f46 !important;
        color: white !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
      }

      .cell-highlight:active {
        transform: translateY(0px) !important;
      }

      /* Navigation feedback */
      .excel-navigation-feedback {
        position: fixed;
        top: 20px;
        right: 20px;
        background: #10b981;
        color: white;
        padding: 12px 16px;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 10000;
        animation: slideInRight 0.3s ease-out;
        max-width: 300px;
      }

      .excel-navigation-error {
        position: fixed;
        top: 20px;
        right: 20px;
        background: #dc2626;
        color: white;
        padding: 12px 16px;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 10000;
        animation: slideInRight 0.3s ease-out;
        max-width: 300px;
      }

      .nav-feedback-header,
      .nav-error-header {
        display: flex;
        align-items: center;
        margin-bottom: 8px;
        font-weight: 600;
      }

      .nav-feedback-icon,
      .nav-error-icon {
        margin-right: 8px;
        font-size: 16px;
      }

      .nav-feedback-details,
      .nav-error-details {
        font-size: 13px;
        opacity: 0.9;
      }

      .nav-feedback-details div,
      .nav-error-details div {
        margin: 2px 0;
      }

      /* Cell preview tooltip */
      .cell-preview-tooltip {
        position: absolute;
        background: #1f2937;
        color: white;
        padding: 8px 12px;
        border-radius: 6px;
        font-size: 12px;
        white-space: nowrap;
        z-index: 1000;
        pointer-events: none;
        box-shadow: 0 2px 8px rgba(0,0,0,0.2);
        top: -40px;
        left: 50%;
        transform: translateX(-50%);
      }

      .cell-preview-tooltip:before {
        content: '';
        position: absolute;
        top: 100%;
        left: 50%;
        transform: translateX(-50%);
        border: 5px solid transparent;
        border-top-color: #1f2937;
      }

      @keyframes slideInRight {
        from {
          transform: translateX(100%);
          opacity: 0;
        }
        to {
          transform: translateX(0);
          opacity: 1;
        }
      }
    `;

    document.head.appendChild(style);
  }
}

// Export for global use
window.ExcelNavigator = ExcelNavigator;

// Initialize global instance
window.excelNavigator = new ExcelNavigator();
window.excelNavigator.injectStyles();

console.log('🧭 Excel Navigator loaded and ready');