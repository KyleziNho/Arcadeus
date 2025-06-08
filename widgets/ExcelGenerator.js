/* global Office, Excel */

// Cell Reference Tracker for dynamic formula generation
class CellReferenceTracker {
  constructor() {
    this.references = {
      assumptions: {},
      pnl: {},
      fcf: {}
    };
  }

  // Track where each assumption is placed
  trackAssumptionCell(itemName, cellAddress, worksheet = 'Assumptions') {
    if (!this.references[worksheet.toLowerCase()]) {
      this.references[worksheet.toLowerCase()] = {};
    }
    this.references[worksheet.toLowerCase()][itemName] = {
      cell: cellAddress,
      worksheet: worksheet
    };
  }

  // Get reference for formulas
  getReference(itemName, currentWorksheet = null) {
    // Check all worksheets for the reference
    for (const [ws, items] of Object.entries(this.references)) {
      if (items[itemName]) {
        const ref = items[itemName];
        // If we're on a different worksheet, include the worksheet name
        if (currentWorksheet && currentWorksheet.toLowerCase() !== ws) {
          return `${ref.worksheet}!${ref.cell}`;
        }
        return ref.cell;
      }
    }
    return null;
  }
}

class ExcelGenerator {
  constructor() {
    this.cellTracker = null;
  }

  async generateModel(modelData) {
    try {
      console.log('Starting model generation...');
      
      // Initialize cell tracker
      this.cellTracker = new CellReferenceTracker();

      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        
        // Create Assumptions sheet first
        console.log('Creating Assumptions sheet...');
        const assumptionsSheet = sheets.add('Assumptions');
        await this.createAssumptionsLayout(context, assumptionsSheet, modelData);
        
        await context.sync();
        console.log('Assumptions sheet created successfully');
      });

      // Wait 1 second before creating P&L sheet
      console.log('Waiting before creating P&L sheet...');
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      // Generate P&L sheet
      await this.generatePLSheet(modelData);
      
      // Wait 1 second before creating FCF sheet  
      console.log('Waiting before creating FCF sheet...');
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      // Generate FCF sheet
      await this.generateFCFSheet(modelData);

      console.log('Model generation completed successfully!');
      return { success: true, message: 'Model generated successfully!' };
      
    } catch (error) {
      console.error('Error generating model:', error);
      return { success: false, error: error.message };
    }
  }

  async generatePLSheet(modelData) {
    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        
        console.log('Creating P&L sheet...');
        const plSheet = sheets.add('P&L Statement');
        await this.createProfitLossLayout(context, plSheet, modelData);
        
        await context.sync();
        console.log('P&L sheet created successfully');
      });
    } catch (error) {
      console.error('Error generating P&L sheet:', error);
      throw error;
    }
  }

  async generateFCFSheet(modelData) {
    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        
        console.log('Creating FCF sheet...');
        const fcfSheet = sheets.add('Free Cash Flow');
        await this.createCashflowsLayout(context, fcfSheet, modelData);
        
        await context.sync();
        console.log('FCF sheet created successfully');
      });
    } catch (error) {
      console.error('Error generating FCF sheet:', error);
      throw error;
    }
  }

  async createAssumptionsLayout(context, sheet, data) {
    console.log('Creating assumptions layout...');
    
    // Set up headers
    const headerRange = sheet.getRange('A1:M1');
    headerRange.values = [['M&A Financial Model - Assumptions']];
    headerRange.format.font.bold = true;
    headerRange.format.font.size = 16;
    headerRange.format.fill.color = '#2D5A5A';
    headerRange.format.font.color = 'white';
    headerRange.merge();

    let currentRow = 3;

    // High-Level Parameters Section
    const hlParamHeader = sheet.getRange(`A${currentRow}:M${currentRow}`);
    hlParamHeader.values = [['High-Level Parameters']];
    hlParamHeader.format.font.bold = true;
    hlParamHeader.format.fill.color = '#4A90A4';
    hlParamHeader.format.font.color = 'white';
    hlParamHeader.merge();
    currentRow += 2;

    // Add high-level parameters
    const hlParams = [
      ['Currency', data.currency || 'USD'],
      ['Project Start Date', this.formatDateForExcel(data.projectStartDate)],
      ['Model Periods', data.modelPeriods || 'Monthly'],
      ['Project End Date', this.formatDateForExcel(data.projectEndDate)],
      ['Holding Periods', data.holdingPeriodsCalculated || '']
    ];

    for (const [label, value] of hlParams) {
      sheet.getRange(`A${currentRow}`).values = [[label]];
      sheet.getRange(`B${currentRow}`).values = [[value]];
      this.cellTracker.trackAssumptionCell(label.toLowerCase().replace(/\s+/g, '_'), `B${currentRow}`);
      currentRow++;
    }

    currentRow += 2;

    // Deal Assumptions Section
    const dealHeader = sheet.getRange(`A${currentRow}:M${currentRow}`);
    dealHeader.values = [['Deal Assumptions']];
    dealHeader.format.font.bold = true;
    dealHeader.format.fill.color = '#4A90A4';
    dealHeader.format.font.color = 'white';
    dealHeader.merge();
    currentRow += 2;

    // Add deal assumptions
    const dealParams = [
      ['Deal Name', data.dealName || ''],
      ['Deal Value', data.dealValue || 0],
      ['Transaction Fee (%)', data.transactionFee || 2.5],
      ['Deal LTV (%)', data.dealLTV || 70],
      ['Equity Contribution', data.equityContribution || ''],
      ['Debt Financing', data.debtFinancing || '']
    ];

    for (const [label, value] of dealParams) {
      sheet.getRange(`A${currentRow}`).values = [[label]];
      sheet.getRange(`B${currentRow}`).values = [[value]];
      this.cellTracker.trackAssumptionCell(label.toLowerCase().replace(/\s+/g, '_').replace(/[()%]/g, ''), `B${currentRow}`);
      currentRow++;
    }

    currentRow += 2;

    // Revenue Items Section
    if (data.revenueItems && data.revenueItems.length > 0) {
      const revenueHeader = sheet.getRange(`A${currentRow}:M${currentRow}`);
      revenueHeader.values = [['Revenue Items']];
      revenueHeader.format.font.bold = true;
      revenueHeader.format.fill.color = '#4A90A4';
      revenueHeader.format.font.color = 'white';
      revenueHeader.merge();
      currentRow += 2;

      // Create period headers for revenue
      const periods = this.calculatePeriods(data.projectStartDate, data.projectEndDate, data.modelPeriods);
      
      // Headers row for revenue items
      const revHeaders = ['Item', 'Base Value'];
      for (let i = 0; i < Math.min(periods, 10); i++) {
        revHeaders.push(this.formatDateHeader(new Date(data.projectStartDate), i, data.modelPeriods));
      }
      
      const revHeaderRange = sheet.getRange(`A${currentRow}:${String.fromCharCode(65 + revHeaders.length - 1)}${currentRow}`);
      revHeaderRange.values = [revHeaders];
      revHeaderRange.format.font.bold = true;
      revHeaderRange.format.fill.color = '#87CEEB';
      currentRow++;

      // Add revenue items
      data.revenueItems.forEach((item, index) => {
        sheet.getRange(`A${currentRow}`).values = [[item.name || `Revenue Item ${index + 1}`]];
        sheet.getRange(`B${currentRow}`).values = [[item.value || 0]];
        
        this.cellTracker.trackAssumptionCell(`revenue_item_${index}`, `B${currentRow}`);
        
        // Add growth rates if available
        if (item.growthType === 'periodic' && item.periods) {
          let col = 2; // Start from column C (index 2)
          item.periods.forEach(period => {
            if (col < revHeaders.length) {
              sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).values = [[period.value || 0]];
              col++;
            }
          });
        } else if (item.growthType === 'annual' && item.annualGrowthRate) {
          // Calculate values with annual growth
          let baseValue = item.value || 0;
          let col = 2;
          for (let i = 0; i < Math.min(periods, 10); i++) {
            const growthFactor = Math.pow(1 + (item.annualGrowthRate / 100), Math.floor(i / 12));
            const value = baseValue * growthFactor;
            sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).values = [[value]];
            col++;
          }
        }
        
        currentRow++;
      });
      
      currentRow += 2;
    }

    // Operating Expenses Section
    if (data.operatingExpenses && data.operatingExpenses.length > 0) {
      const opexHeader = sheet.getRange(`A${currentRow}:M${currentRow}`);
      opexHeader.values = [['Operating Expenses']];
      opexHeader.format.font.bold = true;
      opexHeader.format.fill.color = '#4A90A4';
      opexHeader.format.font.color = 'white';
      opexHeader.merge();
      currentRow += 2;

      data.operatingExpenses.forEach((item, index) => {
        sheet.getRange(`A${currentRow}`).values = [[item.name || `OpEx Item ${index + 1}`]];
        sheet.getRange(`B${currentRow}`).values = [[item.value || 0]];
        this.cellTracker.trackAssumptionCell(`opex_item_${index}`, `B${currentRow}`);
        currentRow++;
      });
      
      currentRow += 2;
    }

    // Capital Expenses Section
    if (data.capitalExpenses && data.capitalExpenses.length > 0) {
      const capexHeader = sheet.getRange(`A${currentRow}:M${currentRow}`);
      capexHeader.values = [['Capital Expenses']];
      capexHeader.format.font.bold = true;
      capexHeader.format.fill.color = '#4A90A4';
      capexHeader.format.font.color = 'white';
      capexHeader.merge();
      currentRow += 2;

      data.capitalExpenses.forEach((item, index) => {
        sheet.getRange(`A${currentRow}`).values = [[item.name || `CapEx Item ${index + 1}`]];
        sheet.getRange(`B${currentRow}`).values = [[item.value || 0]];
        this.cellTracker.trackAssumptionCell(`capex_item_${index}`, `B${currentRow}`);
        currentRow++;
      });
      
      currentRow += 2;
    }

    // Exit Assumptions Section
    const exitHeader = sheet.getRange(`A${currentRow}:M${currentRow}`);
    exitHeader.values = [['Exit Assumptions']];
    exitHeader.format.font.bold = true;
    exitHeader.format.fill.color = '#4A90A4';
    exitHeader.format.font.color = 'white';
    exitHeader.merge();
    currentRow += 2;

    const exitParams = [
      ['Disposal Cost (%)', data.disposalCost || 2.5],
      ['Terminal Cap Rate (%)', data.terminalCapRate || 8.5]
    ];

    for (const [label, value] of exitParams) {
      sheet.getRange(`A${currentRow}`).values = [[label]];
      sheet.getRange(`B${currentRow}`).values = [[value]];
      this.cellTracker.trackAssumptionCell(label.toLowerCase().replace(/\s+/g, '_').replace(/[()%]/g, ''), `B${currentRow}`);
      currentRow++;
    }

    // Auto-size columns
    sheet.getRange('A:M').format.autofitColumns();
    
    console.log('Assumptions layout created successfully');
  }

  async createProfitLossLayout(context, sheet, data) {
    console.log('Creating P&L layout...');
    
    // Set up headers
    const headerRange = sheet.getRange('A1:M1');
    headerRange.values = [['Profit & Loss Statement']];
    headerRange.format.font.bold = true;
    headerRange.format.font.size = 16;
    headerRange.format.fill.color = '#2D5A5A';
    headerRange.format.font.color = 'white';
    headerRange.merge();

    let currentRow = 3;

    // Calculate periods for headers
    const periods = this.calculatePeriods(data.projectStartDate, data.projectEndDate, data.modelPeriods);
    
    // Date headers
    const headers = ['Item'];
    for (let i = 0; i < Math.min(periods, 12); i++) {
      headers.push(this.formatDateHeader(new Date(data.projectStartDate), i, data.modelPeriods));
    }
    
    const headerRow = sheet.getRange(`A${currentRow}:${String.fromCharCode(65 + headers.length - 1)}${currentRow}`);
    headerRow.values = [headers];
    headerRow.format.font.bold = true;
    headerRow.format.fill.color = '#87CEEB';
    currentRow += 2;

    // Revenue Section
    const revenueHeader = sheet.getRange(`A${currentRow}:${String.fromCharCode(65 + headers.length - 1)}${currentRow}`);
    revenueHeader.values = [['REVENUE']];
    revenueHeader.format.font.bold = true;
    revenueHeader.format.fill.color = '#4A90A4';
    revenueHeader.format.font.color = 'white';
    revenueHeader.merge();
    currentRow++;

    // Add revenue items
    if (data.revenueItems && data.revenueItems.length > 0) {
      data.revenueItems.forEach((item, index) => {
        sheet.getRange(`A${currentRow}`).values = [[item.name || `Revenue Item ${index + 1}`]];
        
        // Add formulas referencing assumptions sheet
        for (let col = 1; col < headers.length; col++) {
          const assumptionRef = this.cellTracker.getReference(`revenue_item_${index}`, 'P&L Statement');
          if (assumptionRef) {
            sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).formulas = [[assumptionRef]];
          } else {
            sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).values = [[item.value || 0]];
          }
        }
        currentRow++;
      });
    }

    // Total Revenue row
    sheet.getRange(`A${currentRow}`).values = [['Total Revenue']];
    sheet.getRange(`A${currentRow}`).format.font.bold = true;
    for (let col = 1; col < headers.length; col++) {
      const startRow = currentRow - (data.revenueItems?.length || 0);
      const endRow = currentRow - 1;
      if (startRow <= endRow) {
        sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).formulas = [[`=SUM(${String.fromCharCode(65 + col)}${startRow}:${String.fromCharCode(65 + col)}${endRow})`]];
      }
    }
    currentRow += 2;

    // Operating Expenses Section
    const opexHeader = sheet.getRange(`A${currentRow}:${String.fromCharCode(65 + headers.length - 1)}${currentRow}`);
    opexHeader.values = [['OPERATING EXPENSES']];
    opexHeader.format.font.bold = true;
    opexHeader.format.fill.color = '#4A90A4';
    opexHeader.format.font.color = 'white';
    opexHeader.merge();
    currentRow++;

    // Add operating expenses
    if (data.operatingExpenses && data.operatingExpenses.length > 0) {
      data.operatingExpenses.forEach((item, index) => {
        sheet.getRange(`A${currentRow}`).values = [[item.name || `OpEx Item ${index + 1}`]];
        
        for (let col = 1; col < headers.length; col++) {
          const assumptionRef = this.cellTracker.getReference(`opex_item_${index}`, 'P&L Statement');
          if (assumptionRef) {
            sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).formulas = [[assumptionRef]];
          } else {
            sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).values = [[-(item.value || 0)]];
          }
        }
        currentRow++;
      });
    }

    // Total Operating Expenses row
    sheet.getRange(`A${currentRow}`).values = [['Total Operating Expenses']];
    sheet.getRange(`A${currentRow}`).format.font.bold = true;
    for (let col = 1; col < headers.length; col++) {
      const startRow = currentRow - (data.operatingExpenses?.length || 0);
      const endRow = currentRow - 1;
      if (startRow <= endRow) {
        sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).formulas = [[`=SUM(${String.fromCharCode(65 + col)}${startRow}:${String.fromCharCode(65 + col)}${endRow})`]];
      }
    }
    currentRow += 2;

    // EBITDA row
    sheet.getRange(`A${currentRow}`).values = [['EBITDA']];
    sheet.getRange(`A${currentRow}`).format.font.bold = true;
    const totalRevenueRow = currentRow - (data.operatingExpenses?.length || 0) - 4;
    const totalOpexRow = currentRow - 1;
    for (let col = 1; col < headers.length; col++) {
      sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).formulas = [[`=${String.fromCharCode(65 + col)}${totalRevenueRow}+${String.fromCharCode(65 + col)}${totalOpexRow}`]];
    }
    currentRow += 2;

    // Apply formatting for negative numbers
    const rangeToFormat = sheet.getRange(`B${3}:${String.fromCharCode(65 + headers.length - 1)}${currentRow}`);
    rangeToFormat.numberFormat = [['#,##0.00_);[Red](#,##0.00)']];

    // Auto-size columns
    sheet.getRange(`A:${String.fromCharCode(65 + headers.length - 1)}`).format.autofitColumns();
    
    console.log('P&L layout created successfully');
  }

  async createCashflowsLayout(context, sheet, modelData) {
    console.log('Creating cashflows layout...');
    
    // Set up headers
    const headerRange = sheet.getRange('A1:M1');
    headerRange.values = [['Free Cash Flow Statement']];
    headerRange.format.font.bold = true;
    headerRange.format.font.size = 16;
    headerRange.format.fill.color = '#2D5A5A';
    headerRange.format.font.color = 'white';
    headerRange.merge();

    let currentRow = 3;

    // Calculate periods for headers
    const periods = this.calculatePeriods(modelData.projectStartDate, modelData.projectEndDate, modelData.modelPeriods);
    
    // Date headers
    const headers = ['Item'];
    for (let i = 0; i < Math.min(periods, 12); i++) {
      headers.push(this.formatDateHeader(new Date(modelData.projectStartDate), i, modelData.modelPeriods));
    }
    
    const headerRow = sheet.getRange(`A${currentRow}:${String.fromCharCode(65 + headers.length - 1)}${currentRow}`);
    headerRow.values = [headers];
    headerRow.format.font.bold = true;
    headerRow.format.fill.color = '#87CEEB';
    currentRow += 2;

    // EBITDA (from P&L)
    sheet.getRange(`A${currentRow}`).values = [['EBITDA']];
    for (let col = 1; col < headers.length; col++) {
      sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).formulas = [[`='P&L Statement'!${String.fromCharCode(65 + col)}${currentRow}`]];
    }
    const ebitdaRow = currentRow;
    currentRow += 2;

    // Capital Expenditures Section
    const capexHeader = sheet.getRange(`A${currentRow}:${String.fromCharCode(65 + headers.length - 1)}${currentRow}`);
    capexHeader.values = [['CAPITAL EXPENDITURES']];
    capexHeader.format.font.bold = true;
    capexHeader.format.fill.color = '#4A90A4';
    capexHeader.format.font.color = 'white';
    capexHeader.merge();
    currentRow++;

    // Add capital expenses
    if (modelData.capitalExpenses && modelData.capitalExpenses.length > 0) {
      modelData.capitalExpenses.forEach((item, index) => {
        sheet.getRange(`A${currentRow}`).values = [[item.name || `CapEx Item ${index + 1}`]];
        
        for (let col = 1; col < headers.length; col++) {
          const assumptionRef = this.cellTracker.getReference(`capex_item_${index}`, 'Free Cash Flow');
          if (assumptionRef) {
            sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).formulas = [[`-${assumptionRef}`]];
          } else {
            sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).values = [[-(item.value || 0)]];
          }
        }
        currentRow++;
      });
    }

    // Total CapEx row
    sheet.getRange(`A${currentRow}`).values = [['Total Capital Expenditures']];
    sheet.getRange(`A${currentRow}`).format.font.bold = true;
    for (let col = 1; col < headers.length; col++) {
      const startRow = currentRow - (modelData.capitalExpenses?.length || 0);
      const endRow = currentRow - 1;
      if (startRow <= endRow) {
        sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).formulas = [[`=SUM(${String.fromCharCode(65 + col)}${startRow}:${String.fromCharCode(65 + col)}${endRow})`]];
      }
    }
    const totalCapexRow = currentRow;
    currentRow += 2;

    // Free Cash Flow row
    sheet.getRange(`A${currentRow}`).values = [['Free Cash Flow']];
    sheet.getRange(`A${currentRow}`).format.font.bold = true;
    for (let col = 1; col < headers.length; col++) {
      sheet.getRange(`${String.fromCharCode(65 + col)}${currentRow}`).formulas = [[`=${String.fromCharCode(65 + col)}${ebitdaRow}+${String.fromCharCode(65 + col)}${totalCapexRow}`]];
    }
    const fcfRow = currentRow;
    currentRow += 2;

    // Cumulative FCF row
    sheet.getRange(`A${currentRow}`).values = [['Cumulative Free Cash Flow']];
    sheet.getRange(`A${currentRow}`).format.font.bold = true;
    
    // First period is just the FCF
    sheet.getRange(`B${currentRow}`).formulas = [[`=B${fcfRow}`]];
    
    // Subsequent periods add previous cumulative + current FCF
    for (let col = 2; col < headers.length; col++) {
      const prevCol = String.fromCharCode(65 + col - 1);
      const currentCol = String.fromCharCode(65 + col);
      sheet.getRange(`${currentCol}${currentRow}`).formulas = [[`=${prevCol}${currentRow}+${currentCol}${fcfRow}`]];
    }
    currentRow += 2;

    // Apply formatting for negative numbers
    const rangeToFormat = sheet.getRange(`B${3}:${String.fromCharCode(65 + headers.length - 1)}${currentRow}`);
    rangeToFormat.numberFormat = [['#,##0.00_);[Red](#,##0.00)']];

    // Auto-size columns
    sheet.getRange(`A:${String.fromCharCode(65 + headers.length - 1)}`).format.autofitColumns();
    
    console.log('Cashflows layout created successfully');
  }

  formatDateForExcel(dateString) {
    if (!dateString) return '';
    const date = new Date(dateString);
    return date.toLocaleDateString();
  }

  calculatePeriods(startDate, endDate, periodType) {
    if (!startDate || !endDate) return 12; // Default to 12 periods
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    const diffTime = Math.abs(end - start);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    switch (periodType) {
      case 'daily':
        return Math.min(diffDays, 365); // Cap at 1 year for display
      case 'monthly':
        return Math.min(Math.ceil(diffDays / 30), 60); // Cap at 5 years
      case 'quarterly':
        return Math.min(Math.ceil(diffDays / 90), 20); // Cap at 5 years
      case 'yearly':
        return Math.min(Math.ceil(diffDays / 365), 10); // Cap at 10 years
      default:
        return 12;
    }
  }

  formatDateHeader(startDate, periodIndex, periodType) {
    const date = new Date(startDate);
    
    switch (periodType) {
      case 'daily':
        date.setDate(date.getDate() + periodIndex);
        return date.toLocaleDateString();
      case 'monthly':
        date.setMonth(date.getMonth() + periodIndex);
        return date.toLocaleDateString('en-US', { year: 'numeric', month: 'short' });
      case 'quarterly':
        date.setMonth(date.getMonth() + (periodIndex * 3));
        return `Q${Math.floor(periodIndex % 4) + 1} ${date.getFullYear()}`;
      case 'yearly':
        date.setFullYear(date.getFullYear() + periodIndex);
        return date.getFullYear().toString();
      default:
        date.setMonth(date.getMonth() + periodIndex);
        return date.toLocaleDateString('en-US', { year: 'numeric', month: 'short' });
    }
  }
}

// Export for use in main application
window.ExcelGenerator = ExcelGenerator;
window.CellReferenceTracker = CellReferenceTracker;