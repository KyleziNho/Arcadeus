/* global Office, Excel */

// Simple Cell Reference Tracker - keeps track of where data is stored
class CellTracker {
  constructor() {
    this.cellMap = new Map(); // Map of data keys to cell references
    this.sheetData = new Map(); // Map of sheet names to their data locations
  }

  // Record where a piece of data is stored
  recordCell(dataKey, sheetName, cellAddress) {
    const reference = `${sheetName}!${cellAddress}`;
    this.cellMap.set(dataKey, reference);
    
    // Also store by sheet for easy lookup
    if (!this.sheetData.has(sheetName)) {
      this.sheetData.set(sheetName, new Map());
    }
    this.sheetData.get(sheetName).set(dataKey, cellAddress);
    
    console.log(`📍 Recorded: ${dataKey} = ${reference}`);
  }

  // Get the cell reference for a piece of data
  getCellReference(dataKey) {
    return this.cellMap.get(dataKey) || null;
  }

  // Get all data for a specific sheet
  getSheetData(sheetName) {
    return this.sheetData.get(sheetName) || new Map();
  }

  // Print all tracked cells (for debugging)
  printAllCells() {
    console.log('📋 All tracked cells:');
    for (const [key, reference] of this.cellMap.entries()) {
      console.log(`  ${key}: ${reference}`);
    }
  }
}

class ExcelGenerator {
  constructor() {
    this.cellTracker = new CellTracker();
    this.currentWorkbook = null;
  }

  async generateModel(modelData) {
    try {
      console.log('🚀 Starting fresh model generation...');
      console.log('📊 Model data:', modelData);
      
      // Reset cell tracker
      this.cellTracker = new CellTracker();
      
      // Step 1: Create Assumptions sheet
      await this.createAssumptionsSheet(modelData);
      
      // Step 2: Create P&L sheet (optional - uncomment when ready)
      // await this.createPLSheet(modelData);
      
      console.log('✅ Model generation completed successfully!');
      this.cellTracker.printAllCells();
      
      return { success: true, message: 'Model created successfully!' };
      
    } catch (error) {
      console.error('❌ Error generating model:', error);
      return { success: false, error: error.message };
    }
  }

  async createAssumptionsSheet(modelData) {
    return Excel.run(async (context) => {
      console.log('📄 Creating Assumptions sheet...');
      
      const sheets = context.workbook.worksheets;
      
      // Delete existing Assumptions sheet if it exists
      try {
        const existingSheet = sheets.getItemOrNullObject('Assumptions');
        existingSheet.load('name');
        await context.sync();
        
        if (!existingSheet.isNullObject) {
          console.log('🗑️ Deleting existing Assumptions sheet');
          existingSheet.delete();
          await context.sync();
        }
      } catch (e) {
        // Sheet doesn't exist, continue
      }
      
      // Create new Assumptions sheet
      const sheet = sheets.add('Assumptions');
      sheet.activate();
      
      // Make sheet the active one
      await context.sync();
      
      // Now populate the sheet with data
      await this.populateAssumptionsSheet(context, sheet, modelData);
      
      console.log('✅ Assumptions sheet created successfully');
    });
  }

  async populateAssumptionsSheet(context, sheet, data) {
    console.log('📝 Populating Assumptions sheet with data...');
    
    let currentRow = 1;
    
    // HEADER
    sheet.getRange('A1').values = [['M&A Financial Model - Assumptions']];
    sheet.getRange('A1').format.font.bold = true;
    sheet.getRange('A1').format.font.size = 16;
    currentRow = 3;
    
    // Track section start rows for reference
    const sectionRows = {};
    
    // HIGH-LEVEL PARAMETERS SECTION
    sectionRows['highLevelParams'] = currentRow;
    sheet.getRange(`A${currentRow}`).values = [['HIGH-LEVEL PARAMETERS']];
    sheet.getRange(`A${currentRow}`).format.font.bold = true;
    currentRow += 2;
    
    // Currency
    sheet.getRange(`A${currentRow}`).values = [['Currency']];
    sheet.getRange(`B${currentRow}`).values = [[data.currency || 'USD']];
    this.cellTracker.recordCell('currency', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Project Start Date
    sheet.getRange(`A${currentRow}`).values = [['Project Start Date']];
    sheet.getRange(`B${currentRow}`).values = [[data.projectStartDate || '']];
    this.cellTracker.recordCell('projectStartDate', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Model Periods
    sheet.getRange(`A${currentRow}`).values = [['Model Periods']];
    sheet.getRange(`B${currentRow}`).values = [[data.modelPeriods || 'Monthly']];
    this.cellTracker.recordCell('modelPeriods', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Project End Date
    sheet.getRange(`A${currentRow}`).values = [['Project End Date']];
    sheet.getRange(`B${currentRow}`).values = [[data.projectEndDate || '']];
    this.cellTracker.recordCell('projectEndDate', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    currentRow += 2; // Add space
    
    // DEAL ASSUMPTIONS SECTION
    sectionRows['dealAssumptions'] = currentRow;
    sheet.getRange(`A${currentRow}`).values = [['DEAL ASSUMPTIONS']];
    sheet.getRange(`A${currentRow}`).format.font.bold = true;
    currentRow += 2;
    
    // Deal Name
    sheet.getRange(`A${currentRow}`).values = [['Deal Name']];
    sheet.getRange(`B${currentRow}`).values = [[data.dealName || '']];
    this.cellTracker.recordCell('dealName', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Deal Value
    sheet.getRange(`A${currentRow}`).values = [['Deal Value']];
    sheet.getRange(`B${currentRow}`).values = [[data.dealValue || 0]];
    this.cellTracker.recordCell('dealValue', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Transaction Fee
    sheet.getRange(`A${currentRow}`).values = [['Transaction Fee (%)']];
    sheet.getRange(`B${currentRow}`).values = [[data.transactionFee || 2.5]];
    this.cellTracker.recordCell('transactionFee', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Deal LTV
    sheet.getRange(`A${currentRow}`).values = [['Deal LTV (%)']];
    sheet.getRange(`B${currentRow}`).values = [[data.dealLTV || 70]];
    this.cellTracker.recordCell('dealLTV', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    currentRow += 2; // Add space
    
    // REVENUE ITEMS SECTION
    if (data.revenueItems && data.revenueItems.length > 0) {
      sectionRows['revenueItems'] = currentRow;
      sheet.getRange(`A${currentRow}`).values = [['REVENUE ITEMS']];
      sheet.getRange(`A${currentRow}`).format.font.bold = true;
      currentRow += 2;
      
      const revenueStartRow = currentRow;
      data.revenueItems.forEach((item, index) => {
        const itemName = item.name || `Revenue Item ${index + 1}`;
        sheet.getRange(`A${currentRow}`).values = [[itemName]];
        sheet.getRange(`B${currentRow}`).values = [[item.value || 0]];
        this.cellTracker.recordCell(`revenue_${index}`, 'Assumptions', `B${currentRow}`);
        this.cellTracker.recordCell(`revenue_${index}_name`, 'Assumptions', `A${currentRow}`);
        currentRow++;
      });
      
      // Record the range of revenue items for future reference
      this.cellTracker.recordCell('revenue_range', 'Assumptions', `B${revenueStartRow}:B${currentRow - 1}`);
      this.cellTracker.recordCell('revenue_count', 'Assumptions', data.revenueItems.length.toString());
      
      currentRow += 2; // Add space
    }
    
    // OPERATING EXPENSES SECTION
    if (data.operatingExpenses && data.operatingExpenses.length > 0) {
      sectionRows['operatingExpenses'] = currentRow;
      sheet.getRange(`A${currentRow}`).values = [['OPERATING EXPENSES']];
      sheet.getRange(`A${currentRow}`).format.font.bold = true;
      currentRow += 2;
      
      const opexStartRow = currentRow;
      data.operatingExpenses.forEach((item, index) => {
        const itemName = item.name || `OpEx Item ${index + 1}`;
        sheet.getRange(`A${currentRow}`).values = [[itemName]];
        sheet.getRange(`B${currentRow}`).values = [[item.value || 0]];
        this.cellTracker.recordCell(`opex_${index}`, 'Assumptions', `B${currentRow}`);
        this.cellTracker.recordCell(`opex_${index}_name`, 'Assumptions', `A${currentRow}`);
        currentRow++;
      });
      
      // Record the range of operating expenses for future reference
      this.cellTracker.recordCell('opex_range', 'Assumptions', `B${opexStartRow}:B${currentRow - 1}`);
      this.cellTracker.recordCell('opex_count', 'Assumptions', data.operatingExpenses.length.toString());
      
      currentRow += 2; // Add space
    }
    
    // CAPITAL EXPENSES SECTION
    if (data.capitalExpenses && data.capitalExpenses.length > 0) {
      sectionRows['capitalExpenses'] = currentRow;
      sheet.getRange(`A${currentRow}`).values = [['CAPITAL EXPENSES']];
      sheet.getRange(`A${currentRow}`).format.font.bold = true;
      currentRow += 2;
      
      const capexStartRow = currentRow;
      data.capitalExpenses.forEach((item, index) => {
        const itemName = item.name || `CapEx Item ${index + 1}`;
        sheet.getRange(`A${currentRow}`).values = [[itemName]];
        sheet.getRange(`B${currentRow}`).values = [[item.value || 0]];
        this.cellTracker.recordCell(`capex_${index}`, 'Assumptions', `B${currentRow}`);
        this.cellTracker.recordCell(`capex_${index}_name`, 'Assumptions', `A${currentRow}`);
        currentRow++;
      });
      
      // Record the range of capital expenses for future reference
      this.cellTracker.recordCell('capex_range', 'Assumptions', `B${capexStartRow}:B${currentRow - 1}`);
      this.cellTracker.recordCell('capex_count', 'Assumptions', data.capitalExpenses.length.toString());
      
      currentRow += 2; // Add space
    }
    
    // EXIT ASSUMPTIONS SECTION
    sectionRows['exitAssumptions'] = currentRow;
    sheet.getRange(`A${currentRow}`).values = [['EXIT ASSUMPTIONS']];
    sheet.getRange(`A${currentRow}`).format.font.bold = true;
    currentRow += 2;
    
    // Disposal Cost
    sheet.getRange(`A${currentRow}`).values = [['Disposal Cost (%)']];
    sheet.getRange(`B${currentRow}`).values = [[data.disposalCost || 2.5]];
    this.cellTracker.recordCell('disposalCost', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Terminal Cap Rate
    sheet.getRange(`A${currentRow}`).values = [['Terminal Cap Rate (%)']];
    sheet.getRange(`B${currentRow}`).values = [[data.terminalCapRate || 8.5]];
    this.cellTracker.recordCell('terminalCapRate', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Auto-resize columns
    sheet.getRange('A:B').format.autofitColumns();
    
    // Store section row information for reference
    this.cellTracker.recordCell('section_rows', 'Assumptions', JSON.stringify(sectionRows));
    
    await context.sync();
    console.log('✅ Assumptions sheet populated successfully');
    console.log('📍 Section positions:', sectionRows);
  }

  // P&L Sheet generation using dynamic cell references
  async createPLSheet(modelData) {
    return Excel.run(async (context) => {
      console.log('📈 Creating P&L sheet...');
      
      const sheets = context.workbook.worksheets;
      
      // Delete existing P&L sheet if it exists
      try {
        const existingSheet = sheets.getItemOrNullObject('P&L Statement');
        existingSheet.load('name');
        await context.sync();
        
        if (!existingSheet.isNullObject) {
          console.log('🗑️ Deleting existing P&L sheet');
          existingSheet.delete();
          await context.sync();
        }
      } catch (e) {
        // Sheet doesn't exist, continue
      }
      
      // Create new P&L sheet
      const plSheet = sheets.add('P&L Statement');
      await context.sync();
      
      // Build P&L with formulas referencing Assumptions
      let currentRow = 1;
      
      // HEADER
      plSheet.getRange('A1').values = [['P&L Statement']];
      plSheet.getRange('A1').format.font.bold = true;
      plSheet.getRange('A1').format.font.size = 16;
      currentRow = 3;
      
      // REVENUE SECTION
      plSheet.getRange(`A${currentRow}`).values = [['REVENUE']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      currentRow += 2;
      
      // Get revenue range from tracker
      const revenueRange = this.cellTracker.getCellReference('revenue_range');
      const revenueCount = parseInt(this.cellTracker.getCellReference('revenue_count') || '0');
      
      if (revenueCount > 0) {
        // Add each revenue item with formula
        for (let i = 0; i < revenueCount; i++) {
          const nameRef = this.cellTracker.getCellReference(`revenue_${i}_name`);
          const valueRef = this.cellTracker.getCellReference(`revenue_${i}`);
          
          if (nameRef && valueRef) {
            // Use formula to reference the name
            plSheet.getRange(`A${currentRow}`).formulas = [[`=${nameRef}`]];
            // Use formula to reference the value
            plSheet.getRange(`B${currentRow}`).formulas = [[`=${valueRef}`]];
            currentRow++;
          }
        }
        
        // Total Revenue - sum all revenue items
        currentRow++;
        plSheet.getRange(`A${currentRow}`).values = [['Total Revenue']];
        plSheet.getRange(`A${currentRow}`).format.font.bold = true;
        if (revenueRange) {
          plSheet.getRange(`B${currentRow}`).formulas = [[`=SUM(${revenueRange})`]];
        }
        currentRow += 2;
      }
      
      // OPERATING EXPENSES SECTION
      plSheet.getRange(`A${currentRow}`).values = [['OPERATING EXPENSES']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      currentRow += 2;
      
      const opexCount = parseInt(this.cellTracker.getCellReference('opex_count') || '0');
      
      if (opexCount > 0) {
        // Add each operating expense with formula
        for (let i = 0; i < opexCount; i++) {
          const nameRef = this.cellTracker.getCellReference(`opex_${i}_name`);
          const valueRef = this.cellTracker.getCellReference(`opex_${i}`);
          
          if (nameRef && valueRef) {
            plSheet.getRange(`A${currentRow}`).formulas = [[`=${nameRef}`]];
            plSheet.getRange(`B${currentRow}`).formulas = [[`=${valueRef}`]];
            currentRow++;
          }
        }
        
        // Total Operating Expenses
        currentRow++;
        plSheet.getRange(`A${currentRow}`).values = [['Total Operating Expenses']];
        plSheet.getRange(`A${currentRow}`).format.font.bold = true;
        const opexRange = this.cellTracker.getCellReference('opex_range');
        if (opexRange) {
          plSheet.getRange(`B${currentRow}`).formulas = [[`=SUM(${opexRange})`]];
        }
        currentRow += 2;
      }
      
      // EBITDA CALCULATION
      plSheet.getRange(`A${currentRow}`).values = [['EBITDA']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#87CEEB';
      
      // EBITDA = Total Revenue - Total Operating Expenses
      const totalRevenueRow = currentRow - 2 - opexCount - 2;
      const totalOpexRow = currentRow - 2;
      plSheet.getRange(`B${currentRow}`).formulas = [[`=B${totalRevenueRow}-B${totalOpexRow}`]];
      
      // Format numbers
      plSheet.getRange('B:B').numberFormat = [['#,##0.00']];
      
      // Auto-resize columns
      plSheet.getRange('A:B').format.autofitColumns();
      
      await context.sync();
      console.log('✅ P&L sheet created with dynamic references');
    });
  }

  // Utility method to get all tracked data
  getTrackedData() {
    return {
      cellMap: Object.fromEntries(this.cellTracker.cellMap),
      sheetData: Object.fromEntries(this.cellTracker.sheetData)
    };
  }
}

// Export for use in main application
window.ExcelGenerator = ExcelGenerator;
window.CellTracker = CellTracker;