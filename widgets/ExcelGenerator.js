/* global Office, Excel */

// Excel Formatter class for consistent formatting across sheets
class ExcelFormatter {
  static colors = {
    darkBlue: '#002060',
    lightGrey: '#D3D3D3',
    black: '#000000',
    white: '#FFFFFF',
    backgroundDarker5: '#F2F2F2'
  };
  
  static applyNumberFormat(range) {
    range.numberFormat = [['#,##0;[Red](#,##0);"-"']];
    // Ensure right alignment for dash values (preserves existing alignment)
    range.format.horizontalAlignment = 'Right';
    // Ensure Times New Roman font for all cells
    range.format.font.name = 'Times New Roman';
    range.format.font.size = 12;
  }
}

// Simple Cell Reference Tracker - keeps track of where data is stored
// Updated: Removed all fallback calculations - API required for IRR/MOIC accuracy - v2.1
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
    this.plCellTracker = new CellTracker(); // Track P&L cell references
    this.currentWorkbook = null;
  }

  // Helper function to format dates as XX-XXX-XX (last day of month)
  formatDateAsLastDay(date, periodType) {
    if (!date) return '';
    
    const d = new Date(date);
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    
    // Get the last day of the month
    const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
    const monthName = monthNames[d.getMonth()];
    const year = d.getFullYear().toString().slice(-2); // Last 2 digits of year
    
    return `${lastDay.toString().padStart(2, '0')}-${monthName}-${year}`;
  }

  // Helper function to calculate the previous period label based on period type and start date
  getPreviousPeriodLabel(startDate, periodType) {
    if (!startDate) return 'Period 0';
    
    const date = new Date(startDate);
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    
    switch (periodType) {
      case 'daily':
        // Previous day
        date.setDate(date.getDate() - 1);
        return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
        
      case 'monthly':
        // Previous month
        date.setMonth(date.getMonth() - 1);
        return `${monthNames[date.getMonth()]} ${date.getFullYear()}`;
        
      case 'quarterly':
        // Previous quarter
        const currentQuarter = Math.floor(date.getMonth() / 3) + 1;
        date.setMonth(date.getMonth() - 3);
        const prevQuarter = Math.floor(date.getMonth() / 3) + 1;
        return `Q${prevQuarter} ${date.getFullYear()}`;
        
      case 'yearly':
        // Previous year
        return (date.getFullYear() - 1).toString();
        
      default:
        return 'Period 0';
    }
  }

  async generateModel(modelData) {
    try {
      console.log('🚀 Starting fresh model generation...');
      console.log('📊 Model data:', modelData);
      console.log('📊 Revenue items received:', modelData.revenueItems);
      console.log('📊 Operating expenses received:', modelData.operatingExpenses);
      console.log('📊 Capital expenses received:', modelData.capitalExpenses);
      
      // Only reset cell trackers if they're empty (first time)
      // This preserves references between Assumptions and P&L generation
      if (!this.cellTracker || this.cellTracker.cellMap.size === 0) {
        this.cellTracker = new CellTracker();
      }
      if (!this.plCellTracker || this.plCellTracker.cellMap.size === 0) {
        this.plCellTracker = new CellTracker();
      }
      
      // Step 1: Create Assumptions sheet only
      await this.createAssumptionsSheet(modelData);
      
      console.log('✅ Assumptions sheet generation completed successfully!');
      this.cellTracker.printAllCells();
      
      return { success: true, message: 'Assumptions sheet created successfully! You can now generate the P&L using AI.' };
      
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
    console.log('📝 ====== POPULATING ASSUMPTIONS SHEET ======');
    console.log('📝 Received data object:', data);
    console.log('📝 Revenue items received:', data.revenueItems);
    
    let currentRow = 1;
    
    // Hide gridlines
    sheet.showGridlines = false;
    
    // TITLE: '[Insert company name] - Assumptions' spanning two cells, left-aligned, light grey background
    const title = `${data.dealName || '[Insert company name]'} - Assumptions`;
    sheet.getRange('A1:B1').merge();
    sheet.getRange('A1').values = [[title]];
    const titleRange = sheet.getRange('A1');
    titleRange.format.font.name = 'Times New Roman';
    titleRange.format.font.size = 12;
    titleRange.format.font.bold = true;
    titleRange.format.horizontalAlignment = 'Left';
    titleRange.format.fill.color = ExcelFormatter.colors.backgroundDarker5; // Light grey background
    currentRow = 3;
    
    // Track section start rows for reference
    const sectionRows = {};
    
    // HIGH-LEVEL PARAMETERS (Currency, Start Date, Model Periods, End Date) - right aligned
    sheet.getRange(`A${currentRow}`).values = [['Currency']];
    sheet.getRange(`B${currentRow}`).values = [[data.currency || 'USD']];
    sheet.getRange(`B${currentRow}`).format.horizontalAlignment = 'Right';
    this.cellTracker.recordCell('currency', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    sheet.getRange(`A${currentRow}`).values = [['Start Date']];
    sheet.getRange(`B${currentRow}`).values = [[data.projectStartDate || '']];
    sheet.getRange(`B${currentRow}`).format.horizontalAlignment = 'Right';
    this.cellTracker.recordCell('projectStartDate', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    sheet.getRange(`A${currentRow}`).values = [['Model Periods']];
    sheet.getRange(`B${currentRow}`).values = [[data.modelPeriods || 'Monthly']];
    sheet.getRange(`B${currentRow}`).format.horizontalAlignment = 'Right';
    this.cellTracker.recordCell('modelPeriods', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    sheet.getRange(`A${currentRow}`).values = [['End Date']];
    sheet.getRange(`B${currentRow}`).values = [[data.projectEndDate || '']];
    sheet.getRange(`B${currentRow}`).format.horizontalAlignment = 'Right';
    this.cellTracker.recordCell('projectEndDate', 'Assumptions', `B${currentRow}`);
    currentRow += 2;
    
    // ACQUISITION ASSUMPTIONS - dark blue background, spans two cells, left aligned, white text
    sectionRows['dealAssumptions'] = currentRow;
    sheet.getRange(`A${currentRow}:B${currentRow}`).merge();
    sheet.getRange(`A${currentRow}`).values = [['Acquisition Assumptions']];
    const acqAssumptionsRange = sheet.getRange(`A${currentRow}`);
    acqAssumptionsRange.format.font.name = 'Times New Roman';
    acqAssumptionsRange.format.font.size = 12;
    acqAssumptionsRange.format.font.bold = true;
    acqAssumptionsRange.format.horizontalAlignment = 'Left';
    acqAssumptionsRange.format.fill.color = ExcelFormatter.colors.darkBlue;
    acqAssumptionsRange.format.font.color = ExcelFormatter.colors.white;
    currentRow++;
    
    // Deal Value
    sheet.getRange(`A${currentRow}`).values = [['Deal Value']];
    sheet.getRange(`B${currentRow}`).values = [[data.dealValue || 0]];
    sheet.getRange(`B${currentRow}`).numberFormat = '#,##0';
    this.cellTracker.recordCell('dealValue', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Transaction Fee
    sheet.getRange(`A${currentRow}`).values = [['Transaction Fee (%)']];
    sheet.getRange(`B${currentRow}`).values = [[(data.transactionFee || 2.5) / 100]];
    sheet.getRange(`B${currentRow}`).numberFormat = '0.00%';
    this.cellTracker.recordCell('transactionFee', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Deal LTV
    sheet.getRange(`A${currentRow}`).values = [['Deal LTV (%)']];
    sheet.getRange(`B${currentRow}`).values = [[data.dealLTV || 70]];
    this.cellTracker.recordCell('dealLTV', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Equity Contribution (Calculated)
    sheet.getRange(`A${currentRow}`).values = [['Equity Contribution']];
    const dealValueCell = this.cellTracker.getCellReference('dealValue').split('!')[1];
    const ltvCell = this.cellTracker.getCellReference('dealLTV').split('!')[1];
    sheet.getRange(`B${currentRow}`).formulas = [[`=${dealValueCell}*(1-${ltvCell}/100)`]];
    sheet.getRange(`B${currentRow}`).numberFormat = '#,##0';
    this.cellTracker.recordCell('equityContribution', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Debt Financing (Calculated)
    sheet.getRange(`A${currentRow}`).values = [['Debt Financing']];
    sheet.getRange(`B${currentRow}`).formulas = [[`=${dealValueCell}*${ltvCell}/100`]];
    sheet.getRange(`B${currentRow}`).numberFormat = '#,##0';
    this.cellTracker.recordCell('debtFinancing', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Loan Issuance Fees (if there's debt)
    if (data.dealLTV && parseFloat(data.dealLTV) > 0) {
      sheet.getRange(`A${currentRow}`).values = [['Loan Issuance Fees (%)']];
      sheet.getRange(`B${currentRow}`).values = [[(data.debtSettings?.loanIssuanceFees || 1.5) / 100]];
      sheet.getRange(`B${currentRow}`).numberFormat = '0.00%';
      this.cellTracker.recordCell('loanIssuanceFees', 'Assumptions', `B${currentRow}`);
      currentRow++;
    }
    currentRow++;
    
    // REVENUE ITEMS - dark blue background, spans 3 cells (A:C), white text
    if (data.revenueItems && data.revenueItems.length > 0) {
      sectionRows['revenueItems'] = currentRow;
      sheet.getRange(`A${currentRow}:C${currentRow}`).merge();
      sheet.getRange(`A${currentRow}`).values = [['Revenue Items']];
      const revenueHeaderRange = sheet.getRange(`A${currentRow}`);
      revenueHeaderRange.format.font.name = 'Times New Roman';
      revenueHeaderRange.format.font.size = 12;
      revenueHeaderRange.format.font.bold = true;
      revenueHeaderRange.format.horizontalAlignment = 'Left';
      revenueHeaderRange.format.fill.color = ExcelFormatter.colors.darkBlue;
      revenueHeaderRange.format.font.color = ExcelFormatter.colors.white;
      currentRow++;
      
      // Add subheaders in the same row as header
      sheet.getRange(`A${currentRow-1}:C${currentRow-1}`).unmerge();
      sheet.getRange(`A${currentRow-1}`).values = [['Revenue Items']];
      sheet.getRange(`B${currentRow-1}`).values = [['Value']];
      sheet.getRange(`C${currentRow-1}`).values = [['Growth Rate']];
      const revenueSubHeaderRange = sheet.getRange(`A${currentRow-1}:C${currentRow-1}`);
      revenueSubHeaderRange.format.font.name = 'Times New Roman';
      revenueSubHeaderRange.format.font.size = 12;
      revenueSubHeaderRange.format.font.bold = true;
      revenueSubHeaderRange.format.horizontalAlignment = 'Left';
      revenueSubHeaderRange.format.fill.color = ExcelFormatter.colors.darkBlue;
      revenueSubHeaderRange.format.font.color = ExcelFormatter.colors.white;
      
      const revenueStartRow = currentRow;
      data.revenueItems.forEach((item, index) => {
        const itemName = item.name || `Revenue Item ${index + 1}`;
        const growthRate = item.growthRate || 0;
        sheet.getRange(`A${currentRow}`).values = [[itemName]];
        sheet.getRange(`B${currentRow}`).values = [[item.value || 0]];
        sheet.getRange(`B${currentRow}`).numberFormat = '#,##0';
        sheet.getRange(`C${currentRow}`).values = [[`${growthRate}%`]];
        this.cellTracker.recordCell(`revenue_${index}`, 'Assumptions', `B${currentRow}`);
        this.cellTracker.recordCell(`revenue_${index}_name`, 'Assumptions', `A${currentRow}`);
        this.cellTracker.recordCell(`revenue_${index}_growth_rate`, 'Assumptions', `C${currentRow}`);
        currentRow++;
      });
      
      this.cellTracker.recordCell('revenue_range', 'Assumptions', `B${revenueStartRow}:B${currentRow - 1}`);
      this.cellTracker.recordCell('revenue_count', 'Assumptions', data.revenueItems.length.toString());
      currentRow++;
    }
    
    // COST ITEMS (OPERATING EXPENSES) - dark blue background, spans 3 cells (A:C), white text  
    if (data.operatingExpenses && data.operatingExpenses.length > 0) {
      sectionRows['operatingExpenses'] = currentRow;
      sheet.getRange(`A${currentRow}`).values = [['Cost Items']];
      sheet.getRange(`B${currentRow}`).values = [['Value']];
      sheet.getRange(`C${currentRow}`).values = [['Growth Rate']];
      const costHeaderRange = sheet.getRange(`A${currentRow}:C${currentRow}`);
      costHeaderRange.format.font.name = 'Times New Roman';
      costHeaderRange.format.font.size = 12;
      costHeaderRange.format.font.bold = true;
      costHeaderRange.format.horizontalAlignment = 'Left';
      costHeaderRange.format.fill.color = ExcelFormatter.colors.darkBlue;
      costHeaderRange.format.font.color = ExcelFormatter.colors.white;
      currentRow++;
      
      const opexStartRow = currentRow;
      data.operatingExpenses.forEach((item, index) => {
        const itemName = item.name || `Cost Item ${index + 1}`;
        const growthRate = item.growthRate || 0;
        sheet.getRange(`A${currentRow}`).values = [[itemName]];
        sheet.getRange(`B${currentRow}`).values = [[item.value || 0]];
        sheet.getRange(`B${currentRow}`).numberFormat = '#,##0';
        sheet.getRange(`C${currentRow}`).values = [[`${growthRate}%`]];
        this.cellTracker.recordCell(`opex_${index}`, 'Assumptions', `B${currentRow}`);
        this.cellTracker.recordCell(`opex_${index}_name`, 'Assumptions', `A${currentRow}`);
        this.cellTracker.recordCell(`opex_${index}_growth_rate`, 'Assumptions', `C${currentRow}`);
        currentRow++;
      });
      
      this.cellTracker.recordCell('opex_range', 'Assumptions', `B${opexStartRow}:B${currentRow - 1}`);
      this.cellTracker.recordCell('opex_count', 'Assumptions', data.operatingExpenses.length.toString());
      currentRow++;
    }
    
    // CAPEX ITEMS - dark blue background, spans 3 cells (A:C), white text
    if (data.capEx && data.capEx.length > 0) {
      sectionRows['capEx'] = currentRow;
      sheet.getRange(`A${currentRow}`).values = [['CapEx Items']];
      sheet.getRange(`B${currentRow}`).values = [['Value']];
      sheet.getRange(`C${currentRow}`).values = [['Growth Rate']];
      const capexHeaderRange = sheet.getRange(`A${currentRow}:C${currentRow}`);
      capexHeaderRange.format.font.name = 'Times New Roman';
      capexHeaderRange.format.font.size = 12;
      capexHeaderRange.format.font.bold = true;
      capexHeaderRange.format.horizontalAlignment = 'Left';
      capexHeaderRange.format.fill.color = ExcelFormatter.colors.darkBlue;
      capexHeaderRange.format.font.color = ExcelFormatter.colors.white;
      currentRow++;
      
      const capexStartRow = currentRow;
      data.capEx.forEach((item, index) => {
        const itemName = item.name || `CapEx ${index + 1}`;
        const growthRate = item.growthRate || 0;
        sheet.getRange(`A${currentRow}`).values = [[itemName]];
        sheet.getRange(`B${currentRow}`).values = [[item.value || 0]];
        sheet.getRange(`B${currentRow}`).numberFormat = '#,##0';
        sheet.getRange(`C${currentRow}`).values = [[`${growthRate}%`]];
        this.cellTracker.recordCell(`capex_${index}`, 'Assumptions', `B${currentRow}`);
        this.cellTracker.recordCell(`capex_${index}_name`, 'Assumptions', `A${currentRow}`);
        this.cellTracker.recordCell(`capex_${index}_growth_rate`, 'Assumptions', `C${currentRow}`);
        currentRow++;
      });
      
      this.cellTracker.recordCell('capex_range', 'Assumptions', `B${capexStartRow}:B${currentRow - 1}`);
      this.cellTracker.recordCell('capex_count', 'Assumptions', data.capEx.length.toString());
      currentRow++;
    }
    
    // EXIT ASSUMPTIONS - dark blue background, spans two cells, white text
    sectionRows['exitAssumptions'] = currentRow;
    sheet.getRange(`A${currentRow}:B${currentRow}`).merge();
    sheet.getRange(`A${currentRow}`).values = [['Exit Assumptions']];
    const exitAssumptionsRange = sheet.getRange(`A${currentRow}`);
    exitAssumptionsRange.format.font.name = 'Times New Roman';
    exitAssumptionsRange.format.font.size = 12;
    exitAssumptionsRange.format.font.bold = true;
    exitAssumptionsRange.format.horizontalAlignment = 'Left';
    exitAssumptionsRange.format.fill.color = ExcelFormatter.colors.darkBlue;
    exitAssumptionsRange.format.font.color = ExcelFormatter.colors.white;
    currentRow++;
    
    // Disposal Cost
    sheet.getRange(`A${currentRow}`).values = [['Disposal Cost (%)']];
    sheet.getRange(`B${currentRow}`).values = [[(data.disposalCost || 2.5) / 100]];
    sheet.getRange(`B${currentRow}`).numberFormat = '0.00%';
    this.cellTracker.recordCell('disposalCost', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Terminal Cap Rate  
    sheet.getRange(`A${currentRow}`).values = [['Terminal Cap Rate (%)']];
    sheet.getRange(`B${currentRow}`).values = [[(data.terminalCapRate || 8.5) / 100]];
    sheet.getRange(`B${currentRow}`).numberFormat = '0.00%';
    this.cellTracker.recordCell('terminalCapRate', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Discount Rate (WACC)
    sheet.getRange(`A${currentRow}`).values = [['Discount Rate - WACC (%)']];
    sheet.getRange(`B${currentRow}`).values = [[(data.discountRate || 10.0) / 100]];
    sheet.getRange(`B${currentRow}`).numberFormat = '0.00%';
    this.cellTracker.recordCell('discountRate', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Apply Times New Roman font to all cells
    const allCellsRange = sheet.getUsedRange();
    allCellsRange.format.font.name = 'Times New Roman';
    allCellsRange.format.font.size = 12;
    
    // Auto-resize columns
    sheet.getRange('A:C').format.autofitColumns();
    
    // Store section row information for reference
    this.cellTracker.recordCell('section_rows', 'Assumptions', JSON.stringify(sectionRows));
    
    await context.sync();
    console.log('✅ Assumptions sheet populated successfully');
    console.log('📍 Section positions:', sectionRows);
  }

  // Generate detailed OpenAI prompt for P&L creation
  generateDetailedAIPrompt(modelData) {
    console.log('🤖 Generating AI prompt for P&L creation...');
    
    const assumptions = [];
    const cellRefs = [];
    
    // Compile all assumptions with their cell references
    for (const [key, reference] of this.cellTracker.cellMap.entries()) {
      assumptions.push({
        key: key,
        reference: reference,
        value: this.getValueForKey(key, modelData)
      });
    }
    
    // Calculate the number of periods
    const periods = this.calculatePeriods(modelData.projectStartDate, modelData.projectEndDate, modelData.modelPeriods);
    const maxPeriods = Math.min(periods, 60); // Cap at 60 periods for performance
    
    // Generate period headers
    const periodHeaders = [];
    const startDate = new Date(modelData.projectStartDate);
    for (let i = 0; i < maxPeriods; i++) {
      periodHeaders.push(this.formatPeriodHeader(startDate, i, modelData.modelPeriods));
    }
    
    const prompt = `You are a senior financial analyst at a top-tier investment bank specializing in M&A financial modeling. You have been provided with a complete set of assumptions stored in an Excel 'Assumptions' sheet with specific cell references.

**PROJECT OVERVIEW:**
- Deal Name: ${modelData.dealName}
- Currency: ${modelData.currency}
- Model Period Type: ${modelData.modelPeriods}
- Project Duration: ${modelData.projectStartDate} to ${modelData.projectEndDate}
- Total Periods Required: ${maxPeriods}

**EXACT CELL REFERENCES IN ASSUMPTIONS SHEET:**

**High-Level Parameters:**
- Currency: ${this.cellTracker.getCellReference('currency')}
- Project Start Date: ${this.cellTracker.getCellReference('projectStartDate')}
- Project End Date: ${this.cellTracker.getCellReference('projectEndDate')}
- Model Periods: ${this.cellTracker.getCellReference('modelPeriods')}

**Deal Structure:**
- Deal Name: ${this.cellTracker.getCellReference('dealName')}
- Deal Value: ${this.cellTracker.getCellReference('dealValue')}
- Transaction Fee %: ${this.cellTracker.getCellReference('transactionFee')}
- Deal LTV %: ${this.cellTracker.getCellReference('dealLTV')}
- Equity Contribution (Calculated): ${this.cellTracker.getCellReference('equityContribution')}
- Debt Financing (Calculated): ${this.cellTracker.getCellReference('debtFinancing')}

**Revenue Items with Growth Rates:**
${this.formatDetailedRevenueItems(modelData)}

**Operating Expense Items with Growth Rates:**
${this.formatDetailedOpexItems(modelData)}

**Capital Expense Items with Growth Rates:**
${this.formatDetailedCapexItems(modelData)}

**Exit Assumptions:**
- Disposal Cost %: ${this.cellTracker.getCellReference('disposalCost')}
- Terminal Cap Rate %: ${this.cellTracker.getCellReference('terminalCapRate')}

**Debt Model (if LTV > 0):**
${this.formatDetailedDebtModel(modelData)}

**PERIOD HEADERS REQUIRED:**
${periodHeaders.join(', ')}

**DETAILED INSTRUCTIONS:**

1. **Create a comprehensive P&L Statement** with the following structure:
   - Column A: Line item names
   - Columns B through ${this.getColumnLetter(maxPeriods)}: Period data (${maxPeriods} periods total)

2. **Revenue Section:**
   - List each revenue item from the cell references above
   - For period 1: Use the base value from the assumption cell
   - For subsequent periods: Apply growth formulas adjusted for period type:
     * If growth type is 'annual' and model periods are:
       - Daily: Previous period * (1 + annual_rate/365/100)
       - Monthly: Previous period * (1 + annual_rate/12/100)
       - Quarterly: Previous period * (1 + annual_rate/4/100)
       - Yearly: Previous period * (1 + annual_rate/100)
     * If growth type is 'none': Use same value as previous period
   - Include a 'Total Revenue' row that sums all revenue items

3. **Operating Expenses Section:**
   - List each operating expense item (as negative values)
   - Apply the same growth logic as revenue
   - Include a 'Total Operating Expenses' row

4. **NOI Calculation:**
   - NOI = Total Revenue + Total Operating Expenses (expenses are negative)

5. **Capital Expenses (if any):**
   - List each capital expense item (as negative values)
   - Apply growth formulas
   - Include 'Total CapEx' row

6. **REAL ESTATE MODEL - P&L ENDS AT NOI:**
   - For real estate investments, P&L statement ends at Net Operating Income (NOI)
   - NO interest expense, tax expense, or net income calculations
   - NOI is the key metric for real estate financial analysis

8. **FORMAT REQUIREMENTS:**
   - Use exact Excel formula syntax
   - Reference cells using 'Assumptions!CellAddress' format
   - Provide the complete Excel range setup
   - Include proper headers and formatting instructions
   - Make sure all ${maxPeriods} periods are covered

**CRITICAL:** You must provide exact Excel formulas for every cell, referencing the specific assumption cells listed above. Do not use placeholder values - use actual Excel formulas that will calculate correctly.

Please provide the complete P&L structure with exact cell addresses and formulas for all ${maxPeriods} periods.`;

    console.log('📝 Generated detailed AI prompt with', maxPeriods, 'periods');
    return prompt;
  }

  // Format detailed revenue items with exact cell references
  formatDetailedRevenueItems(modelData) {
    if (!modelData.revenueItems || modelData.revenueItems.length === 0) {
      return 'No revenue items specified.';
    }
    
    let output = '';
    modelData.revenueItems.forEach((item, index) => {
      const nameRef = this.cellTracker.getCellReference(`revenue_${index}_name`);
      const valueRef = this.cellTracker.getCellReference(`revenue_${index}`);
      const growthRateRef = this.cellTracker.getCellReference(`revenue_${index}_growth_rate`);
      
      output += `\n- ${item.name || `Revenue Item ${index + 1}`}:\n`;
      output += `  * Base Value Cell: ${valueRef}\n`;
      output += `  * Growth Rate: ${item.growthRate || 0}% (Cell: ${growthRateRef})\n`;
      output += `  * Growth Type: Linear (annual)\n`;
    });
    return output;
  }
  
  // Format detailed operating expense items
  formatDetailedOpexItems(modelData) {
    if (!modelData.operatingExpenses || modelData.operatingExpenses.length === 0) {
      return 'No operating expense items specified.';
    }
    
    let output = '';
    modelData.operatingExpenses.forEach((item, index) => {
      const nameRef = this.cellTracker.getCellReference(`opex_${index}_name`);
      const valueRef = this.cellTracker.getCellReference(`opex_${index}`);
      const growthRateRef = this.cellTracker.getCellReference(`opex_${index}_growth_rate`);
      
      output += `\n- ${item.name || `OpEx Item ${index + 1}`}:\n`;
      output += `  * Base Value Cell: ${valueRef}\n`;
      output += `  * Growth Rate: ${item.growthRate || 0}% (Cell: ${growthRateRef})\n`;
      output += `  * Growth Type: Linear (annual)\n`;
    });
    return output;
  }
  
  // Format detailed operating expenses (alias for formatDetailedOpexItems)
  formatDetailedOperatingExpenses(modelData) {
    if (!modelData.operatingExpenses || modelData.operatingExpenses.length === 0) {
      return 'No operating expense items specified.';
    }
    
    let output = '';
    modelData.operatingExpenses.forEach((item, index) => {
      const nameRef = this.cellTracker.getCellReference(`opex_${index}_name`);
      const valueRef = this.cellTracker.getCellReference(`opex_${index}`);
      const growthRateRef = this.cellTracker.getCellReference(`opex_${index}_growth_rate`);
      
      output += `\n- ${item.name || `OpEx Item ${index + 1}`}:\n`;
      output += `  * Base Value Cell: ${valueRef}\n`;
      output += `  * Growth Rate: ${item.growthRate || 0}% (Cell: ${growthRateRef})\n`;
      output += `  * Growth Type: Linear (annual)\n`;
    });
    return output;
  }
  
  // Format detailed capital expense items
  formatDetailedCapexItems(modelData) {
    if (!modelData.capitalExpenses || modelData.capitalExpenses.length === 0) {
      return 'No capital expense items specified.';
    }
    
    let output = '';
    modelData.capitalExpenses.forEach((item, index) => {
      const nameRef = this.cellTracker.getCellReference(`capex_${index}_name`);
      const valueRef = this.cellTracker.getCellReference(`capex_${index}`);
      const growthRateRef = this.cellTracker.getCellReference(`capex_${index}_growth_rate`);
      
      output += `\n- ${item.name || `CapEx Item ${index + 1}`}:\n`;
      output += `  * Base Value Cell: ${valueRef}\n`;
      output += `  * Growth Rate: ${item.growthRate || 0}% (Cell: ${growthRateRef})\n`;
      output += `  * Growth Type: Linear (annual)\n`;
    });
    return output;
  }
  
  // Format detailed debt model information
  formatDetailedDebtModel(modelData) {
    const hasDebt = modelData.dealLTV && parseFloat(modelData.dealLTV) > 0;
    
    if (!hasDebt) {
      return 'No debt financing (LTV = 0%)';
    }
    
    let output = `\n- Loan Issuance Fees: ${this.cellTracker.getCellReference('loanIssuanceFees')}\n`;
    output += `- Fixed Interest Rate: ${this.cellTracker.getCellReference('fixedRate')}\n`;
    
    return output;
  }
  
  // Generate P&L with AI using actual cell references
  async generatePLWithAI(modelData) {
    try {
      console.log('📈 Generating P&L Statement with AI...');
      
      // Generate comprehensive AI prompt with cell references
      const aiPrompt = this.generateEnhancedPLPrompt(modelData);
      
      // Call OpenAI to generate P&L formulas
      console.log('🤖 Calling OpenAI for P&L generation...');
      const aiResponse = await this.callOpenAIForPL(aiPrompt);
      
      // Create P&L sheet based on AI response
      await this.createAIPLSheet(modelData, aiResponse);
      
      return { success: true, message: 'AI-powered P&L Statement generated successfully!' };
      
    } catch (error) {
      console.error('❌ Error generating AI P&L:', error);
      // Fallback to hardcoded version if AI fails
      console.log('⚠️ Falling back to template-based P&L generation...');
      await this.createPLSheet(modelData);
      return { success: true, message: 'P&L Statement generated (template mode)' };
    }
  }
  
  // Generate Free Cash Flow with AI
  async generateFCFWithAI(modelData) {
    try {
      console.log('💰 Generating Free Cash Flow Statement with AI...');
      
      // Step 1: Read the actual P&L sheet to get real cell structure
      const plStructure = await this.readPLSheetStructure();
      console.log('📊 P&L Structure discovered:', plStructure);
      
      // Step 2: Read assumption sheet structure
      const assumptionStructure = await this.readAssumptionSheetStructure();
      console.log('📊 Assumption Structure discovered:', assumptionStructure);
      
      // Step 3: Read the existing CapEx sheet structure
      console.log('🔍 About to read CapEx sheet structure...');
      const capExStructure = await this.readCapExSheetStructure() || { totalRow: null, sheet: null };
      console.log('📊 CapEx Structure discovered:', capExStructure);
      console.log('🔍 capExStructure type:', typeof capExStructure);
      
      // Step 4: Generate comprehensive FCF AI prompt with ACTUAL cell references
      console.log('🔍 About to generate FCF prompt...');
      const fcfPrompt = this.generateRealFCFPrompt(modelData, plStructure, assumptionStructure, capExStructure);
      console.log('🔍 FCF prompt generated successfully');
      
      // Step 5: Create professional FCF sheet using discovered cell references
      await this.createAIFCFSheet(modelData, fcfPrompt, plStructure, assumptionStructure, capExStructure);
      
      console.log('📋 REAL FCF AI Prompt for OpenAI:');
      console.log('='.repeat(100));
      console.log(fcfPrompt);
      console.log('='.repeat(100));
      
      return { success: true, message: 'Professional Free Cash Flow Statement generated using real P&L cell references!' };
      
    } catch (error) {
      console.error('❌ Error generating FCF:', error);
      return { success: false, error: error.message };
    }
  }
  
  async generateMultiplesAndIRR(modelData) {
    console.log('🔥 NEW CLEAN VERSION v5.0 - ZERO FALLBACK LOGIC');
    console.log('🔥 TIMESTAMP:', new Date().toISOString());
    
    try {
      // Validate inputs
      if (!modelData.dealValue || modelData.dealValue === 0) {
        throw new Error('Deal value is required for IRR/MOIC calculation');
      }
      
      // Read FCF data from existing sheet
      const fcfData = await this.readFCFSheetData();
      console.log('📊 FCF data read:', fcfData);
      
      // Calculate equity contribution
      let equityContribution = modelData.equityContribution;
      if (!equityContribution || equityContribution === 0) {
        const dealLTV = modelData.dealLTV || 70;
        equityContribution = modelData.dealValue * (100 - dealLTV) / 100;
      }
      
      // Call AI API for IRR/MOIC formulas
      const aiPrompt = `Calculate IRR and MOIC for M&A deal. Return Excel formulas in JSON format.
      
Deal Value: ${modelData.dealValue}
Equity: ${equityContribution}
FCF Data: ${JSON.stringify(fcfData)}

Required format:
{
  "calculations": {
    "leveredIRR": {"formula": "=IRR formula here"},
    "leveredMOIC": {"formula": "=MOIC formula here"}
  }
}`;
      
      console.log('🤖 Calling AI API...');
      const aiResponse = await this.callOpenAIForMultiples(aiPrompt);
      console.log('🤖 AI response:', aiResponse);
      
      // Create Excel sheet with AI results
      await this.createCleanMultiplesSheet(modelData, aiResponse, equityContribution);
      
      return { success: true, message: 'IRR & MOIC Analysis created successfully!' };
      
    } catch (error) {
      console.error('❌ IRR/MOIC generation failed:', error);
      throw new Error(`IRR/MOIC calculation failed: ${error.message}`);
    }
  }
  
  // Create the actual P&L sheet with formulas
  async createPLSheet(modelData) {
    return Excel.run(async (context) => {
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
      
      // Calculate periods and prepare headers
      const periods = this.calculatePeriods(modelData.projectStartDate, modelData.projectEndDate, modelData.modelPeriods);
      const periodColumns = periods; // Use full calculated periods
      const totalColumns = periodColumns + 1; // +1 for Period 0 column
      
      let currentRow = 1;
      
      // TITLE
      plSheet.getRange('A1').values = [[`Profit & Loss Statement`]];
      const titleRange = plSheet.getRange(`A1:${this.getColumnLetter(totalColumns)}1`);
      titleRange.merge();
      titleRange.format.font.name = 'Times New Roman';
      titleRange.format.font.size = 12;
      titleRange.format.font.bold = true;
      titleRange.format.fill.color = ExcelFormatter.colors.backgroundDarker5;
      titleRange.format.horizontalAlignment = 'Left';
      currentRow = 2; // Remove extra blank row
      
      // REVENUE ITEMS SECTION
      const revenueStartRow = currentRow;
      plSheet.getRange(`A${currentRow}`).values = [['Revenue Items']];
      const revenueSectionRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      revenueSectionRange.format.font.name = 'Times New Roman';
      revenueSectionRange.format.font.size = 12;
      revenueSectionRange.format.font.bold = true;
      revenueSectionRange.format.fill.color = ExcelFormatter.colors.darkBlue;
      revenueSectionRange.format.font.color = ExcelFormatter.colors.white;
      
      // Add thin grey underline after section header
      revenueSectionRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
      revenueSectionRange.format.borders.getItem('EdgeBottom').weight = 'Thin';
      revenueSectionRange.format.borders.getItem('EdgeBottom').color = ExcelFormatter.colors.lightGrey;
      currentRow++;
      
      // Skip one row then add DATES (two cells below Revenue Items title)
      // Set the blank row height to 8
      plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`).format.rowHeight = 8;
      currentRow++;
      
      // DATES ROW - two cells below Revenue Items title (includes Period 0)
      const dateHeaders = [''];
      
      // Add Period 0 date first
      const prevPeriodDate = new Date(modelData.projectStartDate);
      switch (modelData.modelPeriods) {
        case 'daily':
          prevPeriodDate.setDate(prevPeriodDate.getDate() - 1);
          break;
        case 'monthly':
          prevPeriodDate.setMonth(prevPeriodDate.getMonth() - 1);
          break;
        case 'quarterly':
          prevPeriodDate.setMonth(prevPeriodDate.getMonth() - 3);
          break;
        case 'yearly':
          prevPeriodDate.setFullYear(prevPeriodDate.getFullYear() - 1);
          break;
      }
      dateHeaders.push(this.formatDateAsLastDay(prevPeriodDate, modelData.modelPeriods));
      
      // Add regular period dates
      const startDate = new Date(modelData.projectStartDate);
      for (let i = 0; i < periodColumns; i++) {
        const periodDate = new Date(startDate);
        
        // Calculate the date for this period
        switch (modelData.modelPeriods) {
          case 'daily':
            periodDate.setDate(periodDate.getDate() + i);
            break;
          case 'monthly':
            periodDate.setMonth(periodDate.getMonth() + i);
            break;
          case 'quarterly':
            periodDate.setMonth(periodDate.getMonth() + (i * 3));
            break;
          case 'yearly':
            periodDate.setFullYear(periodDate.getFullYear() + i);
            break;
        }
        
        dateHeaders.push(this.formatDateAsLastDay(periodDate, modelData.modelPeriods));
      }
      
      const dateRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      dateRange.values = [dateHeaders];
      dateRange.format.font.name = 'Times New Roman';
      dateRange.format.font.size = 12;
      dateRange.format.font.bold = true;
      dateRange.format.fill.color = ExcelFormatter.colors.white;
      dateRange.format.font.color = ExcelFormatter.colors.black;
      
      // Add dashed underline under dates
      dateRange.format.borders.getItem('EdgeBottom').style = 'Dash';
      dateRange.format.borders.getItem('EdgeBottom').weight = 'Thin';
      dateRange.format.borders.getItem('EdgeBottom').color = ExcelFormatter.colors.black;
      currentRow++;
      
      // PERIOD ROW - 'Period' followed by 0, 1, 2, 3, etc.
      const periodHeaders = ['Period'];
      periodHeaders.push('0'); // Period 0
      for (let i = 1; i <= periodColumns; i++) {
        periodHeaders.push(i.toString());
      }
      
      const periodRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      periodRange.values = [periodHeaders];  
      periodRange.format.font.name = 'Times New Roman';
      periodRange.format.font.size = 12;
      periodRange.format.font.bold = false; // Remove bold formatting
      currentRow++;
      
      // Add each revenue item
      if (modelData.revenueItems && modelData.revenueItems.length > 0) {
        modelData.revenueItems.forEach((item, index) => {
          plSheet.getRange(`A${currentRow}`).values = [[item.name || `Revenue ${index + 1}`]];
          plSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
          plSheet.getRange(`A${currentRow}`).format.font.size = 12;
          
          // Add formulas for each period (including Period 0)
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            
            if (col === 1) {
              // Period 0 - show dash for no revenue
              const dashRange = plSheet.getRange(`${colLetter}${currentRow}`);
              dashRange.values = [['-']];
              dashRange.format.horizontalAlignment = 'Right';
              ExcelFormatter.applyNumberFormat(dashRange);
            } else if (col === 2) {
              // First actual period - reference from Assumptions
              const assumptionRef = this.cellTracker.getCellReference(`revenue_${index}`);
              if (assumptionRef) {
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${assumptionRef}`]];
              } else {
                plSheet.getRange(`${colLetter}${currentRow}`).values = [[item.value || 0]];
              }
              ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
            } else {
              // Growth formula for subsequent periods - reference assumptions sheet
              const prevCol = this.getColumnLetter(col - 1);
              const growthRateRef = this.cellTracker.getCellReference(`revenue_${index}_growth_rate`);
              const periodAdjustment = this.getPeriodAdjustment(modelData.modelPeriods);
              
              console.log(`Revenue ${index} growth ref:`, growthRateRef);
              
              if (growthRateRef) {
                // Extract just the cell reference (e.g., "B24" from "Assumptions!B24")
                const cellRef = growthRateRef.includes('!') ? growthRateRef.split('!')[1] : growthRateRef;
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
                  [[`=${prevCol}${currentRow}*(1+Assumptions!${cellRef}${periodAdjustment})`]];
              } else {
                // No growth rate in assumptions - use flat growth
                console.log(`No growth rate reference found for revenue ${index}`);
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevCol}${currentRow}`]];
              }
              ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
            }
          }
          currentRow++;
        });
      }
      
      // Total Revenue
      plSheet.getRange(`A${currentRow}`).values = [['Total Revenue']];
      const totalRevRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      totalRevRange.format.font.name = 'Times New Roman';
      totalRevRange.format.font.size = 12;
      totalRevRange.format.font.bold = true;
      
      // Add thin underline at top of total
      totalRevRange.format.borders.getItem('EdgeTop').style = 'Continuous';
      totalRevRange.format.borders.getItem('EdgeTop').weight = 'Thin';
      totalRevRange.format.borders.getItem('EdgeTop').color = ExcelFormatter.colors.black;
      
      for (let col = 1; col <= totalColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (col === 1) {
          // Period 0 - show dash
          const dashRange = plSheet.getRange(`${colLetter}${currentRow}`);
          dashRange.values = [['-']];
          dashRange.format.horizontalAlignment = 'Right';
        } else if (modelData.revenueItems && modelData.revenueItems.length > 0) {
          const sumFormula = `=SUM(${colLetter}${revenueStartRow + 3}:${colLetter}${currentRow - 1})`;
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[sumFormula]];
        } else {
          plSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
        }
        ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
      }
      const totalRevenueRow = currentRow;
      currentRow++;
      
      // Add blank row with height 8 between Total Revenue and Cost Items
      plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`).format.rowHeight = 8;
      currentRow++;
      
      // COST ITEMS SECTION
      const opexStartRow = currentRow;
      plSheet.getRange(`A${currentRow}`).values = [['Cost Items']];
      const opexSectionRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      opexSectionRange.format.font.name = 'Times New Roman';
      opexSectionRange.format.font.size = 12;
      opexSectionRange.format.font.bold = true;
      opexSectionRange.format.fill.color = ExcelFormatter.colors.darkBlue;
      opexSectionRange.format.font.color = ExcelFormatter.colors.white;
      
      // Add thin grey underline after section header
      opexSectionRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
      opexSectionRange.format.borders.getItem('EdgeBottom').weight = 'Thin';
      opexSectionRange.format.borders.getItem('EdgeBottom').color = ExcelFormatter.colors.lightGrey;
      currentRow++;
      
      // Add each operating expense
      if (modelData.operatingExpenses && modelData.operatingExpenses.length > 0) {
        modelData.operatingExpenses.forEach((item, index) => {
          plSheet.getRange(`A${currentRow}`).values = [[item.name || `OpEx ${index + 1}`]];
          plSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
          plSheet.getRange(`A${currentRow}`).format.font.size = 12;
          
          // Add formulas for each period (including Period 0)
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            
            if (col === 1) {
              // Period 0 - show dash for no costs
              const dashRange = plSheet.getRange(`${colLetter}${currentRow}`);
              dashRange.values = [['-']];
              dashRange.format.horizontalAlignment = 'Right';
              ExcelFormatter.applyNumberFormat(dashRange);
            } else if (col === 2) {
              // First actual period - negative value from Assumptions
              const assumptionRef = this.cellTracker.getCellReference(`opex_${index}`);
              if (assumptionRef) {
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=-${assumptionRef}`]];
              } else {
                plSheet.getRange(`${colLetter}${currentRow}`).values = [[-item.value || 0]];
              }
              ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
            } else {
              // Growth formula for subsequent periods - reference assumptions sheet
              const prevCol = this.getColumnLetter(col - 1);
              const growthRateRef = this.cellTracker.getCellReference(`opex_${index}_growth_rate`);
              const periodAdjustment = this.getPeriodAdjustment(modelData.modelPeriods);
              
              console.log(`OpEx ${index} growth ref:`, growthRateRef);
              
              if (growthRateRef) {
                // Extract just the cell reference (e.g., "B24" from "Assumptions!B24")
                const cellRef = growthRateRef.includes('!') ? growthRateRef.split('!')[1] : growthRateRef;
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
                  [[`=${prevCol}${currentRow}*(1+Assumptions!${cellRef}${periodAdjustment})`]];
              } else {
                // No growth rate in assumptions - use flat growth
                console.log(`No growth rate reference found for opex ${index}`);
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevCol}${currentRow}`]];
              }
              ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
            }
          }
          currentRow++;
        });
      }
      
      // Total Operating Expenses
      plSheet.getRange(`A${currentRow}`).values = [['Total Operating Expenses']];
      const totalOpExRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      totalOpExRange.format.font.name = 'Times New Roman';
      totalOpExRange.format.font.size = 12;
      totalOpExRange.format.font.bold = true;
      
      // Add thin underline at top of total
      totalOpExRange.format.borders.getItem('EdgeTop').style = 'Continuous';
      totalOpExRange.format.borders.getItem('EdgeTop').weight = 'Thin';
      totalOpExRange.format.borders.getItem('EdgeTop').color = ExcelFormatter.colors.black;
      
      for (let col = 1; col <= totalColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (col === 1) {
          // Period 0 - show dash
          const dashRange = plSheet.getRange(`${colLetter}${currentRow}`);
          dashRange.values = [['-']];
          dashRange.format.horizontalAlignment = 'Right';
        } else if (modelData.operatingExpenses && modelData.operatingExpenses.length > 0) {
          const sumFormula = `=SUM(${colLetter}${opexStartRow + 1}:${colLetter}${currentRow - 1})`;
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[sumFormula]];
        } else {
          plSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
        }
        ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
      }
      const totalOpexRow = currentRow;
      currentRow++;
      
      // NOI - moved directly below Total Operating Expenses
      plSheet.getRange(`A${currentRow}`).values = [['NOI']];
      const noiRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      noiRange.format.font.name = 'Times New Roman';
      noiRange.format.font.size = 12;
      noiRange.format.font.bold = true;
      noiRange.format.fill.color = ExcelFormatter.colors.backgroundDarker5; // Light grey background
      noiRange.format.font.color = ExcelFormatter.colors.black; // Black text
      
      // Add thin underline at top of NOI
      noiRange.format.borders.getItem('EdgeTop').style = 'Continuous';
      noiRange.format.borders.getItem('EdgeTop').weight = 'Thin';
      noiRange.format.borders.getItem('EdgeTop').color = ExcelFormatter.colors.black;
      
      for (let col = 1; col <= totalColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (col === 1) {
          // Period 0 - show dash
          const dashRange = plSheet.getRange(`${colLetter}${currentRow}`);
          dashRange.values = [['-']];
          dashRange.format.horizontalAlignment = 'Right';
        } else {
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
            [[`=${colLetter}${totalRevenueRow}+${colLetter}${totalOpexRow}`]];
        }
        ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
      }
      const ebitdaRow = currentRow;
      currentRow += 2;
      
      // Real estate model: NOI is the final metric - no interest, tax, or net income calculations needed
      
      // Track NOI for FCF calculations
      this.plCellTracker.recordCell('noi', 'P&L Statement', `B${ebitdaRow}:${this.getColumnLetter(totalColumns)}${ebitdaRow}`);
      
      // Apply number formatting with brackets for negatives and dash for zeros
      const dataRange = plSheet.getRange(`B5:${this.getColumnLetter(totalColumns)}${ebitdaRow}`);
      ExcelFormatter.applyNumberFormat(dataRange);
      
      // Apply Times New Roman font to all data cells
      const allDataRange = plSheet.getRange(`A1:${this.getColumnLetter(totalColumns)}${ebitdaRow}`);
      allDataRange.format.font.name = 'Times New Roman';
      allDataRange.format.font.size = 12;
      
      // Hide gridlines
      await context.sync();
      plSheet.showGridlines = false;
      
      // Auto-fit columns
      plSheet.getRange(`A:${this.getColumnLetter(periodColumns)}`).format.autofitColumns();
      
      await context.sync();
      console.log('✅ P&L Statement created successfully with formatting');
      
      // Print tracked P&L cells for debugging
      console.log('📊 P&L Cell References:');
      this.plCellTracker.printAllCells();
    });
  }
  
  // Read the actual P&L sheet to discover structure
  async readPLSheetStructure() {
    return Excel.run(async (context) => {
      console.log('🔍 Reading P&L sheet structure...');
      
      try {
        const plSheet = context.workbook.worksheets.getItem('P&L Statement');
        
        // Get the used range to understand the structure
        const usedRange = plSheet.getUsedRange();
        usedRange.load(['values', 'formulas', 'address']);
        
        await context.sync();
        
        const values = usedRange.values;
        const formulas = usedRange.formulas;
        const address = usedRange.address;
        
        console.log('📋 P&L Used Range:', address);
        console.log('📋 P&L Values sample:', values.slice(0, 5));
        
        // Parse the structure to find key line items
        const structure = {
          totalColumns: values[0].length,
          lineItems: {},
          periodColumns: [],
          address: address
        };
        
        // Find period headers (typically row 3)
        if (values.length > 2) {
          structure.periodColumns = values[2].slice(1); // Skip first column (labels)
        }
        
        // Scan for key line items
        values.forEach((row, rowIndex) => {
          const label = row[0];
          if (typeof label === 'string') {
            const labelLower = label.toLowerCase();
            
            // Map key items to their row positions
            if (labelLower.includes('total revenue')) {
              structure.lineItems.totalRevenue = {
                row: rowIndex + 1, // Excel rows are 1-based
                range: `B${rowIndex + 1}:${this.getColumnLetter(structure.totalColumns - 1)}${rowIndex + 1}`,
                label: label
              };
            }
            if (labelLower.includes('total operating expenses')) {
              structure.lineItems.totalOpex = {
                row: rowIndex + 1,
                range: `B${rowIndex + 1}:${this.getColumnLetter(structure.totalColumns - 1)}${rowIndex + 1}`,
                label: label
              };
            }
            if (labelLower.includes('ebitda')) {
              structure.lineItems.ebitda = {
                row: rowIndex + 1,
                range: `B${rowIndex + 1}:${this.getColumnLetter(structure.totalColumns - 1)}${rowIndex + 1}`,
                label: label
              };
            }
            if (labelLower.includes('interest expense')) {
              structure.lineItems.interestExpense = {
                row: rowIndex + 1,
                range: `B${rowIndex + 1}:${this.getColumnLetter(structure.totalColumns - 1)}${rowIndex + 1}`,
                label: label
              };
            }
            if (labelLower.includes('net income')) {
              structure.lineItems.netIncome = {
                row: rowIndex + 1,
                range: `B${rowIndex + 1}:${this.getColumnLetter(structure.totalColumns - 1)}${rowIndex + 1}`,
                label: label
              };
            }
            
            // Also capture individual revenue and expense items
            if (rowIndex > 4 && rowIndex < 50 && label && label.length > 0 && !label.includes('REVENUE') && !label.includes('EXPENSES') && !label.includes('Total')) {
              // This might be an individual line item
              if (!structure.lineItems.individualItems) {
                structure.lineItems.individualItems = [];
              }
              structure.lineItems.individualItems.push({
                row: rowIndex + 1,
                range: `B${rowIndex + 1}:${this.getColumnLetter(structure.totalColumns - 1)}${rowIndex + 1}`,
                label: label,
                type: 'unknown'
              });
            }
          }
        });
        
        console.log('📋 P&L Structure parsed:', structure.lineItems);
        return structure;
        
      } catch (error) {
        console.error('❌ Error reading P&L sheet:', error);
        return { error: 'Could not read P&L sheet' };
      }
    });
  }
  
  // Read the actual Assumption sheet to discover structure
  async readAssumptionSheetStructure() {
    return Excel.run(async (context) => {
      console.log('🔍 Reading Assumption sheet structure...');
      
      try {
        const assumptionSheet = context.workbook.worksheets.getItem('Assumptions');
        
        // Get the used range
        const usedRange = assumptionSheet.getUsedRange();
        usedRange.load(['values', 'address']);
        
        await context.sync();
        
        const values = usedRange.values;
        const address = usedRange.address;
        
        console.log('📋 Assumptions Used Range:', address);
        
        const structure = {
          address: address,
          keyItems: {}
        };
        
        // Parse assumption values
        values.forEach((row, rowIndex) => {
          const label = row[0];
          const value = row[1];
          
          if (typeof label === 'string' && label.length > 0) {
            const labelLower = label.toLowerCase();
            const cellRef = `B${rowIndex + 1}`;
            
            if (labelLower.includes('deal value')) {
              structure.keyItems.dealValue = { cell: cellRef, value: value, label: label };
            }
            if (labelLower.includes('debt financing')) {
              structure.keyItems.debtFinancing = { cell: cellRef, value: value, label: label };
            }
            if (labelLower.includes('terminal cap rate')) {
              structure.keyItems.terminalCapRate = { cell: cellRef, value: value, label: label };
            }
            if (labelLower.includes('fixed interest rate') || labelLower.includes('total interest rate')) {
              structure.keyItems.interestRate = { cell: cellRef, value: value, label: label };
            }
          }
        });
        
        console.log('📋 Assumption Structure parsed:', structure.keyItems);
        return structure;
        
      } catch (error) {
        console.error('❌ Error reading Assumption sheet:', error);
        return { error: 'Could not read Assumption sheet' };
      }
    });
  }
  
  // Generate FCF prompt with REAL cell references from actual sheets
  // DUPLICATE METHOD REMOVED - See line 2476 for the correct implementation
  generateRealFCFPrompt_REMOVED_DISABLED(modelData, plStructure, assumptionStructure) {
    throw new Error('This method has been disabled. Use the method at line 2476 instead.');
    return 'DISABLED METHOD';
    console.log('🤖 Generating FCF AI prompt...');
    
    // Calculate periods
    const periods = this.calculatePeriods(modelData.projectStartDate, modelData.projectEndDate, modelData.modelPeriods);
    const maxPeriods = Math.min(periods, 60);
    
    const prompt = `You are a world-class M&A financial modeling expert. You have been provided with a complete Assumptions sheet and a fully generated P&L Statement. Your task is to create a comprehensive Free Cash Flow Statement that references these existing sheets.

**DEAL OVERVIEW:**
- Deal: ${modelData.dealName}
- Currency: ${modelData.currency} 
- Period Type: ${modelData.modelPeriods}
- Duration: ${modelData.projectStartDate} to ${modelData.projectEndDate}
- Periods Needed: ${maxPeriods}

**AVAILABLE ASSUMPTIONS SHEET DATA:**
${this.formatDetailedAssumptions(modelData)}

**AVAILABLE P&L STATEMENT DATA:**
${this.formatDetailedPLReferences()}

**YOUR TASK:**
Create a complete Free Cash Flow Statement that uses ONLY Excel formulas referencing the above cell locations. Build a comprehensive FCF model with the following sections:

**1. OPERATING CASH FLOW:**
- Start with NOI from P&L: ${this.plCellTracker.getCellReference('noi')}
- Add back: Depreciation & Amortization (create reasonable assumption)
- Less: Working Capital Changes (typical 2-5% of revenue change)
- Less: Tax adjustments if needed

**2. INVESTING CASH FLOW:**
- Initial Investments: Use CapEx from assumptions
- Asset sales/disposals (if applicable)

**3. FINANCING CASH FLOW:**
- Interest payments: Already in P&L at ${this.plCellTracker.getCellReference('interest_expense')}
- Principal payments: Calculate debt amortization schedule
- Dividend payments (if any)

**4. FREE CASH FLOW METRICS:**
- Unlevered FCF (before financing)
- Levered FCF (after debt service)
- Cumulative FCF
- IRR calculation using XIRR
- Terminal Value using cap rate: ${this.cellTracker.getCellReference('terminalCapRate')}

**CRITICAL REQUIREMENTS:**

1. **FORMULAS ONLY**: Every single cell must contain an Excel formula, never hardcoded values
2. **EXACT REFERENCES**: Use the exact cell references provided above
3. **FORMAT**: 'P&L Statement'!B15 or 'Assumptions'!C10
4. **WORKING CAPITAL**: =((Current_Revenue*WC_%) - (Previous_Revenue*WC_%)) as negative cash flow
5. **DEPRECIATION**: =Total_CapEx/Useful_Life or reasonable percentage
6. **DEBT SERVICE**: =Principal_Payment + Interest_Payment
7. **PERIOD ADJUSTMENT**: Adjust annual rates for ${modelData.modelPeriods} periods

**EXPECTED OUTPUT FORMAT:**
Provide the complete Excel structure:
- Row-by-row layout with exact cell addresses
- Exact formulas for each cell
- All ${maxPeriods} period columns
- Proper section headers and formatting
- Terminal value and valuation in final periods

**EXAMPLE FORMULA STYLE:**
Row 15: NOI from P&L
- B15: ='P&L Statement'!B12
- C15: ='P&L Statement'!C12
- D15: ='P&L Statement'!D12

Row 16: Working Capital Change
- B16: =-('P&L Statement'!B6*0.03)
- C16: =-(('P&L Statement'!C6*0.03)-('P&L Statement'!B6*0.03))
- D16: =-(('P&L Statement'!D6*0.03)-('P&L Statement'!C6*0.03))

Provide the COMPLETE Free Cash Flow model with exact Excel formulas for every cell across all periods.`;

    console.log('📝 Generated REAL FCF AI prompt with actual P&L cell references');
    return prompt;
  }
  
  // Format detailed assumptions for AI prompt
  formatDetailedAssumptions(modelData) {
    let output = 'EXACT ASSUMPTIONS SHEET CELL REFERENCES (use these exact references):\n\n';
    
    output += `**DEAL STRUCTURE:**\n`;
    output += `- Deal Value: ${this.cellTracker.getCellReference('dealValue')}\n`;
    output += `- Deal LTV (%): ${this.cellTracker.getCellReference('dealLTV')}\n`;
    output += `- Transaction Fee (%): ${this.cellTracker.getCellReference('transactionFee')}\n`;
    output += `- Equity Contribution: ${this.cellTracker.getCellReference('equityContribution')}\n`;
    output += `- Debt Financing: ${this.cellTracker.getCellReference('debtFinancing')}\n\n`;
    
    output += `**PROJECT PARAMETERS:**\n`;
    output += `- Currency: ${this.cellTracker.getCellReference('currency')}\n`;
    output += `- Project Start: ${this.cellTracker.getCellReference('projectStartDate')}\n`;
    output += `- Project End: ${this.cellTracker.getCellReference('projectEndDate')}\n`;
    output += `- Model Periods: ${this.cellTracker.getCellReference('modelPeriods')}\n\n`;
    
    output += `**DEBT MODEL:**\n`;
    const hasDebt = modelData.dealLTV && parseFloat(modelData.dealLTV) > 0;
    if (hasDebt) {
      output += `- Fixed Interest Rate (%): ${this.cellTracker.getCellReference('fixedRate')}\n`;
      output += `- Loan Issuance Fees (%): ${this.cellTracker.getCellReference('loanIssuanceFees')}\n`;
    } else {
      output += `- No debt financing (LTV = 0%)\n`;
    }
    output += '\n';
    
    output += `**REVENUE ASSUMPTIONS:**\n`;
    if (modelData.revenueItems && modelData.revenueItems.length > 0) {
      modelData.revenueItems.forEach((item, index) => {
        output += `- ${item.name}: ${this.cellTracker.getCellReference(`revenue_${index}`)}\n`;
        output += `  * Growth Type: ${item.growthType || 'none'}\n`;
        if (item.growthType === 'annual' && item.annualGrowthRate) {
          output += `  * Annual Growth Rate: ${item.annualGrowthRate}%\n`;
        }
      });
    } else {
      output += `- No revenue items specified\n`;
    }
    output += '\n';
    
    output += `**OPERATING EXPENSE ASSUMPTIONS:**\n`;
    if (modelData.operatingExpenses && modelData.operatingExpenses.length > 0) {
      modelData.operatingExpenses.forEach((item, index) => {
        output += `- ${item.name}: ${this.cellTracker.getCellReference(`opex_${index}`)}\n`;
        output += `  * Growth Type: ${item.growthType || 'none'}\n`;
        if (item.growthType === 'annual' && item.annualGrowthRate) {
          output += `  * Annual Growth Rate: ${item.annualGrowthRate}%\n`;
        }
      });
    } else {
      output += `- No operating expense items specified\n`;
    }
    output += '\n';
    
    output += `**INITIAL INVESTMENT ASSUMPTIONS:**\n`;
    if (modelData.capitalExpenses && modelData.capitalExpenses.length > 0) {
      modelData.capitalExpenses.forEach((item, index) => {
        output += `- ${item.name}: ${this.cellTracker.getCellReference(`capex_${index}`)}\n`;
        output += `  * Growth Type: ${item.growthType || 'none'}\n`;
        if (item.growthType === 'annual' && item.annualGrowthRate) {
          output += `  * Annual Growth Rate: ${item.annualGrowthRate}%\n`;
        }
      });
    } else {
      output += `- No initial investment items specified\n`;
    }
    output += '\n';
    
    output += `**EXIT ASSUMPTIONS:**\n`;
    output += `- Disposal Cost (%): ${this.cellTracker.getCellReference('disposalCost')}\n`;
    output += `- Terminal Cap Rate (%): ${this.cellTracker.getCellReference('terminalCapRate')}\n\n`;
    
    output += `**USAGE INSTRUCTIONS:**\n`;
    output += `- Reference format: ='Assumptions'!B15 (where B15 is the specific cell)\n`;
    output += `- Always use single quotes around sheet names\n`;
    output += `- These are the ONLY assumption values available - do not create new ones\n`;
    
    return output;
  }
  
  // Format detailed P&L references for AI prompt
  formatDetailedPLReferences() {
    let output = 'EXACT P&L STATEMENT CELL REFERENCES (use these exact references):\n\n';
    
    output += `**REVENUE SECTION:**\n`;
    if (modelData.revenueItems) {
      modelData.revenueItems.forEach((item, index) => {
        const ref = this.plCellTracker.getCellReference(`revenue_item_${index}`);
        if (ref) {
          output += `- ${item.name}: ${ref}\n`;
        }
      });
    }
    output += `- TOTAL REVENUE: ${this.plCellTracker.getCellReference('total_revenue')}\n\n`;
    
    output += `**EXPENSE SECTION:**\n`;
    if (modelData.operatingExpenses) {
      modelData.operatingExpenses.forEach((item, index) => {
        const ref = this.plCellTracker.getCellReference(`opex_item_${index}`);
        if (ref) {
          output += `- ${item.name}: ${ref}\n`;
        }
      });
    }
    output += `- TOTAL OPERATING EXPENSES: ${this.plCellTracker.getCellReference('total_opex')}\n\n`;
    
    output += `**KEY METRICS:**\n`;
    output += `- NOI: ${this.plCellTracker.getCellReference('noi')}\n`;
    output += `- INTEREST EXPENSE: ${this.plCellTracker.getCellReference('interest_expense')}\n`;
    output += `- NET INCOME: ${this.plCellTracker.getCellReference('net_income')}\n\n`;
    
    output += `**USAGE INSTRUCTIONS:**\n`;
    output += `- Reference format: ='P&L Statement'!B15 (where B15 is the specific cell)\n`;
    output += `- For ranges: Use the entire range like B6:AK6 for all periods of a line item\n`;
    output += `- Always use single quotes around sheet names with spaces\n`;
    
    return output;
  }
  
  // Format CapEx references
  formatCapexReferences(modelData) {
    let output = '';
    if (modelData.capitalExpenses && modelData.capitalExpenses.length > 0) {
      modelData.capitalExpenses.forEach((item, index) => {
        output += `${item.name}: ${this.cellTracker.getCellReference(`capex_${index}`)}\n`;
      });
    } else {
      output = 'No initial investments specified.';
    }
    return output;
  }
  
  // Create professional FCF sheet using REAL cell references from discovered P&L structure
  // DUPLICATE METHOD REMOVED - See line 3738 for the correct implementation
  async createAIFCFSheet_REMOVED_DISABLED(modelData, aiPrompt, plStructure, assumptionStructure) {
    throw new Error('This method has been disabled. Use the method at line 3738 instead.');
    return { success: false, error: 'DISABLED METHOD' };
    return Excel.run(async (context) => {
      console.log('💰 Creating professional Free Cash Flow sheet...');
      
      const sheets = context.workbook.worksheets;
      
      // Delete existing FCF sheet if it exists
      try {
        const existingSheet = sheets.getItemOrNullObject('Free Cash Flow');
        existingSheet.load('name');
        await context.sync();
        
        if (!existingSheet.isNullObject) {
          console.log('🗑️ Deleting existing FCF sheet');
          existingSheet.delete();
          await context.sync();
        }
      } catch (e) {
        // Sheet doesn't exist, continue
      }
      
      // Create new FCF sheet
      const fcfSheet = sheets.add('Free Cash Flow');
      await context.sync();
      
      // Calculate periods
      const periods = this.calculatePeriods(modelData.projectStartDate, modelData.projectEndDate, modelData.modelPeriods);
      const periodColumns = periods; // Use full calculated periods
      
      let currentRow = 1;
      
      // TITLE
      fcfSheet.getRange('A1').values = [[`Free Cash Flow Statement (${modelData.currency})`]];
      fcfSheet.getRange('A1').format.font.bold = true;
      fcfSheet.getRange('A1').format.font.size = 16;
      fcfSheet.getRange('A1').format.fill.color = '#1f4e79';
      fcfSheet.getRange('A1').format.font.color = 'white';
      currentRow = 3;
      
      // TIME PERIOD HEADERS - Include Period 0 for Initial Investment
      const headers = [''];
      const prevPeriodLabel = this.getPreviousPeriodLabel(modelData.projectStartDate, modelData.modelPeriods);
      headers.push(prevPeriodLabel); // Period 0 with actual date label
      const startDate = new Date(modelData.projectStartDate);
      for (let i = 0; i < periodColumns; i++) {
        headers.push(this.formatPeriodHeader(startDate, i, modelData.modelPeriods));
      }
      const totalColumns = periodColumns + 1; // +1 for Initial Investment period
      
      const headerRange = fcfSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = '#d9d9d9';
      headerRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
      headerRange.format.borders.getItem('EdgeBottom').weight = 'Medium';
      currentRow += 2;
      
      // OPERATING CASH FLOW SECTION
      fcfSheet.getRange(`A${currentRow}`).values = [['OPERATING CASH FLOW']];
      fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
      fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#b7dee8';
      currentRow++;
      
      // Use REAL P&L references discovered from the actual sheet
      if (plStructure && plStructure.sheetExists) {
        
        // Track FCF line item positions for later reference
        const fcfStructure = {};
        
        // NOI (from actual P&L)
        fcfSheet.getRange(`A${currentRow}`).values = [['NOI']];
        fcfStructure.noi = currentRow;
        if (plStructure.lineItems.noi) {
          const noiRow = plStructure.lineItems.noi.row;
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            if (col === 1) {
              // Period 0 in FCF has no NOI (it's the initial investment period)
              fcfSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
            } else {
              // For operating periods, FCF column maps to P&L column minus 1
              // FCF Period 1 (column 2) = P&L Period 1 (column 1)
              const plCol = this.getColumnLetter(col - 1);
              fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`='P&L Statement'!${plCol}${noiRow}`]];
            }
          }
        }
        currentRow++;
        
        // Real estate model: No tax calculations required
        
        // NOI (No tax calculations for real estate)
        fcfSheet.getRange(`A${currentRow}`).values = [['NOI (Net Operating Income)']];
        fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
        fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#daeef3';
        fcfStructure.nopat = currentRow;
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${fcfStructure.noi}`]];
        }
        currentRow += 2;
        
        // WORKING CAPITAL SECTION
        fcfSheet.getRange(`A${currentRow}`).values = [['WORKING CAPITAL ADJUSTMENTS']];
        fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
        fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#fde9d9';
        currentRow++;
        
        // Working Capital Change using REAL Total Revenue reference
        fcfSheet.getRange(`A${currentRow}`).values = [['Less: Change in Working Capital (3% of Revenue)']];
        fcfStructure.workingCapital = currentRow;
        if (plStructure.lineItems.totalRevenue) {
          const revenueRow = plStructure.lineItems.totalRevenue.row;
          
          // Handle Period 0 (initial period) - Column B is Initial Investment
          const period0Col = this.getColumnLetter(2);
          fcfSheet.getRange(`${period0Col}${currentRow}`).values = [[0]]; // No working capital change in Period 0
          
          for (let col = 3; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            const plCol = this.getColumnLetter(col); // Direct mapping: FCF column C -> P&L column C, etc.
            const prevCol = this.getColumnLetter(col - 1);   // Previous P&L column
            
            if (col === 3) {
              // First operational period - initial working capital investment (3% of revenue)
              fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=-'P&L Statement'!${plCol}${revenueRow}*0.03`]];
            } else {
              // Subsequent periods - change in working capital from previous period
              fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=-('P&L Statement'!${plCol}${revenueRow}*0.03-'P&L Statement'!${prevCol}${revenueRow}*0.03)`]];
            }
          }
        }
        currentRow += 2;
        
        // INITIAL INVESTMENTS SECTION
        fcfSheet.getRange(`A${currentRow}`).values = [['INITIAL INVESTMENTS']];
        fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
        fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#f2f2f2';
        currentRow++;
        
        // Initial Investments from assumptions
        fcfSheet.getRange(`A${currentRow}`).values = [['Less: Initial Investments']];
        fcfStructure.capex = currentRow;
        if (assumptionStructure && assumptionStructure.assumptions.capitalExpenses && assumptionStructure.assumptions.capitalExpenses.length > 0) {
          // Use actual capital expense references from assumptions
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            let capexFormula = '';
            assumptionStructure.assumptions.capitalExpenses.forEach((item, index) => {
              if (index === 0) {
                capexFormula = `-Assumptions!${item.cellRef}`;
              } else {
                capexFormula += `-Assumptions!${item.cellRef}`;
              }
            });
            fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[capexFormula]];
          }
        } else {
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            fcfSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
          }
        }
        currentRow += 2;
        
        // UNLEVERED FREE CASH FLOW
        fcfSheet.getRange(`A${currentRow}`).values = [['Unlevered Free Cash Flow']];
        fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
        fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#c5e0b4';
        fcfSheet.getRange(`A${currentRow}`).format.borders.getItem('EdgeTop').style = 'Continuous';
        fcfSheet.getRange(`A${currentRow}`).format.borders.getItem('EdgeTop').weight = 'Medium';
        fcfStructure.unleveredFCF = currentRow;
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${fcfStructure.nopat}+${colLetter}${fcfStructure.workingCapital}+${colLetter}${fcfStructure.capex}`]];
          fcfSheet.getRange(`${colLetter}${currentRow}`).format.borders.getItem('EdgeTop').style = 'Continuous';
          fcfSheet.getRange(`${colLetter}${currentRow}`).format.borders.getItem('EdgeTop').weight = 'Medium';
        }
        currentRow += 2;
        
        // FINANCING CASH FLOWS SECTION
        fcfSheet.getRange(`A${currentRow}`).values = [['FINANCING CASH FLOWS']];
        fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
        fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#ffc7ce';
        currentRow++;
        
        // Debt Expense - referenced from Debt Model sheet
        fcfSheet.getRange(`A${currentRow}`).values = [['Less: Debt Expense']];
        fcfStructure.interestPayments = currentRow;
        // Reference debt expense directly from Debt Model sheet row 6
        const hasDebt = modelData.dealLTV && parseFloat(modelData.dealLTV) > 0;
        if (hasDebt) {
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            let debtExpenseFormula = '0';
            
            if (col === 1) {
              // Period 0 - no debt expense
              debtExpenseFormula = '0';
            } else {
              // Reference debt expense from Debt Model: Period 1 is in column C
              const debtModelCol = this.getColumnLetter(col + 1); // FCF col 1 -> Debt Model col C (col+1)
              debtExpenseFormula = `=-'Debt Model'!${debtModelCol}8`;
            }
            
            fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[debtExpenseFormula]];
          }
        } else {
          // No debt expense
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            fcfSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
          }
        }
        currentRow++;
        
        // Loan proceeds (positive cash inflow in Period 0)
        fcfSheet.getRange(`A${currentRow}`).values = [['Loan proceeds']];
        fcfStructure.loanProceeds = currentRow;
        if (hasDebt) {
          const debtRef = this.cellTracker.getCellReference('debtFinancing');
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            if (col === 2) { // First operating period (Period 1)
              // Loan proceeds come in during first operating period
              if (debtRef) {
                fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${debtRef}`]];
              } else {
                fcfSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
              }
            } else {
              // All other periods - no loan proceeds
              fcfSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
            }
          }
        } else {
          // No debt - all zeros
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            fcfSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
          }
        }
        currentRow += 2;
        
        // ASSET DISPOSAL PROCEEDS (Final Period Only)
        fcfSheet.getRange(`A${currentRow}`).values = [['Asset Disposal Proceeds']];
        fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
        fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#d5e8d4';
        fcfStructure.assetDisposal = currentRow;
        
        // Calculate disposal proceeds only in final period using general disposal cost %
        for (let col = 1; col <= totalColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          if (col === totalColumns) {
            // Final period: Deal Value * (1 - Disposal Cost %)
            const dealValueRef = this.cellTracker.getCellReference('dealValue');
            const disposalCostRef = this.cellTracker.getCellReference('disposalCost');
            
            if (dealValueRef && disposalCostRef) {
              // Net disposal proceeds = Deal Value * (1 - Disposal Cost %)
              fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${dealValueRef}*(1-${disposalCostRef})`]];
            } else if (modelData.dealValue && modelData.disposalCost) {
              // Fallback: use direct values if cell references not available
              const netProceeds = modelData.dealValue * (1 - (modelData.disposalCost / 100));
              fcfSheet.getRange(`${colLetter}${currentRow}`).values = [[netProceeds]];
            } else {
              fcfSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
            }
          } else {
            // Not final period: No disposal
            fcfSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
          }
        }
        currentRow += 2;
        
        // LEVERED FREE CASH FLOW  
        fcfSheet.getRange(`A${currentRow}`).values = [['Levered Free Cash Flow']];
        fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
        fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#ffeb9c';
        fcfSheet.getRange(`A${currentRow}`).format.borders.getItem('EdgeTop').style = 'Double';
        fcfSheet.getRange(`A${currentRow}`).format.borders.getItem('EdgeTop').weight = 'Thick';
        fcfStructure.leveredFCF = currentRow;
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          // Include loan proceeds in levered cash flow calculation
          fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${fcfStructure.unleveredFCF}+${colLetter}${fcfStructure.interestPayments}+${colLetter}${fcfStructure.loanProceeds}+${colLetter}${fcfStructure.assetDisposal}`]];
          fcfSheet.getRange(`${colLetter}${currentRow}`).format.borders.getItem('EdgeTop').style = 'Double';
          fcfSheet.getRange(`${colLetter}${currentRow}`).format.borders.getItem('EdgeTop').weight = 'Thick';
        }
        currentRow += 2;
        
        // CUMULATIVE METRICS SECTION
        fcfSheet.getRange(`A${currentRow}`).values = [['CUMULATIVE ANALYSIS']];
        fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
        fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#e2efda';
        currentRow++;
        
        // Cumulative FCF using actual tracked levered FCF row
        fcfSheet.getRange(`A${currentRow}`).values = [['Cumulative Free Cash Flow']];
        fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          if (col === 1) {
            fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${fcfStructure.leveredFCF}`]];
          } else {
            const prevCol = this.getColumnLetter(col - 1);
            fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevCol}${currentRow}+${colLetter}${fcfStructure.leveredFCF}`]];
          }
        }
        currentRow += 2;
        
        // NPV CALCULATIONS SECTION
        fcfSheet.getRange(`A${currentRow}`).values = [['NPV CALCULATIONS']];
        fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
        fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#d5e4bc';
        currentRow++;
        
        // Undiscounted NPV (simple sum) using actual levered FCF row
        fcfSheet.getRange(`A${currentRow}`).values = [['Undiscounted NPV (Sum of FCF)']];
        fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
        const undiscountedRange = `B${fcfStructure.leveredFCF}:${this.getColumnLetter(periodColumns)}${fcfStructure.leveredFCF}`;
        fcfSheet.getRange(`B${currentRow}`).formulas = [[`=SUM(${undiscountedRange})`]];
        currentRow++;
        
        // Discounted NPV using WACC from assumptions
        fcfSheet.getRange(`A${currentRow}`).values = [['Discounted NPV @ WACC']];
        fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
        fcfSheet.getRange(`A${currentRow}`).format.font.italic = true;
        if (assumptionStructure && assumptionStructure.assumptions.discountRate) {
          const waccCellRef = assumptionStructure.assumptions.discountRate.cellRef;
          fcfSheet.getRange(`B${currentRow}`).formulas = [[`=NPV(Assumptions!${waccCellRef},${undiscountedRange})`]];
        } else {
          // Fallback to 10% if WACC not found
          fcfSheet.getRange(`B${currentRow}`).formulas = [[`=NPV(0.1,${undiscountedRange})`]];
        }
        currentRow++;
        
        // IRR CALCULATION using Excel's built-in IRR function
        fcfSheet.getRange(`A${currentRow}`).values = [['Internal Rate of Return (IRR)']];
        fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
        fcfSheet.getRange(`A${currentRow}`).format.font.italic = true;
        
        // Create cash flow series starting with negative initial investment
        // Initial investment is negative (cash outflow), followed by positive FCF (cash inflows)
        let irrCashFlowRange = '';
        if (modelData.dealValue) {
          // Use the equity contribution as the initial investment (negative)
          const equityContribution = modelData.dealValue * ((100 - (modelData.dealLTV || 70)) / 100);
          
          // Create a helper row for the IRR cash flow series
          const irrCashFlowRow = currentRow + 1;
          fcfSheet.getRange(`A${irrCashFlowRow}`).values = [['IRR Cash Flow Series:']];
          fcfSheet.getRange(`A${irrCashFlowRow}`).format.font.size = 10;
          fcfSheet.getRange(`A${irrCashFlowRow}`).format.font.italic = true;
          
          // Period 0: Negative initial investment (equity contribution)
          fcfSheet.getRange(`B${irrCashFlowRow}`).values = [[-equityContribution]];
          
          // Periods 1+: Reference the levered FCF values plus disposal proceeds in final period
          for (let col = 2; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            if (col === totalColumns) {
              // Final period: Levered FCF + Asset Disposal Proceeds
              fcfSheet.getRange(`${colLetter}${irrCashFlowRow}`).formulas = [[`=${colLetter}${fcfStructure.leveredFCF}+${colLetter}${fcfStructure.assetDisposal}`]];
            } else {
              // Regular periods: Just the levered FCF
              fcfSheet.getRange(`${colLetter}${irrCashFlowRow}`).formulas = [[`=${colLetter}${fcfStructure.leveredFCF}`]];
            }
          }
          
          // IRR calculation using the cash flow series
          irrCashFlowRange = `B${irrCashFlowRow}:${this.getColumnLetter(totalColumns)}${irrCashFlowRow}`;
          fcfSheet.getRange(`B${currentRow}`).formulas = [[`=IFERROR(IRR(${irrCashFlowRange}),"No Solution")`]];
          fcfSheet.getRange(`B${currentRow}`).format.numberFormat = [['0.00%']];
          
          currentRow++; // Skip the hidden cash flow row
          currentRow++; // Move to next row for MOIC
          
          // MOIC (Multiple of Invested Capital) calculation
          fcfSheet.getRange(`A${currentRow}`).values = [['Multiple of Invested Capital (MOIC)']];
          fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
          fcfSheet.getRange(`A${currentRow}`).format.font.italic = true;
          
          // MOIC = Total Cash Returned / Initial Investment
          // Use the same cash flow range as IRR (excluding initial investment) and sum it
          const cashFlowRangeWithoutInitial = `C${irrCashFlowRow}:${this.getColumnLetter(totalColumns)}${irrCashFlowRow}`;
          fcfSheet.getRange(`B${currentRow}`).formulas = [[`=SUM(${cashFlowRangeWithoutInitial}) / ${equityContribution}`]];
          fcfSheet.getRange(`B${currentRow}`).format.numberFormat = [['0.00"x"']];
          
        } else {
          // Fallback if deal value not available
          fcfSheet.getRange(`B${currentRow}`).values = [['Deal value required for IRR calculation']];
          fcfSheet.getRange(`B${currentRow}`).format.font.italic = true;
          currentRow++;
          
          fcfSheet.getRange(`A${currentRow}`).values = [['Multiple of Invested Capital (MOIC)']];
          fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
          fcfSheet.getRange(`A${currentRow}`).format.font.italic = true;
          fcfSheet.getRange(`B${currentRow}`).values = [['Deal value required for MOIC calculation']];
          fcfSheet.getRange(`B${currentRow}`).format.font.italic = true;
        }
        
      } else {
        // Fallback if P&L structure not found
        fcfSheet.getRange(`A${currentRow}`).values = [['⚠️ P&L Structure Not Found - Manual Input Required']];
        fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#ffcccc';
      }
      
      // Format all numbers without currency symbols, with red brackets for negatives
      const dataRange = fcfSheet.getRange(`B5:${this.getColumnLetter(periodColumns)}${currentRow}`);
      const numberFormat = '#,##0;[Red]-#,##0;"-"'; // Positive, negative with minus sign, zero as dash
      dataRange.numberFormat = [[numberFormat]];
      
      // Auto-fit columns
      fcfSheet.getRange(`A:${this.getColumnLetter(periodColumns)}`).format.autofitColumns();
      
      await context.sync();
      
      console.log('✅ Professional FCF sheet created with REAL P&L references');
      console.log('📊 AI Prompt available in console for future OpenAI integration:');
      console.log('='.repeat(80));
      console.log(aiPrompt);
      console.log('='.repeat(80));
    });
  }
  
  // Helper to get value for a key from modelData
  getValueForKey(key, modelData) {
    if (key.startsWith('revenue_')) {
      const index = parseInt(key.split('_')[1]);
      if (!isNaN(index) && modelData.revenueItems?.[index]) {
        return modelData.revenueItems[index].value || 0;
      }
    }
    // Add similar logic for other keys as needed
    return modelData[key] || '';
  }

  // Format revenue items for prompt
  formatRevenueItems(modelData) {
    if (!modelData.revenueItems || modelData.revenueItems.length === 0) return 'None';
    
    return modelData.revenueItems.map((item, index) => {
      const valueRef = this.cellTracker.getCellReference(`revenue_${index}`);
      const growthTypeRef = this.cellTracker.getCellReference(`revenue_${index}_growth_type`);
      const growthRateRef = this.cellTracker.getCellReference(`revenue_${index}_growth_rate`);
      
      return `- ${item.name}: ${item.value} (Cell: ${valueRef})
  Growth Type: ${item.growthType} (Cell: ${growthTypeRef})
  Growth Rate: ${item.annualGrowthRate || 0}% (Cell: ${growthRateRef})`;
    }).join('\n');
  }

  // Format operating expenses for prompt
  formatOpexItems(modelData) {
    if (!modelData.operatingExpenses || modelData.operatingExpenses.length === 0) return 'None';
    
    return modelData.operatingExpenses.map((item, index) => {
      const valueRef = this.cellTracker.getCellReference(`opex_${index}`);
      const growthTypeRef = this.cellTracker.getCellReference(`opex_${index}_growth_type`);
      const growthRateRef = this.cellTracker.getCellReference(`opex_${index}_growth_rate`);
      
      return `- ${item.name}: ${item.value} (Cell: ${valueRef})
  Growth Type: ${item.growthType} (Cell: ${growthTypeRef})
  Growth Rate: ${item.annualGrowthRate || 0}% (Cell: ${growthRateRef})`;
    }).join('\n');
  }

  // Format capital expenses for prompt
  formatCapexItems(modelData) {
    if (!modelData.capitalExpenses || modelData.capitalExpenses.length === 0) return 'None';
    
    return modelData.capitalExpenses.map((item, index) => {
      const valueRef = this.cellTracker.getCellReference(`capex_${index}`);
      const growthTypeRef = this.cellTracker.getCellReference(`capex_${index}_growth_type`);
      const growthRateRef = this.cellTracker.getCellReference(`capex_${index}_growth_rate`);
      
      return `- ${item.name}: ${item.value} (Cell: ${valueRef})
  Growth Type: ${item.growthType} (Cell: ${growthTypeRef})
  Growth Rate: ${item.annualGrowthRate || 0}% (Cell: ${growthRateRef})`;
    }).join('\n');
  }

  // Format debt model for prompt
  formatDebtModel(modelData) {
    const hasDebt = modelData.dealLTV && parseFloat(modelData.dealLTV) > 0;
    if (!hasDebt) return 'No debt financing (LTV = 0)';
    
    let debtInfo = `- Loan Issuance Fees: ${modelData.debtSettings?.loanIssuanceFees || 1.5}% (Cell: ${this.cellTracker.getCellReference('loanIssuanceFees')})\n`;
    debtInfo += `- Fixed Rate: ${modelData.debtSettings?.fixedRate || 5.5}% (Cell: ${this.cellTracker.getCellReference('fixedRate')})`;
    
    return debtInfo;
  }

  // P&L Sheet generation using proper structure
  // DUPLICATE METHOD - COMMENTED OUT - Use the first createPLSheet method above
  async createPLSheet_OLD_DUPLICATE(modelData) {
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
      
      // Calculate number of periods
      const periods = this.calculatePeriods(modelData.projectStartDate, modelData.projectEndDate, modelData.modelPeriods);
      const periodColumns = periods; // Use full calculated periods
      
      // HEADER
      plSheet.getRange('A1').values = [[`P&L Statement (${modelData.currency})`]];
      plSheet.getRange('A1').format.font.bold = true;
      plSheet.getRange('A1').format.font.size = 16;
      currentRow = 3;
      
      // TIME PERIOD HEADERS
      const headers = [''];
      const startDate = new Date(modelData.projectStartDate);
      for (let i = 0; i < periodColumns; i++) {
        headers.push(this.formatPeriodHeader(startDate, i, modelData.modelPeriods));
      }
      
      const headerRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(periodColumns)}${currentRow}`);
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = '#e0e0e0';
      currentRow += 2;
      
      // REVENUE SECTION
      plSheet.getRange(`A${currentRow}`).values = [['REVENUE']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#87CEEB';
      currentRow += 1;
      
      // Get revenue count from model data
      const revenueCount = modelData.revenueItems ? modelData.revenueItems.length : 0;
      const revenueStartRow = currentRow;
      
      if (revenueCount > 0) {
        // Add each revenue item with growth formulas
        for (let i = 0; i < revenueCount; i++) {
          const nameRef = this.cellTracker.getCellReference(`revenue_${i}_name`);
          const valueRef = this.cellTracker.getCellReference(`revenue_${i}`);
          
          if (nameRef && valueRef) {
            // Item name - use the actual name from modelData
            const itemName = modelData.revenueItems[i]?.name || `Revenue Item ${i + 1}`;
            plSheet.getRange(`A${currentRow}`).values = [[itemName]];
            
            // Values for each period with growth
            for (let col = 1; col <= totalColumns; col++) {
              const colLetter = this.getColumnLetter(col);
              
              if (col === 1) {
                // First period - base value
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=Assumptions!${valueRef.split('!')[1]}`]];
              } else {
                // Subsequent periods - apply growth
                const prevCol = this.getColumnLetter(col - 1);
                let growthFormula;
                const growthType = modelData.revenueItems?.[i]?.growthType;
                const growthRate = modelData.revenueItems?.[i]?.annualGrowthRate;
                
                if (growthType === 'annual' && growthRate) {
                  const periodRate = this.adjustGrowthRateForPeriod(growthRate, modelData.modelPeriods);
                  growthFormula = `=${prevCol}${currentRow}*(1+${periodRate})`;
                } else {
                  growthFormula = `=${prevCol}${currentRow}`;
                }
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[growthFormula]];
              }
            }
            currentRow++;
          }
        }
        
        // Total Revenue row
        currentRow++;
        plSheet.getRange(`A${currentRow}`).values = [['Total Revenue']];
        plSheet.getRange(`A${currentRow}`).format.font.bold = true;
        
        // Sum formulas for each period - only sum actual revenue rows
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          const revenueEndRow = revenueStartRow + revenueCount - 1;
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
            [[`=SUM(${colLetter}${revenueStartRow}:${colLetter}${revenueEndRow})`]];
        }
        currentRow += 2;
      }
      
      // Store total revenue row for NOI calculation
      const totalRevenueRow = revenueCount > 0 ? currentRow - 2 : 0;
      
      // OPERATING EXPENSES SECTION
      plSheet.getRange(`A${currentRow}`).values = [['OPERATING EXPENSES']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#FFB6C1';
      currentRow += 1;
      
      // Get opex count from model data
      const opexCount = modelData.operatingExpenses ? modelData.operatingExpenses.length : 0;
      const opexStartRow = currentRow;
      
      if (opexCount > 0) {
        // Add each operating expense with growth formulas
        for (let i = 0; i < opexCount; i++) {
          const nameRef = this.cellTracker.getCellReference(`opex_${i}_name`);
          const valueRef = this.cellTracker.getCellReference(`opex_${i}`);
          
          if (nameRef && valueRef) {
            // Item name - use the actual name from modelData  
            const itemName = modelData.operatingExpenses[i]?.name || `OpEx Item ${i + 1}`;
            plSheet.getRange(`A${currentRow}`).values = [[itemName]];
            
            // Values for each period with growth
            for (let col = 1; col <= totalColumns; col++) {
              const colLetter = this.getColumnLetter(col);
              
              if (col === 1) {
                // First period - base value (negative for expenses)
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=-Assumptions!${valueRef.split('!')[1]}`]];
              } else {
                // Subsequent periods - apply growth
                const prevCol = this.getColumnLetter(col - 1);
                let growthFormula;
                const growthType = modelData.operatingExpenses?.[i]?.growthType;
                const growthRate = modelData.operatingExpenses?.[i]?.annualGrowthRate;
                
                if (growthType === 'annual' && growthRate) {
                  const periodRate = this.adjustGrowthRateForPeriod(growthRate, modelData.modelPeriods);
                  growthFormula = `=${prevCol}${currentRow}*(1+${periodRate})`;
                } else {
                  growthFormula = `=${prevCol}${currentRow}`;
                }
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[growthFormula]];
              }
            }
            currentRow++;
          }
        }
        
        // Total Operating Expenses row
        currentRow++;
        plSheet.getRange(`A${currentRow}`).values = [['Total Operating Expenses']];
        plSheet.getRange(`A${currentRow}`).format.font.bold = true;
        
        // Sum formulas for each period - only sum actual opex rows
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          const opexEndRow = opexStartRow + opexCount - 1;
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
            [[`=SUM(${colLetter}${opexStartRow}:${colLetter}${opexEndRow})`]];
        }
        currentRow += 2;
      }
      
      // Store total opex row for NOI calculation
      const totalOpexRow = opexCount > 0 ? currentRow - 2 : 0;
      
      // NOI CALCULATION
      plSheet.getRange(`A${currentRow}`).values = [['NOI']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#98FB98';
      
      // NOI formulas for each period
      for (let col = 1; col <= periodColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        let ebitdaFormula = '0';
        
        if (totalRevenueRow > 0 && totalOpexRow > 0) {
          ebitdaFormula = `=${colLetter}${totalRevenueRow}+${colLetter}${totalOpexRow}`;
        } else if (totalRevenueRow > 0) {
          ebitdaFormula = `=${colLetter}${totalRevenueRow}`;
        } else if (totalOpexRow > 0) {
          ebitdaFormula = `=${colLetter}${totalOpexRow}`;
        }
        
        plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[ebitdaFormula]];
      }
      currentRow += 2;
      
      // Real estate model: P&L ends at NOI - no debt service, interest expense, or net income calculations
      
      // Track NOI for FCF calculations
      this.plCellTracker.recordCell('noi', 'P&L Statement', `B${currentRow-2}:${this.getColumnLetter(periodColumns)}${currentRow-2}`);
      
      // Format numbers for all data columns
      const dataRange = plSheet.getRange(`B5:${this.getColumnLetter(periodColumns)}${currentRow-2}`);
      dataRange.numberFormat = [['#,##0;[Red]-#,##0;"-"']];
      
      // Auto-resize columns
      plSheet.getRange(`A:${this.getColumnLetter(periodColumns)}`).format.autofitColumns();
      
      await context.sync();
      console.log('✅ P&L sheet created with dynamic references and growth formulas');
    });
  }

  // Utility method to get all tracked data
  getTrackedData() {
    return {
      cellMap: Object.fromEntries(this.cellTracker.cellMap),
      sheetData: Object.fromEntries(this.cellTracker.sheetData)
    };
  }
  
  // Helper function to get column letter for Excel
  getColumnLetter(columnIndex) {
    if (columnIndex < 0) return 'A';
    if (columnIndex < 26) {
      return String.fromCharCode(65 + columnIndex);
    }
    // For columns beyond Z
    let result = '';
    let temp = columnIndex;
    while (temp >= 0) {
      result = String.fromCharCode(65 + (temp % 26)) + result;
      temp = Math.floor(temp / 26) - 1;
    }
    return result;
  }

  // Helper method to get currency format based on selected currency
  getCurrencyFormat(currency) {
    const currencyFormats = {
      'USD': '[$$-en-US] #,##0;[Red][$$-en-US] -#,##0',
      'EUR': '[$€-en-US] #,##0;[Red][$€-en-US] -#,##0',
      'GBP': '[$£-en-GB] #,##0;[Red][$£-en-GB] -#,##0',
      'JPY': '[$¥-ja-JP] #,##0;[Red][$¥-ja-JP] -#,##0',
      'CAD': '[$C$-en-CA] #,##0;[Red][$C$-en-CA] -#,##0',
      'AUD': '[$A$-en-AU] #,##0;[Red][$A$-en-AU] -#,##0',
      'CHF': '[$CHF-de-CH] #,##0;[Red][$CHF-de-CH] -#,##0',
      'CNY': '[$¥-zh-CN] #,##0;[Red][$¥-zh-CN] -#,##0',
      'SEK': '[$kr-sv-SE] #,##0;[Red][$kr-sv-SE] -#,##0',
      'NOK': '[$kr-nb-NO] #,##0;[Red][$kr-nb-NO] -#,##0'
    };
    
    return currencyFormats[currency] || currencyFormats['USD'];
  }
  
  // Calculate number of periods between dates
  calculatePeriods(startDate, endDate, periodType) {
    if (!startDate || !endDate) {
      return 12; // Default
    }
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    const diffTime = Math.abs(end - start);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    switch (periodType) {
      case 'daily':
        return Math.min(diffDays, 1000); // Increased cap for daily periods
      case 'monthly':
        return Math.ceil(diffDays / 30); // Removed cap for monthly periods
      case 'quarterly':
        return Math.ceil(diffDays / 90); // Removed cap for quarterly periods
      case 'yearly':
        return Math.ceil(diffDays / 365); // Removed cap for yearly periods
      default:
        return 12;
    }
  }
  
  // Format period header based on period type
  formatPeriodHeader(startDate, periodIndex, periodType) {
    const date = new Date(startDate);
    
    switch (periodType) {
      case 'daily':
        date.setDate(date.getDate() + periodIndex);
        return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
      case 'monthly':
        date.setMonth(date.getMonth() + periodIndex);
        return date.toLocaleDateString('en-US', { year: 'numeric', month: 'short' });
      case 'quarterly':
        date.setMonth(date.getMonth() + (periodIndex * 3));
        const quarter = Math.floor((date.getMonth() / 3)) + 1;
        return `Q${quarter} ${date.getFullYear()}`;
      case 'yearly':
        date.setFullYear(date.getFullYear() + periodIndex);
        return date.getFullYear().toString();
      default:
        date.setMonth(date.getMonth() + periodIndex);
        return date.toLocaleDateString('en-US', { year: 'numeric', month: 'short' });
    }
  }
  
  // Generate actual Excel date for XIRR calculations
  formatPeriodDate(startDate, periodIndex, periodType) {
    const date = new Date(startDate);
    
    switch (periodType) {
      case 'daily':
        date.setDate(date.getDate() + periodIndex);
        break;
      case 'monthly':
        date.setMonth(date.getMonth() + periodIndex);
        break;
      case 'quarterly':
        date.setMonth(date.getMonth() + (periodIndex * 3));
        break;
      case 'yearly':
        date.setFullYear(date.getFullYear() + periodIndex);
        break;
      default:
        date.setMonth(date.getMonth() + periodIndex);
    }
    
    // Convert to Excel serial date (days since 1/1/1900)
    const excelEpoch = new Date(1900, 0, 1);
    const msPerDay = 24 * 60 * 60 * 1000;
    const excelDate = Math.floor((date - excelEpoch) / msPerDay) + 2; // +2 for Excel's leap year bug
    return excelDate; // Return Excel serial date number
  }
  
  // Adjust growth rate for period type  
  adjustGrowthRateForPeriod(annualRate, periodType) {
    switch (periodType) {
      case 'daily':
        return annualRate / 365;
      case 'monthly':
        return annualRate / 12;
      case 'quarterly':
        return annualRate / 4;
      case 'yearly':
      default:
        return annualRate;
    }
  }
  
  // Get period adjustment string for formulas
  getPeriodAdjustment(periodType) {
    switch (periodType) {
      case 'daily':
        return '/365';
      case 'monthly':
        return '/12';
      case 'quarterly':
        return '/4';
      case 'yearly':
      default:
        return '';
    }
  }
  
  // Calculate period rate
  calculatePeriodRate(annualRate, periodType) {
    switch (periodType) {
      case 'daily':
        return annualRate / 365;
      case 'monthly':
        return annualRate / 12;
      case 'quarterly':
        return annualRate / 4;
      case 'yearly':
      default:
        return annualRate;
    }
  }
  
  // Generate growth formula based on period type
  getGrowthFormula(prevCellRef, growthRateRef, periodType, growthType) {
    if (!growthType || growthType === 'none' || !growthRateRef) {
      return `=${prevCellRef}`; // No growth
    }
    
    // Adjust growth rate based on period type
    let periodAdjustment = '';
    switch (periodType) {
      case 'daily':
        periodAdjustment = '/365';
        break;
      case 'monthly':
        periodAdjustment = '/12';
        break;
      case 'quarterly':
        periodAdjustment = '/4';
        break;
      case 'yearly':
      default:
        periodAdjustment = ''; // Annual rate as-is
    }
    
    // Create growth formula
    if (growthRateRef) {
      return `=${prevCellRef}*(1+${growthRateRef}${periodAdjustment})`;
    } else {
      return `=${prevCellRef}`;
    }
  }

  // Read actual P&L sheet structure to discover cell locations
  async readPLSheetStructure() {
    return Excel.run(async (context) => {
      console.log('🔍 Reading P&L sheet structure...');
      
      try {
        const plSheet = context.workbook.worksheets.getItem('P&L Statement');
        const usedRange = plSheet.getUsedRange();
        usedRange.load(['values', 'formulas', 'address']);
        await context.sync();

        const structure = {
          sheetExists: true,
          usedRange: usedRange.address,
          lineItems: {},
          periodColumns: 0
        };

        const values = usedRange.values;
        const formulas = usedRange.formulas;
        
        // Parse the P&L structure to find key line items
        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            const cellValue = values[row][col];
            const cellRef = this.getColumnLetter(col) + (row + 1);
            
            // Look for key P&L line items
            if (typeof cellValue === 'string') {
              const lowerValue = cellValue.toLowerCase();
              
              // Map line items to their row positions
              if (lowerValue.includes('revenue') && lowerValue.includes('total')) {
                structure.lineItems.totalRevenue = { row: row + 1, startCol: 'B', cellRef: `B${row + 1}` };
              }
              if (lowerValue.includes('noi') || lowerValue.includes('net operating income') || lowerValue.includes('ebitda') || (lowerValue.includes('ebit') && lowerValue.includes('da'))) {
                structure.lineItems.noi = { row: row + 1, startCol: 'B', cellRef: `B${row + 1}` };
              }
              if (lowerValue.includes('net income') || lowerValue.includes('net profit')) {
                structure.lineItems.netIncome = { row: row + 1, startCol: 'B', cellRef: `B${row + 1}` };
              }
              if (lowerValue.includes('interest') && lowerValue.includes('expense')) {
                structure.lineItems.interestExpense = { row: row + 1, startCol: 'B', cellRef: `B${row + 1}` };
              }
              if (lowerValue.includes('operating') && lowerValue.includes('expense') && lowerValue.includes('total')) {
                structure.lineItems.totalOpEx = { row: row + 1, startCol: 'B', cellRef: `B${row + 1}` };
              }
              if (lowerValue.includes('capital') && (lowerValue.includes('expenditure') || lowerValue.includes('expense'))) {
                structure.lineItems.totalCapEx = { row: row + 1, startCol: 'B', cellRef: `B${row + 1}` };
              }
            }
          }
        }

        // Determine number of period columns (excluding column A for labels)
        if (values.length > 0) {
          structure.periodColumns = values[0].length - 1; // Subtract 1 for the label column
        }

        console.log('📊 P&L Structure discovered:', structure);
        return structure;

      } catch (error) {
        console.log('❌ P&L sheet does not exist or cannot be read:', error.message);
        return {
          sheetExists: false,
          error: error.message,
          lineItems: {},
          periodColumns: 0
        };
      }
    });
  }

  // Read actual Assumptions sheet structure
  async readAssumptionSheetStructure() {
    return Excel.run(async (context) => {
      console.log('🔍 Reading Assumptions sheet structure...');
      
      try {
        const assumptionsSheet = context.workbook.worksheets.getItem('Assumptions');
        const usedRange = assumptionsSheet.getUsedRange();
        usedRange.load(['values', 'formulas', 'address']);
        await context.sync();

        const structure = {
          sheetExists: true,
          usedRange: usedRange.address,
          assumptions: {},
          sectionMap: {}
        };

        const values = usedRange.values;
        
        // Parse the Assumptions structure to find key data points
        for (let row = 0; row < values.length; row++) {
          const labelValue = values[row][0]; // Column A contains labels
          const dataValue = values[row][1]; // Column B contains data
          const cellRef = `B${row + 1}`;
          
          if (typeof labelValue === 'string') {
            const lowerLabel = labelValue.toLowerCase();
            
            // Map key assumption values to their cell references
            if (lowerLabel.includes('currency')) {
              structure.assumptions.currency = { cellRef, value: dataValue };
            }
            if (lowerLabel.includes('deal value')) {
              structure.assumptions.dealValue = { cellRef, value: dataValue };
            }
            if (lowerLabel.includes('deal ltv')) {
              structure.assumptions.dealLTV = { cellRef, value: dataValue };
            }
            if (lowerLabel.includes('transaction fee')) {
              structure.assumptions.transactionFee = { cellRef, value: dataValue };
            }
            if (lowerLabel.includes('disposal cost')) {
              structure.assumptions.disposalCost = { cellRef, value: dataValue };
            }
            if (lowerLabel.includes('terminal cap rate')) {
              structure.assumptions.terminalCapRate = { cellRef, value: dataValue };
            }
            if (lowerLabel.includes('discount rate') || lowerLabel.includes('wacc')) {
              structure.assumptions.discountRate = { cellRef, value: dataValue };
            }
            if (lowerLabel.includes('interest rate') && !lowerLabel.includes('type')) {
              structure.assumptions.interestRate = { cellRef, value: dataValue };
            }
            if (lowerLabel.includes('equity contribution')) {
              structure.assumptions.equityContribution = { cellRef, value: dataValue };
            }
            if (lowerLabel.includes('debt financing')) {
              structure.assumptions.debtFinancing = { cellRef, value: dataValue };
            }
            
            // Track revenue items
            if (labelValue.includes('Revenue Item') || labelValue.includes('Product Sales') || 
                labelValue.includes('Service Revenue') || labelValue.includes('Sales')) {
              if (!structure.assumptions.revenueItems) structure.assumptions.revenueItems = [];
              structure.assumptions.revenueItems.push({ 
                name: labelValue, 
                cellRef, 
                value: dataValue 
              });
            }
            
            // Track operating expenses
            if ((lowerLabel.includes('expense') || lowerLabel.includes('cost')) && 
                !lowerLabel.includes('capital') && !lowerLabel.includes('disposal')) {
              if (!structure.assumptions.operatingExpenses) structure.assumptions.operatingExpenses = [];
              structure.assumptions.operatingExpenses.push({ 
                name: labelValue, 
                cellRef, 
                value: dataValue 
              });
            }
            
            // Track capital expenses
            if (lowerLabel.includes('capital') && (lowerLabel.includes('expense') || lowerLabel.includes('expenditure'))) {
              if (!structure.assumptions.capitalExpenses) structure.assumptions.capitalExpenses = [];
              structure.assumptions.capitalExpenses.push({ 
                name: labelValue, 
                cellRef, 
                value: dataValue 
              });
            }
          }
        }

        console.log('📊 Assumptions Structure discovered:', structure);
        return structure;

      } catch (error) {
        console.log('❌ Assumptions sheet does not exist or cannot be read:', error.message);
        return {
          sheetExists: false,
          error: error.message,
          assumptions: {},
          sectionMap: {}
        };
      }
    });
  }

  // Generate comprehensive FCF prompt using REAL cell references from P&L and Assumptions
  generateRealFCFPrompt(modelData, plStructure, assumptionStructure, capExStructure) {
    console.log('🤖 Generating REAL FCF AI prompt with discovered cell references...');
    console.log('🔍 capExStructure parameter:', capExStructure);
    console.log('🔍 capExStructure type:', typeof capExStructure);
    
    // Ensure capExStructure is properly defined for template literals
    const safeCapExStructure = capExStructure || { totalRow: null, sheet: null };
    console.log('🔍 safeCapExStructure:', safeCapExStructure);
    
    const periods = this.calculatePeriods(modelData.projectStartDate, modelData.projectEndDate, modelData.modelPeriods);
    const maxPeriods = Math.min(periods, 36);
    
    // Build period headers
    const periodHeaders = [];
    const startDate = new Date(modelData.projectStartDate);
    for (let i = 0; i < maxPeriods; i++) {
      periodHeaders.push(this.formatPeriodHeader(startDate, i, modelData.modelPeriods));
    }

    const prompt = `You are a senior financial analyst specializing in M&A Free Cash Flow modeling. You have been provided with ACTUAL cell references from an existing P&L Statement and Assumptions sheet.

**PROJECT OVERVIEW:**
- Deal Name: ${modelData.dealName}
- Currency: ${modelData.currency}
- Period Type: ${modelData.modelPeriods}
- Total Periods: ${maxPeriods}
- Date Range: ${modelData.projectStartDate} to ${modelData.projectEndDate}

**ACTUAL P&L SHEET STRUCTURE DISCOVERED:**
${this.formatPLStructureForPrompt(plStructure, maxPeriods)}

**ACTUAL ASSUMPTIONS SHEET REFERENCES:**
${this.formatAssumptionStructureForPrompt(assumptionStructure, safeCapExStructure, maxPeriods)}

**ACTUAL CAPEX SHEET STRUCTURE:**
${this.formatCapExStructureForPrompt(safeCapExStructure, maxPeriods)}

**REQUIRED FCF SHEET STRUCTURE:**

Create a comprehensive Free Cash Flow statement with the following structure:
- Column A: Line item names  
- Columns B through ${this.getColumnLetter(maxPeriods)}: Period data (${maxPeriods} periods total)
- Period Headers: ${periodHeaders.join(', ')}

**FCF CALCULATION METHODOLOGY - REAL ESTATE MODEL:**

You MUST create a Free Cash Flow Statement with this EXACT structure:

**UNLEVERED CASH FLOWS:**
1. **Period** - Column headers for all periods
2. **Purchase price** - Deal Value in Period 0 only (negative)
3. **Transaction costs** - Transaction fees in Period 0 only (negative)
4. **EBITDA** - Reference NOI from P&L (all operating periods)
5. **CapEx** - Reference Total CapEx from CapEx sheet using ='CapEx'!C${safeCapExStructure?.totalRow || 'X'} format (all periods, negative values)
6. **Sale Price** - Terminal value in final period only (positive)
7. **Disposal Costs** - Disposal costs in final period only (negative)
8. **Unlevered Cashflows** - Sum of all above items per period

**LEVERED CASH FLOWS:**
9. **Debt drawn** - Debt amount in Period 0 only (positive)
10. **Debt upfront costs** - Loan issuance fees in Period 0 only (negative)
11. **Interest Payments** - Interest payments calculated in FCF (all debt periods)
12. **Debt repaid (Interest-only-loan)** - Principal repayment in final period only (negative)
13. **Levered Cashflows** - Unlevered Cashflows + Debt flows per period

**EQUITY FLOWS:**
14. **Equity distributions** - Levered Cashflows (all periods)

**RETURNS:**
16. **Unlevered IRR** - XIRR calculation on Unlevered Cashflows
17. **MOIC** - Multiple of Invested Capital calculation
   - Undiscounted NPV (simple sum of FCF)
   - Discounted NPV using WACC from Assumptions
   - IRR calculation including initial investment

**CRITICAL REQUIREMENTS:**

1. **Use EXACT cell references** from the discovered P&L and Assumptions structures above
2. **Reference format**: Use 'P&L Statement'!B15 or 'Assumptions'!B23 format  
3. **Handle missing data**: If a P&L line item is not found, use conservative estimates
4. **Most Important**: ALWAYS reference the NOI line from P&L as the starting point for FCF calculations
5. **Period consistency**: Ensure all ${maxPeriods} periods are calculated
6. **Formula accuracy**: All formulas must be valid Excel syntax

**NOI PRIORITY:**
The most critical value to extract from the P&L is the NOI (Net Operating Income) for each period, located at: ${plStructure.lineItems?.noi?.cellRef || 'Not Found - Please locate manually'}

**EXPECTED OUTPUT:**
Provide complete Excel range setup with exact cell addresses and formulas for all ${maxPeriods} periods. Include proper formatting instructions and ensure all calculations reference the actual discovered cell locations.

If any critical P&L references are missing, clearly state what assumptions you're making and recommend manual verification.`;

    return prompt;
  }

  // Format P&L structure for AI prompt
  formatPLStructureForPrompt(plStructure, maxPeriods) {
    if (!plStructure.sheetExists) {
      return `❌ P&L Sheet not found or unreadable. Error: ${plStructure.error || 'Unknown error'}`;
    }

    let output = `✅ P&L Sheet discovered with ${plStructure.periodColumns} period columns\n`;
    output += `📍 Used Range: ${plStructure.usedRange}\n\n`;
    
    output += `**KEY P&L LINE ITEMS FOUND:**\n`;
    
    if (plStructure.lineItems.totalRevenue) {
      output += `- Total Revenue: Row ${plStructure.lineItems.totalRevenue.row}, Range B${plStructure.lineItems.totalRevenue.row}:${this.getColumnLetter(maxPeriods)}${plStructure.lineItems.totalRevenue.row}\n`;
    }
    
    if (plStructure.lineItems.totalOpEx) {
      output += `- Total Operating Expenses: Row ${plStructure.lineItems.totalOpEx.row}, Range B${plStructure.lineItems.totalOpEx.row}:${this.getColumnLetter(maxPeriods)}${plStructure.lineItems.totalOpEx.row}\n`;
    }
    
    if (plStructure.lineItems.noi) {
      output += `- NOI: Row ${plStructure.lineItems.noi.row}, Range B${plStructure.lineItems.noi.row}:${this.getColumnLetter(maxPeriods)}${plStructure.lineItems.noi.row}\n`;
    }
    
    if (plStructure.lineItems.noi) {
      output += `- NOI (Net Operating Income): Row ${plStructure.lineItems.noi.row}, Range B${plStructure.lineItems.noi.row}:${this.getColumnLetter(maxPeriods)}${plStructure.lineItems.noi.row}\n`;
    }
    
    if (plStructure.lineItems.netIncome) {
      output += `- 🎯 NET INCOME (CRITICAL): Row ${plStructure.lineItems.netIncome.row}, Range B${plStructure.lineItems.netIncome.row}:${this.getColumnLetter(maxPeriods)}${plStructure.lineItems.netIncome.row}\n`;
    }
    
    if (plStructure.lineItems.totalCapEx) {
      output += `- Total CapEx: Row ${plStructure.lineItems.totalCapEx.row}, Range B${plStructure.lineItems.totalCapEx.row}:${this.getColumnLetter(maxPeriods)}${plStructure.lineItems.totalCapEx.row}\n`;
    }

    output += `\n**REFERENCE FORMAT:** Use 'P&L Statement'!B[row]:[column][row] for ranges\n`;
    output += `**EXAMPLE:** 'P&L Statement'!B${plStructure.lineItems.noi?.row || '15'}:${this.getColumnLetter(maxPeriods)}${plStructure.lineItems.noi?.row || '15'} for NOI across all periods\n\n`;

    // CAPEX STRUCTURE REMOVED - This belongs in formatCapExStructureForPrompt, not here

    return output;
  }

  // Format Assumptions structure for AI prompt  
  formatCapExStructureForPrompt(capExStructure, maxPeriods) {
    if (!capExStructure || !capExStructure.totalRow) {
      return 'No CapEx sheet found - use zero values for all periods';
    }
    
    let output = `CapEx sheet 'CapEx' exists with:\n`;
    output += `- Total CapEx row: ${capExStructure?.totalRow}\n`;
    output += `- Reference format: ='CapEx'!C${capExStructure?.totalRow} (for period 1)\n`;
    output += `- Use columns C through ${this.getColumnLetter(maxPeriods + 2)} for all periods\n`;
    output += `- IMPORTANT: CapEx should be NEGATIVE in FCF (cash outflow)\n`;
    
    return output;
  }

  formatAssumptionStructureForPrompt(assumptionStructure, capExStructure, maxPeriods) {
    if (!assumptionStructure.sheetExists) {
      return `❌ Assumptions Sheet not found or unreadable. Error: ${assumptionStructure.error || 'Unknown error'}`;
    }

    let output = `✅ Assumptions Sheet discovered\n`;
    output += `📍 Used Range: ${assumptionStructure.usedRange}\n\n`;
    
    output += `**KEY ASSUMPTIONS AVAILABLE:**\n`;
    
    const assumptions = assumptionStructure.assumptions;
    
    if (assumptions.currency) {
      output += `- Currency: ${assumptions.currency.cellRef} (${assumptions.currency.value})\n`;
    }
    if (assumptions.dealValue) {
      output += `- Deal Value: ${assumptions.dealValue.cellRef} (${assumptions.dealValue.value})\n`;
    }
    if (assumptions.dealLTV) {
      output += `- Deal LTV: ${assumptions.dealLTV.cellRef} (${assumptions.dealLTV.value}%)\n`;
    }
    if (assumptions.interestRate) {
      output += `- Interest Rate: ${assumptions.interestRate.cellRef} (${assumptions.interestRate.value}%)\n`;
    }
    if (assumptions.terminalCapRate) {
      output += `- Terminal Cap Rate: ${assumptions.terminalCapRate.cellRef} (${assumptions.terminalCapRate.value}%)\n`;
    }
    if (assumptions.disposalCost) {
      output += `- Disposal Cost: ${assumptions.disposalCost.cellRef} (${assumptions.disposalCost.value}%)\n`;
    }
    if (assumptions.discountRate) {
      output += `- Discount Rate (WACC): ${assumptions.discountRate.cellRef} (${assumptions.discountRate.value}%)\n`;
    }

    if (assumptions.revenueItems && assumptions.revenueItems.length > 0) {
      output += `\n**REVENUE ITEMS:**\n`;
      assumptions.revenueItems.forEach(item => {
        output += `- ${item.name}: ${item.cellRef} (${item.value})\n`;
      });
    }

    if (assumptions.operatingExpenses && assumptions.operatingExpenses.length > 0) {
      output += `\n**OPERATING EXPENSES:**\n`;
      assumptions.operatingExpenses.forEach(item => {
        output += `- ${item.name}: ${item.cellRef} (${item.value})\n`;
      });
    }

    if (assumptions.capitalExpenses && assumptions.capitalExpenses.length > 0) {
      output += `\n**CAPITAL EXPENSES:**\n`;
      assumptions.capitalExpenses.forEach(item => {
        output += `- ${item.name}: ${item.cellRef} (${item.value})\n`;
      });
    }

    output += `\n**REFERENCE FORMAT:** Use 'Assumptions'!B[row] for individual cells\n\n`;

    return output;
  }
  
  // Read FCF sheet structure to get cell references
  async readFCFSheetStructure() {
    return Excel.run(async (context) => {
      try {
        const sheets = context.workbook.worksheets;
        const fcfSheet = sheets.getItemOrNullObject('Free Cash Flow');
        fcfSheet.load('name');
        await context.sync();
        
        if (fcfSheet.isNullObject) {
          throw new Error('Free Cash Flow sheet not found');
        }
        
        // Read the entire sheet to find structure
        const range = fcfSheet.getUsedRange();
        range.load('values');
        await context.sync();
        
        const values = range.values;
        const structure = {
          sheetName: 'Free Cash Flow',
          periodColumns: 0,
          leveredFCF: null,
          unleveredFCF: null,
          cumulativeFCF: null,
          discountedNPV: null,
          undiscountedNPV: null,
          cashFlowRange: null
        };
        
        // Find key rows and structure
        for (let i = 0; i < values.length; i++) {
          const row = values[i];
          if (row && row[0]) {
            const cellValue = row[0].toString().toLowerCase();
            
            if (cellValue.includes('levered free cash flow')) {
              structure.leveredFCF = i + 1;
            } else if (cellValue.includes('unlevered free cash flow')) {
              structure.unleveredFCF = i + 1;
            } else if (cellValue.includes('cumulative free cash flow')) {
              structure.cumulativeFCF = i + 1;
            } else if (cellValue.includes('discounted npv')) {
              structure.discountedNPV = i + 1;
            } else if (cellValue.includes('undiscounted npv')) {
              structure.undiscountedNPV = i + 1;
            }
          }
        }
        
        // Find number of period columns
        if (values.length > 0) {
          structure.periodColumns = values[0].length - 1; // Subtract 1 for the label column
        }
        
        // Define cash flow range for IRR calculations
        if (structure.leveredFCF) {
          structure.cashFlowRange = `B${structure.leveredFCF}:${this.getColumnLetter(structure.periodColumns)}${structure.leveredFCF}`;
        }
        
        console.log('📊 FCF Structure discovered:', structure);
        return structure;
        
      } catch (error) {
        console.error('Error reading FCF sheet structure:', error);
        throw error;
      }
    });
  }
  
  // Read actual FCF sheet data values
  async readFCFSheetData() {
    return Excel.run(async (context) => {
      try {
        const sheets = context.workbook.worksheets;
        const fcfSheet = sheets.getItemOrNullObject('Free Cash Flow');
        fcfSheet.load('name');
        await context.sync();
        
        if (fcfSheet.isNullObject) {
          throw new Error('Free Cash Flow sheet not found');
        }
        
        // Read the entire sheet data
        const range = fcfSheet.getUsedRange();
        range.load('values');
        await context.sync();
        
        const values = range.values;
        const fcfData = {
          sheetName: 'Free Cash Flow',
          rawData: values,
          leveredFCFValues: [],
          unleveredFCFValues: [],
          periodHeaders: [],
          cashFlowPeriods: []
        };
        
        // Extract period headers (first row)
        if (values.length > 0) {
          fcfData.periodHeaders = values[0].slice(1); // Remove first column (labels)
        }
        
        // Find and extract actual cash flow values
        for (let i = 0; i < values.length; i++) {
          const row = values[i];
          if (row && row[0]) {
            const cellValue = row[0].toString().toLowerCase();
            
            if (cellValue.includes('levered free cash flow')) {
              fcfData.leveredFCFValues = row.slice(1); // Remove first column (label)
            } else if (cellValue.includes('unlevered free cash flow')) {
              fcfData.unleveredFCFValues = row.slice(1); // Remove first column (label)
            }
          }
        }
        
        // Create cash flow periods for IRR calculation
        fcfData.cashFlowPeriods = fcfData.leveredFCFValues.map((value, index) => ({
          period: fcfData.periodHeaders[index] || `Period ${index + 1}`,
          leveredFCF: parseFloat(value) || 0,
          unleveredFCF: parseFloat(fcfData.unleveredFCFValues[index]) || 0
        }));
        
        console.log('💰 FCF Data extracted:', fcfData);
        return fcfData;
        
      } catch (error) {
        console.error('Error reading FCF sheet data:', error);
        throw error;
      }
    });
  }
  
  // Read actual P&L sheet data values
  async readPLSheetData() {
    return Excel.run(async (context) => {
      try {
        const sheets = context.workbook.worksheets;
        const plSheet = sheets.getItemOrNullObject('P&L Statement');
        plSheet.load('name');
        await context.sync();
        
        if (plSheet.isNullObject) {
          throw new Error('P&L Statement sheet not found');
        }
        
        // Read the entire sheet data
        const range = plSheet.getUsedRange();
        range.load('values');
        await context.sync();
        
        const values = range.values;
        const plData = {
          sheetName: 'P&L Statement',
          rawData: values,
          revenue: [],
          expenses: [],
          ebitda: [],
          netIncome: [],
          periodHeaders: []
        };
        
        // Extract period headers (first row)
        if (values.length > 0) {
          plData.periodHeaders = values[0].slice(1); // Remove first column (labels)
        }
        
        // Find and extract key P&L line items
        for (let i = 0; i < values.length; i++) {
          const row = values[i];
          if (row && row[0]) {
            const cellValue = row[0].toString().toLowerCase();
            
            if (cellValue.includes('total revenue')) {
              plData.revenue = row.slice(1);
            } else if (cellValue.includes('noi') || cellValue.includes('net operating income') || cellValue.includes('ebitda')) {
              plData.noi = row.slice(1);
            } else if (cellValue.includes('net income')) {
              plData.netIncome = row.slice(1);
            }
          }
        }
        
        console.log('📊 P&L Data extracted:', plData);
        return plData;
        
      } catch (error) {
        console.error('Error reading P&L sheet data:', error);
        throw error;
      }
    });
  }
  
  // Call OpenAI API for Multiples & IRR calculation
  async callOpenAIForMultiples(prompt) {
    try {
      // Use the same API endpoint as AIExtractionService
      const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
      const apiEndpoint = isLocal ? 'http://localhost:8888/.netlify/functions/chat' : '/.netlify/functions/chat';
      
      // Calculate approximate token count
      const tokenCount = Math.ceil(prompt.length / 4); // Rough estimate: 4 chars per token
      console.log(`🤖 API Request Details:
        - Endpoint: ${apiEndpoint}
        - Prompt length: ${prompt.length} characters
        - Estimated tokens: ${tokenCount}
        - Request type: financial_analysis`);
      
      const requestBody = {
        message: prompt,
        autoFillMode: true, // Set to true to use batch processing
        batchType: 'financial_analysis',
        systemPrompt: null, // Let chat.js handle the system prompt based on batchType
        temperature: 0.1,
        maxTokens: 1500 // Increased for comprehensive IRR/MOIC analysis
      };
      
      console.log('🤖 Sending Multiples & IRR request to OpenAI...');
      console.log('📝 Request body:', JSON.stringify(requestBody, null, 2));
      
      const response = await fetch(apiEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        body: JSON.stringify(requestBody)
      });

      console.log(`📡 API Response Status: ${response.status} ${response.statusText}`);
      
      if (!response.ok) {
        // Try to get error details from response
        let errorDetails = response.statusText;
        try {
          const errorData = await response.text();
          console.log('❌ API Error Response:', errorData);
          errorDetails = errorData || response.statusText;
          
          // Check for specific error types
          if (response.status === 502) {
            console.log('❌ 502 Bad Gateway - likely token limit or API timeout');
          } else if (response.status === 413) {
            console.log('❌ 413 Payload Too Large - request body too large');
          } else if (response.status === 429) {
            console.log('❌ 429 Too Many Requests - rate limit exceeded');
          }
        } catch (e) {
          console.log('❌ Could not read error response');
        }
        throw new Error(`API error: ${response.status} ${errorDetails}`);
      }

      const data = await response.json();
      console.log('✅ API Response Data:', data);
      
      if (data.error) {
        throw new Error(data.error);
      }

      return data;
      
    } catch (error) {
      console.error('❌ Error calling OpenAI for Multiples & IRR:', error);
      console.error('❌ Error details:', {
        message: error.message,
        stack: error.stack
      });
      throw error;
    }
  }
  
  // Generate comprehensive AI prompt with ALL actual data for IRR/MOIC calculations
  generateMultiplesAndIRRPrompt(modelData, fcfStructure, fcfData, plData, assumptionStructure) {
    console.log('🤖 Generating comprehensive AI prompt with actual financial data...');
    
    // Calculate periods and other key metrics
    const periods = this.calculatePeriods(modelData.projectStartDate, modelData.projectEndDate, modelData.modelPeriods);
    const holdingPeriodYears = periods / (modelData.modelPeriods === 'monthly' ? 12 : 
                                          modelData.modelPeriods === 'quarterly' ? 4 : 
                                          modelData.modelPeriods === 'yearly' ? 1 : 12);
    
    // Get actual FCF values
    const leveredFCFValues = fcfData.leveredFCFValues || [];
    const unleveredFCFValues = fcfData.unleveredFCFValues || [];
    
    // Calculate equity contribution if missing
    let equityContribution = modelData.equityContribution;
    if (!equityContribution || equityContribution === 0) {
      const dealValue = modelData.dealValue || 0;
      const dealLTV = modelData.dealLTV || 70;
      equityContribution = dealValue * (100 - dealLTV) / 100;
    }
    
    const prompt = `You are an expert M&A financial analyst. I need you to calculate IRR and MOIC for this investment and provide Excel formulas.

**INVESTMENT DETAILS:**
- Deal Value: ${modelData.dealValue} ${modelData.currency}
- Equity Investment: ${equityContribution} ${modelData.currency}
- Investment Period: ${holdingPeriodYears} years (${modelData.projectStartDate} to ${modelData.projectEndDate})
- Model Frequency: ${modelData.modelPeriods}

**ACTUAL CASH FLOW DATA:**
Levered Free Cash Flow (${leveredFCFValues.length} periods): [${leveredFCFValues.slice(0, 10).join(', ')}${leveredFCFValues.length > 10 ? '...' : ''}]
Unlevered Free Cash Flow (${unleveredFCFValues.length} periods): [${unleveredFCFValues.slice(0, 10).join(', ')}${unleveredFCFValues.length > 10 ? '...' : ''}]

**EXCEL SHEET STRUCTURE:**
- FCF Sheet Name: "${fcfStructure.sheetName}"
- Levered FCF Row: ${fcfStructure.leveredFCF}
- Unlevered FCF Row: ${fcfStructure.unleveredFCF}
- Data Columns: B to ${this.getColumnLetter(fcfStructure.periodColumns || 10)}

**TASK:**
Create Excel formulas for:
1. Levered IRR - Include initial equity investment as Year 0 cash flow
2. Unlevered IRR - Include initial equity investment as Year 0 cash flow  
3. Levered MOIC - Total cash returns / Initial investment
4. Unlevered MOIC - Total unlevered returns / Initial investment

**CRITICAL REQUIREMENTS:**
- IRR formulas MUST include the initial investment (-${equityContribution}) as the first cash flow
- Use proper Excel IRR syntax with cash flow arrays
- Handle potential #NUM! errors with IFERROR
- Formulas must reference the actual Excel sheet ranges provided

**RETURN FORMAT:**
{
  "calculations": {
    "leveredIRR": {
      "formula": "=IFERROR(IRR({initial_investment;cash_flows}), 0)",
      "description": "Levered IRR including initial investment"
    },
    "unleveredIRR": {
      "formula": "=IFERROR(IRR({initial_investment;cash_flows}), 0)", 
      "description": "Unlevered IRR including initial investment"
    },
    "leveredMOIC": {
      "formula": "=SUM(range)/${equityContribution}",
      "description": "Levered MOIC calculation"
    },
    "unleveredMOIC": {
      "formula": "=SUM(range)/${equityContribution}",
      "description": "Unlevered MOIC calculation"
    }
  }
}

Generate working Excel formulas using the provided data and structure.`;

    return prompt;
  }
  
  // NEW: Clean AI-only sheet creation with NO fallback logic
  async createCleanMultiplesSheet(modelData, aiResponse, equityContribution) {
    return Excel.run(async (context) => {
      console.log('📊 Creating CLEAN IRR/MOIC sheet (v5.0)...');
      
      // Parse AI response
      let calculations = {};
      try {
        if (typeof aiResponse.response === 'string') {
          const jsonMatch = aiResponse.response.match(/\{[\s\S]*\}/);
          if (jsonMatch) {
            const parsed = JSON.parse(jsonMatch[0]);
            calculations = parsed.calculations || {};
          }
        } else if (aiResponse.calculations) {
          calculations = aiResponse.calculations;
        } else if (aiResponse.extractedData) {
          calculations = aiResponse.extractedData.calculations || {};
        }
        console.log('🤖 Parsed calculations:', calculations);
      } catch (error) {
        throw new Error(`Failed to parse AI response: ${error.message}`);
      }
      
      // Delete existing sheet if it exists
      const sheets = context.workbook.worksheets;
      try {
        const existingSheet = sheets.getItemOrNullObject('IRR & MOIC Analysis');
        existingSheet.load('name');
        await context.sync();
        if (!existingSheet.isNullObject) {
          existingSheet.delete();
          await context.sync();
        }
      } catch (e) {}
      
      // Create new sheet with different name
      const sheet = sheets.add('IRR & MOIC Analysis');
      sheet.activate();
      await context.sync();
      
      // Set title - NO FALLBACK ANYWHERE
      sheet.getRange('A1').values = [['AI-POWERED IRR & MOIC ANALYSIS']];
      sheet.getRange('A1').format.font.bold = true;
      sheet.getRange('A1').format.font.size = 18;
      sheet.getRange('A1').format.fill.color = '#2E8B57';
      sheet.getRange('A1').format.font.color = 'white';
      
      let row = 3;
      
      // Deal summary
      sheet.getRange(`A${row}`).values = [['Deal Value:']];
      sheet.getRange(`B${row}`).values = [[modelData.dealValue]];
      sheet.getRange(`B${row}`).format.numberFormat = [['#,##0;[Red]-#,##0']];
      row++;
      
      sheet.getRange(`A${row}`).values = [['Equity Investment:']];
      sheet.getRange(`B${row}`).values = [[equityContribution]];
      sheet.getRange(`B${row}`).format.numberFormat = [['#,##0;[Red]-#,##0']];
      row += 2;
      
      // AI-generated calculations
      if (calculations.leveredIRR && calculations.leveredIRR.formula) {
        sheet.getRange(`A${row}`).values = [['Levered IRR:']];
        try {
          sheet.getRange(`B${row}`).formulas = [[calculations.leveredIRR.formula]];
          sheet.getRange(`B${row}`).format.numberFormat = [['0.00%']];
        } catch (formulaError) {
          console.error('IRR formula error:', formulaError);
          sheet.getRange(`B${row}`).values = [['AI Formula Error']];
        }
        row++;
      }
      
      if (calculations.leveredMOIC && calculations.leveredMOIC.formula) {
        sheet.getRange(`A${row}`).values = [['Levered MOIC:']];
        try {
          sheet.getRange(`B${row}`).formulas = [[calculations.leveredMOIC.formula]];
          sheet.getRange(`B${row}`).format.numberFormat = [['0.00"x"']];
        } catch (formulaError) {
          console.error('MOIC formula error:', formulaError);
          sheet.getRange(`B${row}`).values = [['AI Formula Error']];
        }
        row++;
      }
      
      // Auto-fit columns
      sheet.getRange('A:B').format.autofitColumns();
      
      console.log('✅ Clean IRR/MOIC sheet created successfully - NO FALLBACK MODE');
    });
  }
  
  // Generate enhanced P&L prompt with specific cell references and formula examples
  generateEnhancedPLPrompt(modelData) {
    console.log('🤖 Generating enhanced P&L prompt with cell references...');
    
    const periods = this.calculatePeriods(modelData.projectStartDate, modelData.projectEndDate, modelData.modelPeriods);
    const maxPeriods = Math.min(periods, 60);
    
    let prompt = `You are an Excel financial modeling expert. Create a P&L Statement with EXACT formulas.

**CRITICAL REQUIREMENTS:**
1. Use EXACT cell references provided below
2. Reference Assumptions sheet for ALL growth rates
3. Use proper period adjustments for growth calculations

**PROJECT DETAILS:**
- Currency: ${modelData.currency}
- Period Type: ${modelData.modelPeriods}
- Total Periods: ${maxPeriods}

**REVENUE ITEMS WITH CELL REFERENCES:**\n`;

    // Add revenue items with specific formula examples
    if (modelData.revenueItems) {
      modelData.revenueItems.forEach((item, index) => {
        const valueRef = this.cellTracker.getCellReference(`revenue_${index}`);
        const growthRateRef = this.cellTracker.getCellReference(`revenue_${index}_growth_rate`);
        
        prompt += `\n${index + 1}. ${item.name}:
   - Period 1: =${valueRef}
   - Period 2+: `;
        
        if (growthRateRef && item.growthType === 'annual') {
          const cellRef = growthRateRef.includes('!') ? growthRateRef.split('!')[1] : growthRateRef;
          if (modelData.modelPeriods === 'quarterly') {
            prompt += `=PreviousCell*(1+Assumptions!${cellRef}/4)`;
          } else if (modelData.modelPeriods === 'monthly') {
            prompt += `=PreviousCell*(1+Assumptions!${cellRef}/12)`;
          } else if (modelData.modelPeriods === 'yearly') {
            prompt += `=PreviousCell*(1+Assumptions!${cellRef})`;
          }
          prompt += `\n   - Growth Rate Location: Assumptions!${cellRef}`;
        } else {
          prompt += `=PreviousCell (no growth)`;
        }
      });
    }

    prompt += `\n\n**OPERATING EXPENSES WITH CELL REFERENCES:**\n`;

    // Add operating expenses
    if (modelData.operatingExpenses) {
      modelData.operatingExpenses.forEach((item, index) => {
        const valueRef = this.cellTracker.getCellReference(`opex_${index}`);
        const growthRateRef = this.cellTracker.getCellReference(`opex_${index}_growth_rate`);
        
        prompt += `\n${index + 1}. ${item.name}:
   - Period 1: =-${valueRef}
   - Period 2+: `;
        
        if (growthRateRef && item.growthType === 'annual') {
          const cellRef = growthRateRef.includes('!') ? growthRateRef.split('!')[1] : growthRateRef;
          if (modelData.modelPeriods === 'quarterly') {
            prompt += `=PreviousCell*(1+Assumptions!${cellRef}/4)`;
          } else if (modelData.modelPeriods === 'monthly') {
            prompt += `=PreviousCell*(1+Assumptions!${cellRef}/12)`;
          } else if (modelData.modelPeriods === 'yearly') {
            prompt += `=PreviousCell*(1+Assumptions!${cellRef})`;
          }
          prompt += `\n   - Growth Rate Location: Assumptions!${cellRef}`;
        } else {
          prompt += `=PreviousCell (no growth)`;
        }
      });
    }

    // Real estate model: No depreciation calculations required

    prompt += `\n\n**REQUIRED P&L STRUCTURE:**
You MUST create a P&L Statement with this EXACT structure:

**PERIODS:**
- Period 1: First operating period (${modelData.projectStartDate})
- Period 2+: Subsequent operating periods

**P&L STRUCTURE:**
1. Title: Profit & Loss Statement (no extra blank row below)
2. Revenue Items section header
3. Skip one row
4. Dates row (two cells below Revenue Items) - white background, black text, dashed underline
5. Period row (Period 1, Period 2, etc.)
6. Revenue items with growth formulas
7. Total Revenue (underline at top, no background)
8. Cost Items section header
9. Cost items with growth formulas  
10. Total Operating Expenses (underline at top, no background)
11. NOI (Total Revenue + Total Operating Expenses) - positioned directly below Total Operating Expenses

**FORMATTING REQUIREMENTS (CRITICAL):**
- Font: Times New Roman, Size 12 for ALL cells
- Section Headers (Revenue Items, Cost Items): Dark blue (#002060) background with white text, span all columns
- Title (Profit & Loss Statement): Background color #F2F2F2, left-aligned, span all columns, NO extra blank row below
- Dates row: White background, black text, bold, dashed underline
- Period row: Bold text
- NOI row: Light grey (#F2F2F2) background with black text, underline at top, span all columns
- Total rows (Total Revenue, Total Operating Expenses): Bold with thin black underline at TOP of cell, no background
- Section headers: Add thin grey (#D3D3D3) underline
- Number format: #,##0;[Red](#,##0);"-" (positive, negative with red brackets, zero as dash)
- Hide gridlines on the sheet

**CRITICAL REQUIREMENTS:**
- Real estate model: No depreciation or tax calculations required
- ALL formulas must reference exact cells provided above
- Apply ALL formatting requirements listed above
- Return complete P&L structure in JSON format with exact formulas

**JSON FORMAT REQUIRED:**
{
  "plStructure": {
    "revenueItems": [{"name": "...", "formulas": ["=cell1", "=cell2", ...]}],
    "totalRevenue": {"formula": "=SUM(...)"},
    "operatingExpenses": [{"name": "...", "formulas": ["=cell1", "=cell2", ...]}],
    "totalOpEx": {"formula": "=SUM(...)"},
    "noi": {"formula": "=TotalRevenue-TotalOpEx"}
  }
}`;

    console.log('📋 Enhanced P&L Prompt:', prompt);
    return prompt;
  }

  // Call OpenAI API for P&L generation
  async callOpenAIForPL(prompt) {
    try {
      const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
      const apiEndpoint = isLocal ? 'http://localhost:8888/.netlify/functions/chat' : '/.netlify/functions/chat';
      
      const requestBody = {
        message: prompt,
        autoFillMode: true,
        batchType: 'pl_generation',
        systemPrompt: 'You are an Excel expert. Generate P&L formulas exactly as specified.',
        temperature: 0.1,
        maxTokens: 2000
      };
      
      console.log('🤖 Calling OpenAI for P&L generation...');
      
      const response = await fetch(apiEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        body: JSON.stringify(requestBody)
      });

      if (!response.ok) {
        throw new Error(`API error: ${response.status}`);
      }

      const data = await response.json();
      return data;
      
    } catch (error) {
      console.error('❌ Error calling OpenAI for P&L:', error);
      throw error;
    }
  }

  // Create P&L sheet from AI response
  async createAIPLSheet(modelData, aiResponse) {
    console.log('📊 Creating P&L from AI response...');
    console.log('AI Response:', aiResponse);
    
    try {
      // Parse AI response to get P&L structure
      let plStructure = null;
      if (aiResponse && aiResponse.content) {
        try {
          // Try to extract JSON from AI response
          const jsonMatch = aiResponse.content.match(/\{[\s\S]*\}/);
          if (jsonMatch) {
            plStructure = JSON.parse(jsonMatch[0]);
          }
        } catch (parseError) {
          console.log('⚠️ Could not parse AI response as JSON:', parseError);
        }
      }
      
      if (plStructure && plStructure.plStructure) {
        console.log('✅ Using AI-generated P&L structure');
        await this.createPLSheet(modelData);
      } else {
        console.log('⚠️ Falling back to template P&L (no valid AI structure)');
        await this.createEnhancedPLSheet(modelData);
      }
    } catch (error) {
      console.error('❌ Error creating AI P&L sheet:', error);
      console.log('⚠️ Falling back to template P&L');
      await this.createEnhancedPLSheet(modelData);
    }
  }

  // Create P&L sheet with enhanced depreciation handling
  async createEnhancedPLSheet(modelData) {
    return Excel.run(async (context) => {
      console.log('📈 Creating enhanced P&L Statement with depreciation...');
      
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

      // Calculate periods
      const periods = this.calculatePeriods(modelData.projectStartDate, modelData.projectEndDate, modelData.modelPeriods);
      const periodColumns = periods;
      const totalColumns = periodColumns + 1; // +1 for Initial Investment period

      let currentRow = 1;

      // TITLE
      plSheet.getRange('A1').values = [['Profit & Loss Statement']];
      const titleRange = plSheet.getRange(`A1:${this.getColumnLetter(totalColumns)}1`);
      titleRange.merge();
      titleRange.format.font.name = 'Times New Roman';
      titleRange.format.font.size = 12;
      titleRange.format.font.bold = true;
      titleRange.format.fill.color = ExcelFormatter.colors.backgroundDarker5;
      titleRange.format.horizontalAlignment = 'Left';
      currentRow = 2; // Remove extra blank row

      // REVENUE ITEMS SECTION
      plSheet.getRange(`A${currentRow}`).values = [['Revenue Items']];
      const revenueSectionRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      revenueSectionRange.format.font.name = 'Times New Roman';
      revenueSectionRange.format.font.size = 12;
      revenueSectionRange.format.font.bold = true;
      revenueSectionRange.format.fill.color = ExcelFormatter.colors.darkBlue;
      revenueSectionRange.format.font.color = ExcelFormatter.colors.white;
      
      // Add thin grey underline after section header
      revenueSectionRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
      revenueSectionRange.format.borders.getItem('EdgeBottom').weight = 'Thin';
      revenueSectionRange.format.borders.getItem('EdgeBottom').color = ExcelFormatter.colors.lightGrey;
      currentRow++;
      
      // Skip one row then add DATES (two cells below Revenue Items title)
      // Set the blank row height to 8
      plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`).format.rowHeight = 8;
      currentRow++;
      
      // DATES ROW - two cells below Revenue Items title (includes Period 0)
      const dateHeaders = [''];
      const prevPeriodDate = new Date(modelData.projectStartDate);
      
      // Calculate previous period date
      switch (modelData.modelPeriods) {
        case 'daily':
          prevPeriodDate.setDate(prevPeriodDate.getDate() - 1);
          break;
        case 'monthly':
          prevPeriodDate.setMonth(prevPeriodDate.getMonth() - 1);
          break;
        case 'quarterly':
          prevPeriodDate.setMonth(prevPeriodDate.getMonth() - 3);
          break;
        case 'yearly':
          prevPeriodDate.setFullYear(prevPeriodDate.getFullYear() - 1);
          break;
      }
      
      dateHeaders.push(this.formatDateAsLastDay(prevPeriodDate, modelData.modelPeriods)); // Period 0 with XX-XXX-XX format
      const startDate = new Date(modelData.projectStartDate);
      for (let i = 0; i < periodColumns; i++) {
        const periodDate = new Date(startDate);
        
        // Calculate the date for this period
        switch (modelData.modelPeriods) {
          case 'daily':
            periodDate.setDate(periodDate.getDate() + i);
            break;
          case 'monthly':
            periodDate.setMonth(periodDate.getMonth() + i);
            break;
          case 'quarterly':
            periodDate.setMonth(periodDate.getMonth() + (i * 3));
            break;
          case 'yearly':
            periodDate.setFullYear(periodDate.getFullYear() + i);
            break;
        }
        
        dateHeaders.push(this.formatDateAsLastDay(periodDate, modelData.modelPeriods));
      }
      
      const dateRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      dateRange.values = [dateHeaders];
      dateRange.format.font.name = 'Times New Roman';
      dateRange.format.font.size = 12;
      dateRange.format.font.bold = true;
      dateRange.format.fill.color = ExcelFormatter.colors.white;
      dateRange.format.font.color = ExcelFormatter.colors.black;
      
      // Add dashed underline under dates
      dateRange.format.borders.getItem('EdgeBottom').style = 'Dash';
      dateRange.format.borders.getItem('EdgeBottom').weight = 'Thin';
      dateRange.format.borders.getItem('EdgeBottom').color = ExcelFormatter.colors.black;
      currentRow++;
      
      // PERIOD ROW - 'Period' followed by 0, then 1, 2, 3, etc.
      const periodHeaders = ['Period'];
      periodHeaders.push('0'); // Period 0
      for (let i = 1; i <= periodColumns; i++) {
        periodHeaders.push(i.toString());
      }
      
      const periodRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      periodRange.values = [periodHeaders];  
      periodRange.format.font.name = 'Times New Roman';
      periodRange.format.font.size = 12;
      periodRange.format.font.bold = false; // Remove bold formatting
      currentRow++;

      // Add revenue items with growth
      if (modelData.revenueItems && modelData.revenueItems.length > 0) {
        modelData.revenueItems.forEach((item, index) => {
          plSheet.getRange(`A${currentRow}`).values = [[item.name]];
          plSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
          plSheet.getRange(`A${currentRow}`).format.font.size = 12;
          
          const valueRef = this.cellTracker.getCellReference(`revenue_${index}`);
          const growthRateRef = this.cellTracker.getCellReference(`revenue_${index}_growth_rate`);
          
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            if (col === 1) {
              // Period 0 (Initial Investment): No revenue - show dash
              const dashRange = plSheet.getRange(`${colLetter}${currentRow}`);
              dashRange.values = [['-']];
              dashRange.format.horizontalAlignment = 'Right';
              ExcelFormatter.applyNumberFormat(dashRange);
            } else if (col === 2) {
              // First operating period: base value
              plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${valueRef}`]];
              ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
            } else {
              // Subsequent periods: apply growth
              if (growthRateRef) {
                const prevColLetter = this.getColumnLetter(col - 1);
                if (modelData.modelPeriods === 'monthly') {
                  plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef}/12)`]];
                } else if (modelData.modelPeriods === 'quarterly') {
                  plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef}/4)`]];
                } else {
                  plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef})`]];
                }
              } else {
                const prevColLetter = this.getColumnLetter(col - 1);
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}`]];
              }
              ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
            }
          }
          currentRow++;
        });
      }

      // Total Revenue
      plSheet.getRange(`A${currentRow}`).values = [['Total Revenue']];
      const totalRevRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      totalRevRange.format.font.name = 'Times New Roman';
      totalRevRange.format.font.size = 12;
      totalRevRange.format.font.bold = true;
      
      // Add thin underline at top of total
      totalRevRange.format.borders.getItem('EdgeTop').style = 'Continuous';
      totalRevRange.format.borders.getItem('EdgeTop').weight = 'Thin';
      totalRevRange.format.borders.getItem('EdgeTop').color = ExcelFormatter.colors.black;
      
      const totalRevenueRow = currentRow;
      for (let col = 1; col <= totalColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (col === 1) {
          // Period 0 - show dash
          const dashRange = plSheet.getRange(`${colLetter}${currentRow}`);
          dashRange.values = [['-']];
          dashRange.format.horizontalAlignment = 'Right';
        } else {
          const revenueStartRow = totalRevenueRow - modelData.revenueItems.length;
          const revenueEndRow = totalRevenueRow - 1;
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=SUM(${colLetter}${revenueStartRow + 1}:${colLetter}${revenueEndRow})`]];
        }
        ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
      }
      currentRow++;
      
      // Add blank row with height 8 between Total Revenue and Cost Items
      plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`).format.rowHeight = 8;
      currentRow++;

      // COST ITEMS SECTION
      plSheet.getRange(`A${currentRow}`).values = [['Cost Items']];
      const opexSectionRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      opexSectionRange.format.font.name = 'Times New Roman';
      opexSectionRange.format.font.size = 12;
      opexSectionRange.format.font.bold = true;
      opexSectionRange.format.fill.color = ExcelFormatter.colors.darkBlue;
      opexSectionRange.format.font.color = ExcelFormatter.colors.white;
      
      // Add thin grey underline after section header
      opexSectionRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
      opexSectionRange.format.borders.getItem('EdgeBottom').weight = 'Thin';
      opexSectionRange.format.borders.getItem('EdgeBottom').color = ExcelFormatter.colors.lightGrey;
      currentRow++;

      // Add operating expense items
      if (modelData.operatingExpenses && modelData.operatingExpenses.length > 0) {
        modelData.operatingExpenses.forEach((item, index) => {
          plSheet.getRange(`A${currentRow}`).values = [[item.name]];
          plSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
          plSheet.getRange(`A${currentRow}`).format.font.size = 12;
          
          const valueRef = this.cellTracker.getCellReference(`opex_${index}`);
          const growthRateRef = this.cellTracker.getCellReference(`opex_${index}_growth_rate`);
          
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            if (col === 1) {
              // Period 0 (Initial Investment): No operating expenses - show dash
              const dashRange = plSheet.getRange(`${colLetter}${currentRow}`);
              dashRange.values = [['-']];
              dashRange.format.horizontalAlignment = 'Right';
              ExcelFormatter.applyNumberFormat(dashRange);
            } else if (col === 2) {
              // First operating period: base value
              plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=-${valueRef}`]];
              ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
            } else {
              // Subsequent periods: apply growth
              if (growthRateRef) {
                const prevColLetter = this.getColumnLetter(col - 1);
                if (modelData.modelPeriods === 'monthly') {
                  plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef}/12)`]];
                } else if (modelData.modelPeriods === 'quarterly') {
                  plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef}/4)`]];
                } else {
                  plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef})`]];
                }
              } else {
                const prevColLetter = this.getColumnLetter(col - 1);
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}`]];
              }
              ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
            }
          }
          currentRow++;
        });
      }

      // Total Operating Expenses
      plSheet.getRange(`A${currentRow}`).values = [['Total Operating Expenses']];
      const totalOpExRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      totalOpExRange.format.font.name = 'Times New Roman';
      totalOpExRange.format.font.size = 12;
      totalOpExRange.format.font.bold = true;
      
      // Add thin underline at top of total
      totalOpExRange.format.borders.getItem('EdgeTop').style = 'Continuous';
      totalOpExRange.format.borders.getItem('EdgeTop').weight = 'Thin';
      totalOpExRange.format.borders.getItem('EdgeTop').color = ExcelFormatter.colors.black;
      
      const totalOpExRow = currentRow;
      for (let col = 1; col <= totalColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (col === 1) {
          // Period 0 - show dash
          const dashRange = plSheet.getRange(`${colLetter}${currentRow}`);
          dashRange.values = [['-']];
          dashRange.format.horizontalAlignment = 'Right';
        } else {
          const opexStartRow = totalOpExRow - modelData.operatingExpenses.length;
          const opexEndRow = totalOpExRow - 1;
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=SUM(${colLetter}${opexStartRow + 1}:${colLetter}${opexEndRow})`]];
        }
        ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
      }
      currentRow++;

      // NOI - moved directly below Total Operating Expenses
      plSheet.getRange(`A${currentRow}`).values = [['NOI']];
      const noiRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      noiRange.format.font.name = 'Times New Roman';
      noiRange.format.font.size = 12;
      noiRange.format.font.bold = true;
      noiRange.format.fill.color = ExcelFormatter.colors.backgroundDarker5; // Light grey background
      noiRange.format.font.color = ExcelFormatter.colors.black; // Black text
      
      // Add thin underline at top of NOI
      noiRange.format.borders.getItem('EdgeTop').style = 'Continuous';
      noiRange.format.borders.getItem('EdgeTop').weight = 'Thin';
      noiRange.format.borders.getItem('EdgeTop').color = ExcelFormatter.colors.black;
      
      const ebitdaRow = currentRow;
      for (let col = 1; col <= totalColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (col === 1) {
          // Period 0 - show dash
          const dashRange = plSheet.getRange(`${colLetter}${currentRow}`);
          dashRange.values = [['-']];
          dashRange.format.horizontalAlignment = 'Right';
        } else {
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${totalRevenueRow}+${colLetter}${totalOpExRow}`]];
        }
        ExcelFormatter.applyNumberFormat(plSheet.getRange(`${colLetter}${currentRow}`));
      }
      currentRow += 2;

      // Real estate model: NOI is the final metric - no interest, tax, or net income calculations needed
      
      // Track NOI for FCF calculations
      this.plCellTracker.recordCell('noi', 'P&L Statement', `B${ebitdaRow}:${this.getColumnLetter(totalColumns)}${ebitdaRow}`);

      // Apply number formatting with brackets for negatives and dash for zeros
      const dataRange = plSheet.getRange(`B5:${this.getColumnLetter(totalColumns)}${currentRow}`);
      ExcelFormatter.applyNumberFormat(dataRange);
      
      // Apply Times New Roman font to all data cells
      const allDataRange = plSheet.getRange(`A1:${this.getColumnLetter(totalColumns)}${currentRow}`);
      allDataRange.format.font.name = 'Times New Roman';
      allDataRange.format.font.size = 12;
      
      // Hide gridlines
      await context.sync();
      plSheet.showGridlines = false;
      
      // Auto-fit columns
      plSheet.getRange(`A:${this.getColumnLetter(totalColumns)}`).format.autofitColumns();
      
      await context.sync();
      console.log('✅ Enhanced P&L Statement created successfully with formatting');
    });
  }

  // Read existing CapEx sheet structure
  async readCapExSheetStructure() {
    return Excel.run(async (context) => {
      try {
        const sheets = context.workbook.worksheets;
        const capExSheet = sheets.getItemOrNullObject('CapEx');
        capExSheet.load('name');
        await context.sync();
        
        if (capExSheet.isNullObject) {
          console.log('⚠️ No CapEx sheet found');
          return { totalRow: null, sheet: null };
        }
        
        // Find the total row by looking for "Total CapEx" text
        const usedRange = capExSheet.getUsedRange();
        usedRange.load('values');
        await context.sync();
        
        const values = usedRange.values;
        let totalRow = null;
        
        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            if (values[row][col] === 'Total CapEx') {
              totalRow = row + 1; // Excel is 1-indexed
              break;
            }
          }
          if (totalRow) break;
        }
        
        console.log('📊 CapEx sheet structure discovered:', { totalRow, sheet: 'CapEx' });
        return { totalRow, sheet: 'CapEx' };
        
      } catch (error) {
        console.error('❌ Error reading CapEx sheet structure:', error);
        return { totalRow: null, sheet: null };
      }
    });
  }

  // Generate CapEx sheet with P&L-style formatting
  async generateCapExSheet(modelData) {
    return Excel.run(async (context) => {
      console.log('📊 Creating CapEx Summary Sheet...');
      
      const sheets = context.workbook.worksheets;
      
      // Delete existing CapEx sheet if it exists
      try {
        const existingSheet = sheets.getItemOrNullObject('CapEx');
        existingSheet.load('name');
        await context.sync();
        
        if (!existingSheet.isNullObject) {
          console.log('🗑️ Deleting existing CapEx sheet');
          existingSheet.delete();
          await context.sync();
        }
      } catch (e) {
        // Sheet doesn't exist, continue
      }
      
      // Create new CapEx sheet
      const capExSheet = sheets.add('CapEx');
      capExSheet.position = 2; // After P&L
      capExSheet.activate();
      
      // Hide gridlines
      capExSheet.showGridlines = false;
      
      // Calculate periods
      const startDate = new Date(modelData.projectStartDate);
      const endDate = new Date(modelData.projectEndDate);
      const periods = this.calculatePeriods(startDate, endDate, modelData.modelPeriods);
      const periodColumns = periods;
      const totalColumns = periodColumns + 1; // +1 for Period 0
      
      let currentRow = 1;
      
      // TITLE - P&L style formatting
      capExSheet.getRange('A1').values = [['Capital Expenditures Summary']];
      const titleRange = capExSheet.getRange(`A1:${this.getColumnLetter(totalColumns)}1`);
      titleRange.merge();
      titleRange.format.font.name = 'Times New Roman';
      titleRange.format.font.size = 12;
      titleRange.format.font.bold = true;
      titleRange.format.fill.color = ExcelFormatter.colors.backgroundDarker5;
      titleRange.format.horizontalAlignment = 'Left';
      currentRow = 2; // Remove extra blank row
      
      // CAPEX ITEMS SECTION HEADER
      capExSheet.getRange(`A${currentRow}`).values = [['CapEx Items']];
      const capExSectionRange = capExSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      capExSectionRange.format.font.name = 'Times New Roman';
      capExSectionRange.format.font.size = 12;
      capExSectionRange.format.font.bold = true;
      capExSectionRange.format.fill.color = ExcelFormatter.colors.darkBlue;
      capExSectionRange.format.font.color = ExcelFormatter.colors.white;
      
      // Add thin grey underline after section header
      capExSectionRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
      capExSectionRange.format.borders.getItem('EdgeBottom').weight = 'Thin';
      capExSectionRange.format.borders.getItem('EdgeBottom').color = ExcelFormatter.colors.lightGrey;
      currentRow++;
      
      // Skip one row then add DATES (height 8)
      capExSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`).format.rowHeight = 8;
      currentRow++;
      
      // DATES ROW - includes Period 0
      const dateHeaders = [''];
      const prevPeriodDate = new Date(modelData.projectStartDate);
      
      // Calculate previous period date
      switch (modelData.modelPeriods) {
        case 'daily':
          prevPeriodDate.setDate(prevPeriodDate.getDate() - 1);
          break;
        case 'monthly':
          prevPeriodDate.setMonth(prevPeriodDate.getMonth() - 1);
          break;
        case 'quarterly':
          prevPeriodDate.setMonth(prevPeriodDate.getMonth() - 3);
          break;
        case 'yearly':
          prevPeriodDate.setFullYear(prevPeriodDate.getFullYear() - 1);
          break;
      }
      
      dateHeaders.push(this.formatDateAsLastDay(prevPeriodDate, modelData.modelPeriods)); // Period 0
      for (let i = 0; i < periodColumns; i++) {
        const periodDate = new Date(startDate);
        
        switch (modelData.modelPeriods) {
          case 'daily':
            periodDate.setDate(periodDate.getDate() + i);
            break;
          case 'monthly':
            periodDate.setMonth(periodDate.getMonth() + i);
            break;
          case 'quarterly':
            periodDate.setMonth(periodDate.getMonth() + (i * 3));
            break;
          case 'yearly':
            periodDate.setFullYear(periodDate.getFullYear() + i);
            break;
        }
        
        dateHeaders.push(this.formatDateAsLastDay(periodDate, modelData.modelPeriods));
      }
      
      const dateRange = capExSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      dateRange.values = [dateHeaders];
      dateRange.format.font.name = 'Times New Roman';
      dateRange.format.font.size = 12;
      dateRange.format.font.bold = true;
      dateRange.format.fill.color = ExcelFormatter.colors.white;
      dateRange.format.font.color = ExcelFormatter.colors.black;
      
      // Add dashed underline under dates
      dateRange.format.borders.getItem('EdgeBottom').style = 'Dash';
      dateRange.format.borders.getItem('EdgeBottom').weight = 'Thin';
      dateRange.format.borders.getItem('EdgeBottom').color = ExcelFormatter.colors.black;
      currentRow++;
      
      // PERIOD ROW
      const periodHeaders = ['Period'];
      periodHeaders.push('0'); // Period 0
      for (let i = 1; i <= periodColumns; i++) {
        periodHeaders.push(i.toString());
      }
      
      const periodRange = capExSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      periodRange.values = [periodHeaders];
      periodRange.format.font.name = 'Times New Roman';
      periodRange.format.font.size = 12;
      periodRange.format.font.bold = false;
      currentRow++;
      
      // Individual CapEx items
      const capExStructure = { items: [], totalRow: null };
      
      if (modelData.capEx && modelData.capEx.length > 0) {
        modelData.capEx.forEach((item, index) => {
          capExSheet.getRange(`A${currentRow}`).values = [[item.name || `CapEx ${index + 1}`]];
          capExSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
          capExSheet.getRange(`A${currentRow}`).format.font.size = 12;
          
          const valueRef = this.cellTracker.getCellReference(`capex_${index}`);
          const growthRateRef = this.cellTracker.getCellReference(`capex_${index}_growth_rate`);
          
          capExStructure.items.push({
            name: item.name,
            row: currentRow,
            valueRef: valueRef,
            growthRateRef: growthRateRef
          });
          
          // Fill periods with data
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            if (col === 1) {
              // Period 0: No CapEx - show dash
              const dashRange = capExSheet.getRange(`${colLetter}${currentRow}`);
              dashRange.values = [['-']];
              dashRange.format.horizontalAlignment = 'Right';
              ExcelFormatter.applyNumberFormat(dashRange);
            } else if (col === 2) {
              // Period 1: Base CapEx value (negative cash flow)
              if (valueRef) {
                capExSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=-${valueRef}`]];
                ExcelFormatter.applyNumberFormat(capExSheet.getRange(`${colLetter}${currentRow}`));
              }
            } else {
              // Subsequent periods: Apply growth
              if (valueRef && growthRateRef) {
                const prevColLetter = this.getColumnLetter(col - 1);
                if (modelData.modelPeriods === 'monthly') {
                  capExSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef}/12)`]];
                } else if (modelData.modelPeriods === 'quarterly') {
                  capExSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef}/4)`]];
                } else if (modelData.modelPeriods === 'daily') {
                  capExSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef}/365)`]];
                } else {
                  capExSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef})`]];
                }
                ExcelFormatter.applyNumberFormat(capExSheet.getRange(`${colLetter}${currentRow}`));
              }
            }
          }
          currentRow++;
        });
        
        // Total CapEx row
        capExSheet.getRange(`A${currentRow}`).values = [['Total CapEx']];
        const totalCapExRange = capExSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
        totalCapExRange.format.font.name = 'Times New Roman';
        totalCapExRange.format.font.size = 12;
        totalCapExRange.format.font.bold = true;
        
        // Add thin underline at top of total and light grey background
        totalCapExRange.format.borders.getItem('EdgeTop').style = 'Continuous';
        totalCapExRange.format.borders.getItem('EdgeTop').weight = 'Thin';
        totalCapExRange.format.borders.getItem('EdgeTop').color = ExcelFormatter.colors.black;
        totalCapExRange.format.fill.color = ExcelFormatter.colors.backgroundDarker5;
        
        capExStructure.totalRow = currentRow;
        
        for (let col = 1; col <= totalColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          if (col === 1) {
            // Period 0 - show dash
            const dashRange = capExSheet.getRange(`${colLetter}${currentRow}`);
            dashRange.values = [['-']];
            dashRange.format.horizontalAlignment = 'Right';
            ExcelFormatter.applyNumberFormat(dashRange);
          } else {
            const startRow = currentRow - modelData.capEx.length;
            const endRow = currentRow - 1;
            capExSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=SUM(${colLetter}${startRow}:${colLetter}${endRow})`]];
            ExcelFormatter.applyNumberFormat(capExSheet.getRange(`${colLetter}${currentRow}`));
          }
        }
        
        // Remove thick bottom underline (only keep top underline)
        
      } else {
        // No CapEx items
        capExSheet.getRange(`A${currentRow}`).values = [['No CapEx Items']];
        capExSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
        capExSheet.getRange(`A${currentRow}`).format.font.size = 12;
        capExStructure.totalRow = currentRow;
        
        for (let col = 1; col <= totalColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          const dashRange = capExSheet.getRange(`${colLetter}${currentRow}`);
          dashRange.values = [['-']];
          dashRange.format.horizontalAlignment = 'Right';
          ExcelFormatter.applyNumberFormat(dashRange);
        }
      }
      
      // Auto-resize columns
      capExSheet.getUsedRange().format.autofitColumns();
      
      // Ensure all cells use Times New Roman font
      const allCapExCellsRange = capExSheet.getUsedRange();
      allCapExCellsRange.format.font.name = 'Times New Roman';
      allCapExCellsRange.format.font.size = 12;
      
      console.log('✅ CapEx Summary Sheet created successfully!');
      return capExStructure;
    });
  }

  // Create FCF sheet with the new structure
  async createAIFCFSheet(modelData, fcfPrompt, plStructure, assumptionStructure, capExStructure) {
    return Excel.run(async (context) => {
      console.log('💰 Creating Free Cash Flow Sheet...');
      
      const sheets = context.workbook.worksheets;
      
      // Delete existing FCF sheet if it exists
      try {
        const existingSheet = sheets.getItemOrNullObject('FCF');
        existingSheet.load('name');
        await context.sync();
        
        if (!existingSheet.isNullObject) {
          console.log('🗑️ Deleting existing FCF sheet');
          existingSheet.delete();
          await context.sync();
        }
      } catch (e) {
        // Sheet doesn't exist, continue
      }
      
      // Create new FCF sheet
      const fcfSheet = sheets.add('FCF');
      fcfSheet.position = 3; // After CapEx
      fcfSheet.activate();
      
      // Hide gridlines
      fcfSheet.showGridlines = false;
      
      // Calculate periods
      const startDate = new Date(modelData.projectStartDate);
      const endDate = new Date(modelData.projectEndDate);
      const periods = this.calculatePeriods(startDate, endDate, modelData.modelPeriods);
      const totalColumns = periods + 1; // +1 for Period 0 (periods is a number, not array)
      
      // TITLE - P&L style formatting
      fcfSheet.getRange('A1').values = [['Cashflows']];
      const titleRange = fcfSheet.getRange(`A1:${this.getColumnLetter(totalColumns)}1`);
      titleRange.merge();
      titleRange.format.font.name = 'Times New Roman';
      titleRange.format.font.size = 12;
      titleRange.format.font.bold = true;
      titleRange.format.fill.color = ExcelFormatter.colors.backgroundDarker5;
      titleRange.format.horizontalAlignment = 'Left';
      titleRange.format.font.color = ExcelFormatter.colors.black;
      
      let currentRow = 2; // Remove extra blank row
      
      // Skip one row then add DATES (height 8)
      fcfSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`).format.rowHeight = 8;
      currentRow++;
      
      // DATES ROW - includes Period 0
      const dateHeaders = [''];
      const prevPeriodDate = new Date(modelData.projectStartDate);
      
      // Calculate previous period date
      switch (modelData.modelPeriods) {
        case 'daily':
          prevPeriodDate.setDate(prevPeriodDate.getDate() - 1);
          break;
        case 'monthly':
          prevPeriodDate.setMonth(prevPeriodDate.getMonth() - 1);
          break;
        case 'quarterly':
          prevPeriodDate.setMonth(prevPeriodDate.getMonth() - 3);
          break;
        case 'yearly':
          prevPeriodDate.setFullYear(prevPeriodDate.getFullYear() - 1);
          break;
      }
      
      dateHeaders.push(this.formatDateAsLastDay(prevPeriodDate, modelData.modelPeriods)); // Period 0
      for (let i = 0; i < periods; i++) {
        const periodDate = new Date(startDate);
        
        switch (modelData.modelPeriods) {
          case 'daily':
            periodDate.setDate(periodDate.getDate() + i);
            break;
          case 'monthly':
            periodDate.setMonth(periodDate.getMonth() + i);
            break;
          case 'quarterly':
            periodDate.setMonth(periodDate.getMonth() + (i * 3));
            break;
          case 'yearly':
            periodDate.setFullYear(periodDate.getFullYear() + i);
            break;
        }
        
        dateHeaders.push(this.formatDateAsLastDay(periodDate, modelData.modelPeriods));
      }
      
      const dateRange = fcfSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      dateRange.values = [dateHeaders];
      dateRange.format.font.name = 'Times New Roman';
      dateRange.format.font.size = 12;
      dateRange.format.font.bold = true;
      dateRange.format.fill.color = ExcelFormatter.colors.white;
      dateRange.format.font.color = ExcelFormatter.colors.black;
      
      // Add dashed underline under dates
      dateRange.format.borders.getItem('EdgeBottom').style = 'Dash';
      dateRange.format.borders.getItem('EdgeBottom').weight = 'Thin';
      dateRange.format.borders.getItem('EdgeBottom').color = ExcelFormatter.colors.black;
      currentRow++;
      
      // PERIOD ROW
      const periodHeaders = ['Period'];
      periodHeaders.push('0'); // Period 0
      for (let i = 1; i <= periods; i++) {
        periodHeaders.push(i.toString());
      }
      
      const periodRange = fcfSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      periodRange.values = [periodHeaders];
      periodRange.format.font.name = 'Times New Roman';
      periodRange.format.font.size = 12;
      periodRange.format.font.bold = false;
      periodRange.format.font.color = ExcelFormatter.colors.black;
      currentRow++;
      
      // Remove section headers - start directly with data rows
      
      // Purchase price (Period 0 only)
      fcfSheet.getRange(`A${currentRow}`).values = [['Purchase price']];
      fcfSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
      fcfSheet.getRange(`A${currentRow}`).format.font.size = 12;
      fcfSheet.getRange(`A${currentRow}`).format.font.color = ExcelFormatter.colors.black;
      
      fcfSheet.getRange('B' + currentRow).formulas = [[`=-Assumptions!${this.cellTracker.getCellReference('dealValue')?.split('!')[1] || 'B10'}`]];
      ExcelFormatter.applyNumberFormat(fcfSheet.getRange('B' + currentRow));
      for (let i = 1; i <= periods; i++) {
        const colLetter = this.getColumnLetter(i + 1);
        const dashRange = fcfSheet.getRange(colLetter + currentRow);
        dashRange.values = [['-']];
        dashRange.format.horizontalAlignment = 'Right';
        ExcelFormatter.applyNumberFormat(dashRange);
      }
      currentRow++;
      
      // Transaction costs (Period 0 only)
      fcfSheet.getRange(`A${currentRow}`).values = [['Transaction costs']];
      fcfSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
      fcfSheet.getRange(`A${currentRow}`).format.font.size = 12;
      fcfSheet.getRange(`A${currentRow}`).format.font.color = ExcelFormatter.colors.black;
      
      const dealValueRef = this.cellTracker.getCellReference('dealValue');
      const transactionFeeRef = this.cellTracker.getCellReference('transactionFee');
      if (dealValueRef && transactionFeeRef) {
        fcfSheet.getRange('B' + currentRow).formulas = [[`=-${dealValueRef}*${transactionFeeRef}`]];
        ExcelFormatter.applyNumberFormat(fcfSheet.getRange('B' + currentRow));
      }
      for (let i = 1; i <= periods; i++) {
        const colLetter = this.getColumnLetter(i + 1);
        const dashRange = fcfSheet.getRange(colLetter + currentRow);
        dashRange.values = [['-']];
        dashRange.format.horizontalAlignment = 'Right';
        ExcelFormatter.applyNumberFormat(dashRange);
      }
      currentRow++;
      
      // NOI from P&L
      fcfSheet.getRange(`A${currentRow}`).values = [['NOI']];
      fcfSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
      fcfSheet.getRange(`A${currentRow}`).format.font.size = 12;
      fcfSheet.getRange(`A${currentRow}`).format.font.color = ExcelFormatter.colors.black;
      
      if (plStructure.lineItems.noi) {
        for (let i = 0; i <= periods; i++) {
          const colLetter = this.getColumnLetter(i + 1); // FCF column (Period i is in column i+1)
          const plCol = this.getColumnLetter(i + 1); // P&L column mapping: Period 0 FCF = Period 1 P&L
          fcfSheet.getRange(colLetter + currentRow).formulas = [[`='P&L Statement'!${plCol}${plStructure.lineItems.noi.row}`]];
          ExcelFormatter.applyNumberFormat(fcfSheet.getRange(colLetter + currentRow));
        }
      } else if (plStructure.lineItems.netIncome) {
        // Fallback: Use Net Income if NOI not found
        for (let i = 1; i <= periods; i++) {
          const colLetter = this.getColumnLetter(i + 1);
          const plCol = this.getColumnLetter(i);
          fcfSheet.getRange(colLetter + currentRow).formulas = [[`='P&L Statement'!${plCol}${plStructure.lineItems.netIncome.row}`]];
          ExcelFormatter.applyNumberFormat(fcfSheet.getRange(colLetter + currentRow));
        }
      } else {
        // Last fallback: Use zeros and log warning
        for (let i = 1; i <= periods; i++) {
          const colLetter = this.getColumnLetter(i + 1);
          fcfSheet.getRange(colLetter + currentRow).values = [[0]];
          ExcelFormatter.applyNumberFormat(fcfSheet.getRange(colLetter + currentRow));
        }
        console.warn('⚠️ NOI not found in P&L structure for operating cash flows');
      }
      currentRow++;
      
      // CapEx (from CapEx sheet)
      fcfSheet.getRange(`A${currentRow}`).values = [['CapEx']];
      fcfSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
      fcfSheet.getRange(`A${currentRow}`).format.font.size = 12;
      fcfSheet.getRange(`A${currentRow}`).format.font.color = ExcelFormatter.colors.black;
      
      if (capExStructure && capExStructure.totalRow) {
        for (let i = 0; i <= periods; i++) {
          const colLetter = this.getColumnLetter(i + 1); // FCF column (B, C, D, E...)
          const capExCol = this.getColumnLetter(i + 1); // CapEx column (B, C, D, E...) - same mapping now
          fcfSheet.getRange(colLetter + currentRow).formulas = [[`='CapEx'!${capExCol}${capExStructure?.totalRow}`]];
          ExcelFormatter.applyNumberFormat(fcfSheet.getRange(colLetter + currentRow));
        }
      }
      currentRow++;
      
      // Sale Price (final period only)
      fcfSheet.getRange(`A${currentRow}`).values = [['Sale Price']];
      fcfSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
      fcfSheet.getRange(`A${currentRow}`).format.font.size = 12;
      fcfSheet.getRange(`A${currentRow}`).format.font.color = ExcelFormatter.colors.black;
      
      // Fill Period 0 through Period N-1 with dashes
      for (let i = 0; i < periods; i++) {
        const colLetter = this.getColumnLetter(i + 1);
        const dashRange = fcfSheet.getRange(colLetter + currentRow);
        dashRange.values = [['-']];
        dashRange.format.horizontalAlignment = 'Right';
        ExcelFormatter.applyNumberFormat(dashRange);
      }
      
      const finalPeriodCol = this.getColumnLetter(periods + 1); // FIXED: Correct final period column
      if (plStructure.lineItems.noi) {
        // Terminal value = Final NOI / Terminal Cap Rate
        const terminalCapRateRef = this.cellTracker.getCellReference('terminalCapRate');
        if (terminalCapRateRef) {
          fcfSheet.getRange(finalPeriodCol + currentRow).formulas = [[`='P&L Statement'!${this.getColumnLetter(periods + 1)}${plStructure.lineItems.noi.row}/${terminalCapRateRef}`]];
          ExcelFormatter.applyNumberFormat(fcfSheet.getRange(finalPeriodCol + currentRow));
        }
      } else if (plStructure.lineItems.netIncome) {
        // Fallback: Use Net Income if NOI not found
        const terminalCapRateRef = this.cellTracker.getCellReference('terminalCapRate');
        if (terminalCapRateRef) {
          fcfSheet.getRange(finalPeriodCol + currentRow).formulas = [[`='P&L Statement'!${this.getColumnLetter(periods + 1)}${plStructure.lineItems.netIncome.row}/${terminalCapRateRef}`]];
          ExcelFormatter.applyNumberFormat(fcfSheet.getRange(finalPeriodCol + currentRow));
        }
      } else {
        // Last fallback: If no NOI or Net Income found, use a placeholder note
        fcfSheet.getRange(finalPeriodCol + currentRow).values = [['Check P&L for NOI']];
        console.warn('⚠️ NOI not found in P&L structure for terminal value calculation');
      }
      currentRow++;
      
      // Disposal Costs (final period only)
      fcfSheet.getRange(`A${currentRow}`).values = [['Disposal Costs']];
      fcfSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
      fcfSheet.getRange(`A${currentRow}`).format.font.size = 12;
      fcfSheet.getRange(`A${currentRow}`).format.font.color = ExcelFormatter.colors.black;
      
      // Fill Period 0 through Period N-1 with dashes
      for (let i = 0; i < periods; i++) {
        const colLetter = this.getColumnLetter(i + 1);
        const dashRange = fcfSheet.getRange(colLetter + currentRow);
        dashRange.values = [['-']];
        dashRange.format.horizontalAlignment = 'Right';
        ExcelFormatter.applyNumberFormat(dashRange);
      }
      
      const disposalCostRef = this.cellTracker.getCellReference('disposalCost');
      if (disposalCostRef) {
        fcfSheet.getRange(finalPeriodCol + currentRow).formulas = [[`=${finalPeriodCol}${currentRow - 1}*${disposalCostRef}`]];
        ExcelFormatter.applyNumberFormat(fcfSheet.getRange(finalPeriodCol + currentRow));
      }
      currentRow++;
      
      // Unlevered Cashflows (sum of all above)
      fcfSheet.getRange(`A${currentRow}`).values = [['Unlevered Cashflows']];
      const unleverCashflowRange = fcfSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      unleverCashflowRange.format.font.name = 'Times New Roman';
      unleverCashflowRange.format.font.size = 12;
      unleverCashflowRange.format.font.color = ExcelFormatter.colors.black;
      
      // Add thin underline at top of unlevered cashflow row
      unleverCashflowRange.format.borders.getItem('EdgeTop').style = 'Continuous';
      unleverCashflowRange.format.borders.getItem('EdgeTop').weight = 'Thin';
      unleverCashflowRange.format.borders.getItem('EdgeTop').color = ExcelFormatter.colors.black;
      
      const unlevereCashflowsRow = currentRow;
      for (let i = 0; i <= periods; i++) {
        const colLetter = this.getColumnLetter(i + 1);
        fcfSheet.getRange(colLetter + currentRow).formulas = [[`=SUM(${colLetter}${currentRow - 6}:${colLetter}${currentRow - 1})`]];
        ExcelFormatter.applyNumberFormat(fcfSheet.getRange(colLetter + currentRow));
      }
      currentRow++;
      
      // Add empty row between Unlevered cashflows and debt upfront costs as requested
      fcfSheet.getRange(`A${currentRow}`).format.rowHeight = 8;
      currentRow++;
      
      // Get debt financing reference for later use
      const debtFinancingRef = this.cellTracker.getCellReference('debtFinancing');
      
      // Debt upfront costs (Period 0 only)
      fcfSheet.getRange(`A${currentRow}`).values = [['Debt upfront costs']];
      fcfSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
      fcfSheet.getRange(`A${currentRow}`).format.font.size = 12;
      fcfSheet.getRange(`A${currentRow}`).format.font.color = ExcelFormatter.colors.black;
      
      const loanIssuanceFeesRef = this.cellTracker.getCellReference('loanIssuanceFees');
      if (debtFinancingRef && loanIssuanceFeesRef) {
        // Reference loan issuance fees from Assumptions sheet
        fcfSheet.getRange('B' + currentRow).formulas = [[`=-${debtFinancingRef}*${loanIssuanceFeesRef}`]];
        ExcelFormatter.applyNumberFormat(fcfSheet.getRange('B' + currentRow));
      } else if (debtFinancingRef) {
        // Fallback if loan issuance fees not found
        fcfSheet.getRange('B' + currentRow).formulas = [[`=-${debtFinancingRef}*1.5/100`]];
        ExcelFormatter.applyNumberFormat(fcfSheet.getRange('B' + currentRow));
      }
      
      // Fill Period 1 through Period N with dashes
      for (let i = 1; i <= periods; i++) {
        const colLetter = this.getColumnLetter(i + 1);
        const dashRange = fcfSheet.getRange(colLetter + currentRow);
        dashRange.values = [['-']];
        dashRange.format.horizontalAlignment = 'Right';
        ExcelFormatter.applyNumberFormat(dashRange);
      }
      currentRow++;
      
      // Debt Expense (referenced directly from Debt Model)
      fcfSheet.getRange(`A${currentRow}`).values = [['Debt Expense']];
      fcfSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
      fcfSheet.getRange(`A${currentRow}`).format.font.size = 12;
      fcfSheet.getRange(`A${currentRow}`).format.font.color = ExcelFormatter.colors.black;
      
      if (modelData.dealLTV && parseFloat(modelData.dealLTV) > 0) {
        // Reference debt expense directly from the Debt Model sheet row 8
        for (let i = 0; i <= periods; i++) {
          const colLetter = this.getColumnLetter(i + 1); // FCF column
          if (i === 0) {
            const dashRange = fcfSheet.getRange(colLetter + currentRow);
            dashRange.values = [['-']];
            dashRange.format.horizontalAlignment = 'Right';
            ExcelFormatter.applyNumberFormat(dashRange);
          } else {
            // Reference debt expense from Debt Model: Period 1 is column C, Period 2 is column D, etc.
            const debtModelCol = this.getColumnLetter(i + 1); // FCF Period i maps to Debt Model column (i+1)
            fcfSheet.getRange(colLetter + currentRow).formulas = [[`=-'Debt Model'!${debtModelCol}8`]];
            ExcelFormatter.applyNumberFormat(fcfSheet.getRange(colLetter + currentRow));
          }
        }
      } else {
        // No debt - all zeros
        for (let i = 0; i <= periods; i++) {
          const colLetter = this.getColumnLetter(i + 1);
          const dashRange = fcfSheet.getRange(colLetter + currentRow);
          dashRange.values = [['-']];
          dashRange.format.horizontalAlignment = 'Right';
          ExcelFormatter.applyNumberFormat(dashRange);
        }
      }
      currentRow++;
      
      // Loan proceeds (positive cash inflow in Period 0)
      fcfSheet.getRange(`A${currentRow}`).values = [['Loan proceeds']];
      fcfSheet.getRange(`A${currentRow}`).format.font.name = 'Times New Roman';
      fcfSheet.getRange(`A${currentRow}`).format.font.size = 12;
      fcfSheet.getRange(`A${currentRow}`).format.font.color = ExcelFormatter.colors.black;
      
      if (modelData.dealLTV && parseFloat(modelData.dealLTV) > 0) {
        const debtFinancingRef = this.cellTracker.getCellReference('debtFinancing');
        for (let i = 0; i <= periods; i++) {
          const colLetter = this.getColumnLetter(i + 1); // FCF column
          if (i === 0) {
            // Initial period - loan proceeds come in
            if (debtFinancingRef) {
              fcfSheet.getRange(colLetter + currentRow).formulas = [[`=${debtFinancingRef}`]];
              ExcelFormatter.applyNumberFormat(fcfSheet.getRange(colLetter + currentRow));
            } else {
              // Fallback calculation
              const dealValue = parseFloat(modelData.dealValue) || 0;
              const ltvRatio = parseFloat(modelData.dealLTV) / 100 || 0;
              const loanAmount = dealValue * ltvRatio;
              fcfSheet.getRange(colLetter + currentRow).values = [[loanAmount]];
              ExcelFormatter.applyNumberFormat(fcfSheet.getRange(colLetter + currentRow));
            }
          } else {
            // All other periods - no loan proceeds
            const dashRange = fcfSheet.getRange(colLetter + currentRow);
            dashRange.values = [['-']];
            dashRange.format.horizontalAlignment = 'Right';
            ExcelFormatter.applyNumberFormat(dashRange);
          }
        }
      } else {
        // No debt - all zeros
        for (let i = 0; i <= periods; i++) {
          const colLetter = this.getColumnLetter(i + 1);
          const dashRange = fcfSheet.getRange(colLetter + currentRow);
          dashRange.values = [['-']];
          dashRange.format.horizontalAlignment = 'Right';
          ExcelFormatter.applyNumberFormat(dashRange);
        }
      }
      currentRow++;
      
      // Levered Cashflows
      fcfSheet.getRange(`A${currentRow}`).values = [['Levered Cashflows']];
      const leveredCashflowRange = fcfSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      leveredCashflowRange.format.font.name = 'Times New Roman';
      leveredCashflowRange.format.font.size = 12;
      leveredCashflowRange.format.font.color = ExcelFormatter.colors.black;
      leveredCashflowRange.format.fill.color = ExcelFormatter.colors.backgroundDarker5;
      
      // Make only the label in column A bold, not the values
      const leveredCashflowLabelRange = fcfSheet.getRange(`A${currentRow}`);
      leveredCashflowLabelRange.format.font.bold = true;
      
      // Add thin underline at top of levered cashflow row
      leveredCashflowRange.format.borders.getItem('EdgeTop').style = 'Continuous';
      leveredCashflowRange.format.borders.getItem('EdgeTop').weight = 'Thin';
      leveredCashflowRange.format.borders.getItem('EdgeTop').color = ExcelFormatter.colors.black;
      
      const leveredCashflowsRow = currentRow;
      // Store row numbers for correct formula references (adjusted for empty row)
      const debtUpfrontCostsRow = currentRow - 4;
      const debtExpenseRow = currentRow - 3;
      const loanProceedsRow = currentRow - 2;
      
      for (let i = 0; i <= periods; i++) {
        const colLetter = this.getColumnLetter(i + 1);
        // Levered CF = Unlevered CF + Debt upfront costs + Debt Expense + Loan proceeds (using SUM)
        fcfSheet.getRange(colLetter + currentRow).formulas = [[`=SUM(${colLetter}${unlevereCashflowsRow},${colLetter}${debtUpfrontCostsRow},${colLetter}${debtExpenseRow},${colLetter}${loanProceedsRow})`]];
        ExcelFormatter.applyNumberFormat(fcfSheet.getRange(colLetter + currentRow));
      }
      currentRow += 2;
      
      // Equity distributions (Levered Cashflows) - removed EQUITY FLOWS header as requested
      fcfSheet.getRange(`A${currentRow}`).values = [['Equity distributions']];
      const equityDistRange = fcfSheet.getRange(`A${currentRow}`);
      equityDistRange.format.font.name = 'Times New Roman';
      equityDistRange.format.font.size = 12;
      equityDistRange.format.font.color = ExcelFormatter.colors.black;
      
      for (let i = 0; i <= periods; i++) {
        const colLetter = this.getColumnLetter(i + 1);
        fcfSheet.getRange(colLetter + currentRow).formulas = [[`=${colLetter}${leveredCashflowsRow}`]];
        ExcelFormatter.applyNumberFormat(fcfSheet.getRange(colLetter + currentRow));
      }
      currentRow += 2;
      
      // Use proper column range: B to final period column
      const finalCol = this.getColumnLetter(periods + 1);
      
      // Unlevered IRR - dark blue background only around text and value as requested
      fcfSheet.getRange(`A${currentRow}`).values = [['Unlevered IRR']];
      const unleverIRRLabelRange = fcfSheet.getRange(`A${currentRow}:B${currentRow}`);
      unleverIRRLabelRange.format.font.name = 'Times New Roman';
      unleverIRRLabelRange.format.font.size = 12;
      unleverIRRLabelRange.format.font.bold = true;
      unleverIRRLabelRange.format.font.color = ExcelFormatter.colors.white;
      unleverIRRLabelRange.format.fill.color = ExcelFormatter.colors.darkBlue;
      
      // Add 50% starting guess for high IRR calculations
      fcfSheet.getRange('B' + currentRow).formulas = [[`=IFERROR(IRR(B${unlevereCashflowsRow}:${finalCol}${unlevereCashflowsRow}),"No Solution")`]];
      fcfSheet.getRange('B' + currentRow).numberFormat = [['0.00%']];
      fcfSheet.getRange('B' + currentRow).format.font.bold = true;
      currentRow++;
      
      // Levered IRR - dark blue background only around text and value as requested
      fcfSheet.getRange(`A${currentRow}`).values = [['Levered IRR']];
      const leverIRRLabelRange = fcfSheet.getRange(`A${currentRow}:B${currentRow}`);
      leverIRRLabelRange.format.font.name = 'Times New Roman';
      leverIRRLabelRange.format.font.size = 12;
      leverIRRLabelRange.format.font.bold = true;
      leverIRRLabelRange.format.font.color = ExcelFormatter.colors.white;
      leverIRRLabelRange.format.fill.color = ExcelFormatter.colors.darkBlue;
      
      // Add 100% starting guess for very high IRR calculations
      fcfSheet.getRange('B' + currentRow).formulas = [[`=IFERROR(IRR(B${leveredCashflowsRow}:${finalCol}${leveredCashflowsRow}),"No Solution")`]];
      fcfSheet.getRange('B' + currentRow).numberFormat = [['0.00%']];
      fcfSheet.getRange('B' + currentRow).format.font.bold = true;
      currentRow++;
      
      // MOIC - dark blue background only around text and value as requested
      fcfSheet.getRange(`A${currentRow}`).values = [['MOIC']];
      const moicLabelRange = fcfSheet.getRange(`A${currentRow}:B${currentRow}`);
      moicLabelRange.format.font.name = 'Times New Roman';
      moicLabelRange.format.font.size = 12;
      moicLabelRange.format.font.bold = true;
      moicLabelRange.format.font.color = ExcelFormatter.colors.white;
      moicLabelRange.format.fill.color = ExcelFormatter.colors.darkBlue;
      
      // MOIC: Sum of levered cash inflows / Initial equity investment
      fcfSheet.getRange('B' + currentRow).formulas = [[`=SUM(C${leveredCashflowsRow}:${finalCol}${leveredCashflowsRow})/ABS(B${leveredCashflowsRow})`]];
      fcfSheet.getRange('B' + currentRow).numberFormat = [['0.0"x"']];
      fcfSheet.getRange('B' + currentRow).format.font.bold = true;
      
      // Format all numbers with standard negative format (excluding IRR/MOIC rows)
      // Stop formatting before the RETURNS section to preserve percentage formatting
      const dataRange = fcfSheet.getRange(`B5:${finalCol}${currentRow - 4}`); // Exclude RETURNS, IRR, and MOIC rows
      dataRange.numberFormat = [['#,##0;[Red](#,##0);"-"']];
      
      // Auto-resize columns
      fcfSheet.getUsedRange().format.autofitColumns();
      
      // Ensure all cells use Times New Roman font
      const allCellsRange = fcfSheet.getUsedRange();
      allCellsRange.format.font.name = 'Times New Roman';
      allCellsRange.format.font.size = 12;
      
      console.log('✅ Free Cash Flow Sheet created successfully!');
    });
  }

  // Generate simplified Debt Model sheet (Fixed Rate Only)
  async generateDebtModelSheet(modelData) {
    return Excel.run(async (context) => {
      console.log('🏦 Creating Debt Model sheet with P&L-style formatting...');
      
      const sheets = context.workbook.worksheets;
      
      // Delete existing sheet if it exists
      try {
        const existingSheet = sheets.getItemOrNullObject('Debt Model');
        existingSheet.load('name');
        await context.sync();
        
        if (!existingSheet.isNullObject) {
          existingSheet.delete();
          await context.sync();
        }
      } catch (e) {
        // Continue if sheet doesn't exist
      }
      
      // Create new sheet
      const debtSheet = sheets.add('Debt Model');
      await context.sync();
      
      // Hide gridlines
      debtSheet.showGridlines = false;
      
      // Check if debt financing is enabled
      const hasDebt = modelData.dealLTV && parseFloat(modelData.dealLTV) > 0;
      
      if (!hasDebt) {
        // Title for no debt scenario
        debtSheet.getRange('A1').values = [['Debt Model']];
        const titleRange = debtSheet.getRange('A1:F1');
        titleRange.merge();
        titleRange.format.font.name = 'Times New Roman';
        titleRange.format.font.size = 12;
        titleRange.format.font.bold = true;
        titleRange.format.fill.color = ExcelFormatter.colors.backgroundDarker5;
        titleRange.format.horizontalAlignment = 'Left';
        
        debtSheet.getRange('A3').values = [['No debt financing (LTV = 0%)']];
        debtSheet.getRange('A3').format.font.name = 'Times New Roman';
        debtSheet.getRange('A3').format.font.size = 12;
        await context.sync();
        return { success: true, message: 'Debt Model sheet created (no debt)' };
      }
      
      // P&L-style title formatting
      const periods = this.calculatePeriods(modelData.projectStartDate, modelData.projectEndDate, modelData.modelPeriods);
      const totalColumns = periods + 1; // +1 for Period 0
      
      debtSheet.getRange('A1').values = [['Debt Model']];
      const titleRange = debtSheet.getRange(`A1:${this.getColumnLetter(totalColumns)}1`);
      titleRange.merge();
      titleRange.format.font.name = 'Times New Roman';
      titleRange.format.font.size = 12;
      titleRange.format.font.bold = true;
      titleRange.format.fill.color = ExcelFormatter.colors.backgroundDarker5;
      titleRange.format.horizontalAlignment = 'Left';
      
      // Add section header
      debtSheet.getRange('A2').values = [['Debt Components']];
      const debtSectionRange = debtSheet.getRange(`A2:${this.getColumnLetter(totalColumns)}2`);
      debtSectionRange.format.font.name = 'Times New Roman';
      debtSectionRange.format.font.size = 12;
      debtSectionRange.format.font.bold = true;
      debtSectionRange.format.fill.color = ExcelFormatter.colors.darkBlue;
      debtSectionRange.format.font.color = ExcelFormatter.colors.white;
      
      // Set column B width to 10
      debtSheet.getRange('B:B').format.columnWidth = 10;
      
      // Skip one row then add DATES (height 8) - row 3
      debtSheet.getRange(`A3:${this.getColumnLetter(totalColumns)}3`).format.rowHeight = 8;
      
      // DATES ROW - row 4 (includes Period 0)
      const dateHeaders = [''];
      const prevPeriodDate = new Date(modelData.projectStartDate);
      
      // Calculate previous period date
      switch (modelData.modelPeriods) {
        case 'daily':
          prevPeriodDate.setDate(prevPeriodDate.getDate() - 1);
          break;
        case 'monthly':
          prevPeriodDate.setMonth(prevPeriodDate.getMonth() - 1);
          break;
        case 'quarterly':
          prevPeriodDate.setMonth(prevPeriodDate.getMonth() - 3);
          break;
        case 'yearly':
          prevPeriodDate.setFullYear(prevPeriodDate.getFullYear() - 1);
          break;
      }
      
      dateHeaders.push(this.formatDateAsLastDay(prevPeriodDate, modelData.modelPeriods)); // Period 0
      const startDate = new Date(modelData.projectStartDate);
      for (let i = 0; i < periods; i++) {
        const periodDate = new Date(startDate);
        
        switch (modelData.modelPeriods) {
          case 'daily':
            periodDate.setDate(periodDate.getDate() + i);
            break;
          case 'monthly':
            periodDate.setMonth(periodDate.getMonth() + i);
            break;
          case 'quarterly':
            periodDate.setMonth(periodDate.getMonth() + (i * 3));
            break;
          case 'yearly':
            periodDate.setFullYear(periodDate.getFullYear() + i);
            break;
        }
        
        dateHeaders.push(this.formatDateAsLastDay(periodDate, modelData.modelPeriods));
      }
      
      const dateRange = debtSheet.getRange(`A4:${this.getColumnLetter(totalColumns)}4`);
      dateRange.values = [dateHeaders];
      dateRange.format.font.name = 'Times New Roman';
      dateRange.format.font.size = 12;
      dateRange.format.font.bold = true;
      dateRange.format.fill.color = ExcelFormatter.colors.white;
      dateRange.format.font.color = ExcelFormatter.colors.black;
      
      // Add dashed underline under dates
      dateRange.format.borders.getItem('EdgeBottom').style = 'Dash';
      dateRange.format.borders.getItem('EdgeBottom').weight = 'Thin';
      dateRange.format.borders.getItem('EdgeBottom').color = ExcelFormatter.colors.black;
      
      // Set up period headers (row 5)
      debtSheet.getRange('A5').values = [['Period']];
      debtSheet.getRange('B5').values = [['0']];
      
      for (let i = 1; i <= periods; i++) {
        const colLetter = this.getColumnLetter(i + 1); // C, D, E, etc. for Periods 1, 2, 3...
        debtSheet.getRange(colLetter + '5').values = [[i]];
      }
      
      // Apply font formatting to period headers
      const periodRange = debtSheet.getRange(`A5:${this.getColumnLetter(totalColumns)}5`);
      periodRange.format.font.name = 'Times New Roman';
      periodRange.format.font.size = 12;
      periodRange.format.font.bold = false;
      
      // Calculate debt amount
      const dealValue = parseFloat(modelData.dealValue) || 0;
      const ltvRatio = parseFloat(modelData.dealLTV) / 100 || 0.7;
      const debtAmount = dealValue * ltvRatio;
      
      console.log('📊 Debt calculations:', { dealValue, ltvRatio, debtAmount });
      
      // Set up debt balance row (row 6)
      debtSheet.getRange('A6').values = [['Outstanding Debt Balance']];
      debtSheet.getRange('A6').format.font.name = 'Times New Roman';
      debtSheet.getRange('A6').format.font.size = 12;
      
      // Period 0: Show dash for debt balance (no debt issued yet)
      const dashRange = debtSheet.getRange('B6');
      dashRange.values = [['-']];
      dashRange.format.horizontalAlignment = 'Right';
      ExcelFormatter.applyNumberFormat(dashRange);
      
      // Debt balance for all periods (constant for interest-only loan)
      for (let i = 1; i <= periods; i++) {
        const colLetter = this.getColumnLetter(i + 1); // C, D, E, etc.
        if (i === periods) {
          // Final period: debt is repaid
          debtSheet.getRange(colLetter + '6').values = [[0]];
          ExcelFormatter.applyNumberFormat(debtSheet.getRange(colLetter + '6'));
        } else {
          // Operating periods: debt balance remains the same
          debtSheet.getRange(colLetter + '6').values = [[debtAmount]];
          ExcelFormatter.applyNumberFormat(debtSheet.getRange(colLetter + '6'));
        }
      }
      
      // Set up interest rate row (row 7) - Fixed rate only
      debtSheet.getRange('A7').values = [['Interest Rate (%)']];
      debtSheet.getRange('A7').format.font.name = 'Times New Roman';
      debtSheet.getRange('A7').format.font.size = 12;
      
      // Period 0: Show dash for interest rate
      const interestDashRange = debtSheet.getRange('B7');
      interestDashRange.values = [['-']];
      interestDashRange.format.horizontalAlignment = 'Right';
      
      console.log('📊 Using fixed rate:', modelData.debtSettings?.fixedRate || 5.5);
      
      const fixedRate = (modelData.debtSettings?.fixedRate || 5.5) / 100;
      
      for (let i = 1; i <= periods; i++) {
        const colLetter = this.getColumnLetter(i + 1); // C, D, E, etc.
        debtSheet.getRange(colLetter + '7').values = [[fixedRate]];
        debtSheet.getRange(colLetter + '7').numberFormat = [['0.0%']];
      }
      
      // Set up debt expense row (row 8) - This is what goes to FCF
      debtSheet.getRange('A8').values = [['Debt Expense per Period']];
      const debtExpenseRange = debtSheet.getRange(`A8:${this.getColumnLetter(totalColumns)}8`);
      debtExpenseRange.format.font.name = 'Times New Roman';
      debtExpenseRange.format.font.size = 12;
      debtExpenseRange.format.font.bold = true;
      debtExpenseRange.format.fill.color = ExcelFormatter.colors.backgroundDarker5;
      
      // Add thin underline at top of debt expense row
      debtExpenseRange.format.borders.getItem('EdgeTop').style = 'Continuous';
      debtExpenseRange.format.borders.getItem('EdgeTop').weight = 'Thin';
      debtExpenseRange.format.borders.getItem('EdgeTop').color = ExcelFormatter.colors.black;
      
      // Period 0: Show dash for debt expense
      const expenseDashRange = debtSheet.getRange('B8');
      expenseDashRange.values = [['-']];
      expenseDashRange.format.horizontalAlignment = 'Right';
      ExcelFormatter.applyNumberFormat(expenseDashRange);
      
      // Calculate debt expense for each period
      for (let i = 1; i <= periods; i++) {
        const colLetter = this.getColumnLetter(i + 1); // Current period column (C, D, E...)
        const currentBalanceCol = colLetter; // Use current period's balance for calculation
        
        if (i === periods) {
          // Final period: interest + principal repayment
          // Use previous period's balance (Period N-1) for both interest and principal
          const prevBalanceCol = this.getColumnLetter(i); // Previous period's balance (B for Period 0, C for Period 1, etc.)
          debtSheet.getRange(colLetter + '8').formulas = [[`=${prevBalanceCol}6*${colLetter}7+${prevBalanceCol}6`]];
          ExcelFormatter.applyNumberFormat(debtSheet.getRange(colLetter + '8'));
        } else {
          // Interest-only periods: just interest
          // Interest = Current Balance * Rate
          debtSheet.getRange(colLetter + '8').formulas = [[`=${currentBalanceCol}6*${currentBalanceCol}7`]];
          ExcelFormatter.applyNumberFormat(debtSheet.getRange(colLetter + '8'));
        }
      }
      
      await context.sync();
      
      // Re-apply percentage formatting to interest rate row
      for (let i = 1; i <= periods; i++) {
        const colLetter = this.getColumnLetter(i + 1); // C, D, E, etc.
        debtSheet.getRange(colLetter + '7').numberFormat = [['0.0%']];
      }
      
      // Auto-resize columns
      debtSheet.getUsedRange().format.autofitColumns();
      
      // Ensure all cells use Times New Roman font
      const allDebtCellsRange = debtSheet.getUsedRange();
      allDebtCellsRange.format.font.name = 'Times New Roman';
      allDebtCellsRange.format.font.size = 12;
      
      await context.sync();
      
      // Store for FCF reference
      this.debtModelInterestRow = 8; // Updated to point to debt expense row
      
      console.log('✅ Fixed rate debt model sheet created successfully!');
      return { success: true, message: 'Fixed rate debt model sheet created successfully' };
    });
  }

  // REMOVED: Old complex AI sheet creation - replaced with simple version above
}

// Export for use in main application
window.ExcelGenerator = ExcelGenerator;
window.CellTracker = CellTracker;

// Debug: Confirm export successful
console.log('🔧 ExcelGenerator.js: Classes exported to window', {
  ExcelGenerator: typeof window.ExcelGenerator,
  CellTracker: typeof window.CellTracker
});