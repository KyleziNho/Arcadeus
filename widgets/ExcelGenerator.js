/* global Office, Excel */

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
    
    // Equity Contribution (Calculated)
    sheet.getRange(`A${currentRow}`).values = [['Equity Contribution (Calculated)']];
    const dealValueCell = this.cellTracker.getCellReference('dealValue').split('!')[1];
    const ltvCell = this.cellTracker.getCellReference('dealLTV').split('!')[1];
    sheet.getRange(`B${currentRow}`).formulas = [[`=${dealValueCell}*(1-${ltvCell}/100)`]];
    sheet.getRange(`B${currentRow}`).format.font.italic = true;
    this.cellTracker.recordCell('equityContribution', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    // Debt Financing (Calculated)
    sheet.getRange(`A${currentRow}`).values = [['Debt Financing (Calculated)']];
    sheet.getRange(`B${currentRow}`).formulas = [[`=${dealValueCell}*${ltvCell}/100`]];
    sheet.getRange(`B${currentRow}`).format.font.italic = true;
    this.cellTracker.recordCell('debtFinancing', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    currentRow += 2; // Add space
    
    // REVENUE ITEMS SECTION
    if (data.revenueItems && data.revenueItems.length > 0) {
      console.log('📝 ====== WRITING REVENUE ITEMS TO EXCEL ======');
      console.log('📝 Number of revenue items:', data.revenueItems.length);
      
      sectionRows['revenueItems'] = currentRow;
      sheet.getRange(`A${currentRow}`).values = [['REVENUE ITEMS']];
      sheet.getRange(`A${currentRow}`).format.font.bold = true;
      currentRow += 1;
      
      // Add column headers
      sheet.getRange(`A${currentRow}`).values = [['Name']];
      sheet.getRange(`B${currentRow}`).values = [['Base Value']];
      sheet.getRange(`C${currentRow}`).values = [['Growth Rate']];
      sheet.getRange(`A${currentRow}:C${currentRow}`).format.font.bold = true;
      currentRow += 1;
      
      const revenueStartRow = currentRow;
      data.revenueItems.forEach((item, index) => {
        console.log(`📝 Writing revenue item ${index + 1}:`, item);
        const itemName = item.name || `Revenue Item ${index + 1}`;
        const growthRate = item.growthRate || 0;
        sheet.getRange(`A${currentRow}`).values = [[itemName]];
        sheet.getRange(`B${currentRow}`).values = [[item.value || 0]];
        sheet.getRange(`C${currentRow}`).values = [[`${growthRate}%`]];
        this.cellTracker.recordCell(`revenue_${index}`, 'Assumptions', `B${currentRow}`);
        this.cellTracker.recordCell(`revenue_${index}_name`, 'Assumptions', `A${currentRow}`);
        this.cellTracker.recordCell(`revenue_${index}_growth_rate`, 'Assumptions', `C${currentRow}`);
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
      currentRow += 1;
      
      // Add column headers
      sheet.getRange(`A${currentRow}`).values = [['Name']];
      sheet.getRange(`B${currentRow}`).values = [['Base Value']];
      sheet.getRange(`C${currentRow}`).values = [['Growth Rate']];
      sheet.getRange(`A${currentRow}:C${currentRow}`).format.font.bold = true;
      currentRow += 1;
      
      const opexStartRow = currentRow;
      data.operatingExpenses.forEach((item, index) => {
        const itemName = item.name || `OpEx Item ${index + 1}`;
        const growthRate = item.growthRate || 0;
        sheet.getRange(`A${currentRow}`).values = [[itemName]];
        sheet.getRange(`B${currentRow}`).values = [[item.value || 0]];
        sheet.getRange(`C${currentRow}`).values = [[`${growthRate}%`]];
        this.cellTracker.recordCell(`opex_${index}`, 'Assumptions', `B${currentRow}`);
        this.cellTracker.recordCell(`opex_${index}_name`, 'Assumptions', `A${currentRow}`);
        this.cellTracker.recordCell(`opex_${index}_growth_rate`, 'Assumptions', `C${currentRow}`);
        currentRow++;
      });
      
      // Record the range of operating expenses for future reference
      this.cellTracker.recordCell('opex_range', 'Assumptions', `B${opexStartRow}:B${currentRow - 1}`);
      this.cellTracker.recordCell('opex_count', 'Assumptions', data.operatingExpenses.length.toString());
      
      currentRow += 2; // Add space
    }
    
    // CAPEX SECTION
    if (data.capEx && data.capEx.length > 0) {
      sectionRows['capEx'] = currentRow;
      sheet.getRange(`A${currentRow}`).values = [['CAPITAL EXPENDITURES (CAPEX)']];
      sheet.getRange(`A${currentRow}`).format.font.bold = true;
      currentRow += 1;
      
      // Add column headers
      sheet.getRange(`A${currentRow}`).values = [['CapEx Name']];
      sheet.getRange(`B${currentRow}`).values = [['Annual Value']];
      sheet.getRange(`C${currentRow}`).values = [['Growth Rate (%)']];
      sheet.getRange(`A${currentRow}:C${currentRow}`).format.font.bold = true;
      currentRow += 1;
      
      const capexStartRow = currentRow;
      data.capEx.forEach((item, index) => {
        const itemName = item.name || `CapEx ${index + 1}`;
        const growthRate = item.growthRate || 0;
        sheet.getRange(`A${currentRow}`).values = [[itemName]];
        sheet.getRange(`B${currentRow}`).values = [[item.value || 0]];
        sheet.getRange(`C${currentRow}`).values = [[`${growthRate}%`]];
        this.cellTracker.recordCell(`capex_${index}`, 'Assumptions', `B${currentRow}`);
        this.cellTracker.recordCell(`capex_${index}_name`, 'Assumptions', `A${currentRow}`);
        this.cellTracker.recordCell(`capex_${index}_growth_rate`, 'Assumptions', `C${currentRow}`);
        currentRow++;
      });
      
      // Record the range of CapEx for future reference
      this.cellTracker.recordCell('capex_range', 'Assumptions', `B${capexStartRow}:B${currentRow - 1}`);
      this.cellTracker.recordCell('capex_count', 'Assumptions', data.capEx.length.toString());
      
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
    
    // Discount Rate (WACC)
    sheet.getRange(`A${currentRow}`).values = [['Discount Rate - WACC (%)']];
    sheet.getRange(`B${currentRow}`).values = [[data.discountRate || 10.0]];
    this.cellTracker.recordCell('discountRate', 'Assumptions', `B${currentRow}`);
    currentRow++;
    
    currentRow += 2; // Add space
    
    // DEBT MODEL SECTION
    sectionRows['debtModel'] = currentRow;
    sheet.getRange(`A${currentRow}`).values = [['DEBT MODEL']];
    sheet.getRange(`A${currentRow}`).format.font.bold = true;
    currentRow += 2;
    
    // Check if debt financing is enabled (LTV > 0)
    const hasDebt = data.dealLTV && parseFloat(data.dealLTV) > 0;
    
    if (hasDebt) {
      // Loan Issuance Fees
      sheet.getRange(`A${currentRow}`).values = [['Loan Issuance Fees (%)']];
      sheet.getRange(`B${currentRow}`).values = [[data.loanIssuanceFees || 1.5]];
      this.cellTracker.recordCell('loanIssuanceFees', 'Assumptions', `B${currentRow}`);
      currentRow++;
      
      // Interest Rate Type
      sheet.getRange(`A${currentRow}`).values = [['Interest Rate Type']];
      sheet.getRange(`B${currentRow}`).values = [[data.interestRateType || 'fixed']];
      this.cellTracker.recordCell('interestRateType', 'Assumptions', `B${currentRow}`);
      currentRow++;
      
      // Interest Rate Details
      if (data.interestRateType === 'floating') {
        // Base Rate
        sheet.getRange(`A${currentRow}`).values = [['Base Interest Rate (%)']];
        sheet.getRange(`B${currentRow}`).values = [[data.baseRate || 3.9]];
        this.cellTracker.recordCell('baseRate', 'Assumptions', `B${currentRow}`);
        currentRow++;
        
        // Credit Margin
        sheet.getRange(`A${currentRow}`).values = [['Credit Margin (%)']];
        sheet.getRange(`B${currentRow}`).values = [[data.creditMargin || 2.0]];
        this.cellTracker.recordCell('creditMargin', 'Assumptions', `B${currentRow}`);
        currentRow++;
        
        // Total Floating Rate
        sheet.getRange(`A${currentRow}`).values = [['Total Interest Rate (%)']];
        sheet.getRange(`B${currentRow}`).formulas = [[`=B${currentRow-2}+B${currentRow-1}`]];
        this.cellTracker.recordCell('totalInterestRate', 'Assumptions', `B${currentRow}`);
        currentRow++;
      } else {
        // Fixed Rate
        sheet.getRange(`A${currentRow}`).values = [['Fixed Interest Rate (%)']];
        sheet.getRange(`B${currentRow}`).values = [[data.fixedRate || 5.5]];
        this.cellTracker.recordCell('fixedRate', 'Assumptions', `B${currentRow}`);
        currentRow++;
      }
    } else {
      sheet.getRange(`A${currentRow}`).values = [['No Debt Financing (LTV = 0)']];
      sheet.getRange(`A${currentRow}`).format.font.italic = true;
      currentRow++;
    }
    
    // Auto-resize columns
    sheet.getRange('A:B').format.autofitColumns();
    
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

6. **Interest Expense (if debt exists):**
   - Calculate based on debt financing amount and interest rate
   - Adjust for period type:
     * Daily: Debt * Interest_Rate / 365 / 100
     * Monthly: Debt * Interest_Rate / 12 / 100
     * Quarterly: Debt * Interest_Rate / 4 / 100
     * Yearly: Debt * Interest_Rate / 100

7. **Net Income:**
   - Net Income = NOI - CapEx - Interest Expense

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
    output += `- Interest Rate Type: ${modelData.interestRateType || 'fixed'}\n`;
    
    if (modelData.interestRateType === 'floating') {
      output += `- Base Rate: ${this.cellTracker.getCellReference('baseRate')}\n`;
      output += `- Credit Margin: ${this.cellTracker.getCellReference('creditMargin')}\n`;
      output += `- Total Interest Rate: ${this.cellTracker.getCellReference('totalInterestRate')}\n`;
    } else {
      output += `- Fixed Interest Rate: ${this.cellTracker.getCellReference('fixedRate')}\n`;
    }
    
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
      
      // Step 3: Generate comprehensive FCF AI prompt with ACTUAL cell references
      const fcfPrompt = this.generateRealFCFPrompt(modelData, plStructure, assumptionStructure);
      
      // Step 4: Create professional FCF sheet using discovered cell references
      await this.createAIFCFSheet(modelData, fcfPrompt, plStructure, assumptionStructure);
      
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
      
      let currentRow = 1;
      
      // TITLE
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
      const revenueStartRow = currentRow;
      plSheet.getRange(`A${currentRow}`).values = [['REVENUE']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#87CEEB';
      currentRow++;
      
      // Add each revenue item
      if (modelData.revenueItems && modelData.revenueItems.length > 0) {
        modelData.revenueItems.forEach((item, index) => {
          plSheet.getRange(`A${currentRow}`).values = [[item.name || `Revenue ${index + 1}`]];
          
          // Add formulas for each period
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            
            if (col === 1) {
              // First period - reference from Assumptions
              const assumptionRef = this.cellTracker.getCellReference(`revenue_${index}`);
              if (assumptionRef) {
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${assumptionRef}`]];
              } else {
                plSheet.getRange(`${colLetter}${currentRow}`).values = [[item.value || 0]];
              }
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
                  [[`=${prevCol}${currentRow}*(1+Assumptions!${cellRef}${periodAdjustment}/100)`]];
              } else {
                // No growth rate in assumptions - use flat growth
                console.log(`No growth rate reference found for revenue ${index}`);
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevCol}${currentRow}`]];
              }
            }
          }
          currentRow++;
        });
      }
      
      // Total Revenue
      plSheet.getRange(`A${currentRow}`).values = [['Total Revenue']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      
      for (let col = 1; col <= periodColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (modelData.revenueItems && modelData.revenueItems.length > 0) {
          const sumFormula = `=SUM(${colLetter}${revenueStartRow + 1}:${colLetter}${currentRow - 1})`;
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[sumFormula]];
        } else {
          plSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
        }
      }
      const totalRevenueRow = currentRow;
      currentRow += 2;
      
      // OPERATING EXPENSES SECTION
      const opexStartRow = currentRow;
      plSheet.getRange(`A${currentRow}`).values = [['OPERATING EXPENSES']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#FFB6C1';
      currentRow++;
      
      // Add each operating expense
      if (modelData.operatingExpenses && modelData.operatingExpenses.length > 0) {
        modelData.operatingExpenses.forEach((item, index) => {
          plSheet.getRange(`A${currentRow}`).values = [[item.name || `OpEx ${index + 1}`]];
          
          // Add formulas for each period
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            
            if (col === 1) {
              // First period - negative value from Assumptions
              const assumptionRef = this.cellTracker.getCellReference(`opex_${index}`);
              if (assumptionRef) {
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=-${assumptionRef}`]];
              } else {
                plSheet.getRange(`${colLetter}${currentRow}`).values = [[-item.value || 0]];
              }
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
                  [[`=${prevCol}${currentRow}*(1+Assumptions!${cellRef}${periodAdjustment}/100)`]];
              } else {
                // No growth rate in assumptions - use flat growth
                console.log(`No growth rate reference found for opex ${index}`);
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevCol}${currentRow}`]];
              }
            }
          }
          currentRow++;
        });
      }
      
      // Total Operating Expenses
      plSheet.getRange(`A${currentRow}`).values = [['Total Operating Expenses']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      
      for (let col = 1; col <= periodColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (modelData.operatingExpenses && modelData.operatingExpenses.length > 0) {
          const sumFormula = `=SUM(${colLetter}${opexStartRow + 1}:${colLetter}${currentRow - 1})`;
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[sumFormula]];
        } else {
          plSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
        }
      }
      const totalOpexRow = currentRow;
      currentRow += 2;
      
      // NOI
      plSheet.getRange(`A${currentRow}`).values = [['NOI']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#98FB98';
      
      for (let col = 1; col <= periodColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
          [[`=${colLetter}${totalRevenueRow}+${colLetter}${totalOpexRow}`]];
      }
      const ebitdaRow = currentRow;
      currentRow += 2;
      
      // INTEREST EXPENSE (if debt exists)
      const hasDebt = modelData.dealLTV && parseFloat(modelData.dealLTV) > 0;
      let interestRow = 0;
      
      if (hasDebt) {
        plSheet.getRange(`A${currentRow}`).values = [['Interest Expense']];
        plSheet.getRange(`A${currentRow}`).format.font.bold = true;
        
        const debtRef = this.cellTracker.getCellReference('debtFinancing');
        let rateRef;
        
        if (modelData.interestRateType === 'floating') {
          rateRef = this.cellTracker.getCellReference('totalInterestRate');
        } else {
          rateRef = this.cellTracker.getCellReference('fixedRate');
        }
        
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          const periodAdjustment = this.getPeriodAdjustment(modelData.modelPeriods);
          
          if (debtRef && rateRef) {
            plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
              [[`=-${debtRef}*${rateRef}${periodAdjustment}/100`]];
          } else {
            // Fallback calculation
            const debtAmount = modelData.dealValue * (modelData.dealLTV / 100);
            const rate = modelData.fixedRate || 5.5;
            const periodRate = this.calculatePeriodRate(rate, modelData.modelPeriods);
            plSheet.getRange(`${colLetter}${currentRow}`).values = [[-debtAmount * periodRate / 100]];
          }
        }
        interestRow = currentRow;
        currentRow += 2;
      }
      
      // NET INCOME
      plSheet.getRange(`A${currentRow}`).values = [['Net Income']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#FFD700';
      
      for (let col = 1; col <= periodColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (hasDebt && interestRow > 0) {
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
            [[`=${colLetter}${ebitdaRow}+${colLetter}${interestRow}`]];
        } else {
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
            [[`=${colLetter}${ebitdaRow}`]];
        }
      }
      
      // Track Net Income row
      this.plCellTracker.recordCell('net_income', 'P&L Statement', `B${currentRow}:${this.getColumnLetter(periodColumns)}${currentRow}`);
      
      // Format numbers with red brackets for negatives and dash for empty cells
      const dataRange = plSheet.getRange(`B5:${this.getColumnLetter(periodColumns)}${currentRow}`);
      const numberFormat = '#,##0_);[Red](#,##0);"-"'; // Positive, negative (red brackets), zero as dash
      dataRange.numberFormat = [[numberFormat]];
      
      // Auto-fit columns
      plSheet.getRange(`A:${this.getColumnLetter(periodColumns)}`).format.autofitColumns();
      
      await context.sync();
      console.log('✅ P&L Statement created successfully');
      
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
  generateRealFCFPrompt(modelData, plStructure, assumptionStructure) {
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
- Start with Net Income from P&L: ${this.plCellTracker.getCellReference('net_income')}
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
      output += `- Interest Rate Type: ${this.cellTracker.getCellReference('interestRateType')}\n`;
      if (modelData.interestRateType === 'floating') {
        output += `- Base Rate (%): ${this.cellTracker.getCellReference('baseRate')}\n`;
        output += `- Credit Margin (%): ${this.cellTracker.getCellReference('creditMargin')}\n`;
        output += `- Total Interest Rate (%): ${this.cellTracker.getCellReference('totalInterestRate')}\n`;
      } else {
        output += `- Fixed Interest Rate (%): ${this.cellTracker.getCellReference('fixedRate')}\n`;
      }
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
  async createAIFCFSheet(modelData, aiPrompt, plStructure, assumptionStructure) {
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
      headers.push('Initial Investment'); // Period 0
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
            const plCol = this.getColumnLetter(col === 1 ? 2 : col + 1); // Period 0 references P&L column B, others offset by 1
            fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`='P&L Statement'!${plCol}${noiRow}`]];
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
        
        // Interest Payments using REAL Interest Expense reference
        fcfSheet.getRange(`A${currentRow}`).values = [['Less: Interest Payments']];
        fcfStructure.interestPayments = currentRow;
        if (plStructure.lineItems.interestExpense) {
          const interestRow = plStructure.lineItems.interestExpense.row;
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            const plCol = this.getColumnLetter(col === 1 ? 2 : col + 1); // Period 0 references P&L column B, others offset by 1
            fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`='P&L Statement'!${plCol}${interestRow}`]];
          }
        } else {
          // No interest expense found
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
              fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${dealValueRef}*(1-${disposalCostRef}/100)`]];
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
          fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${fcfStructure.unleveredFCF}+${colLetter}${fcfStructure.interestPayments}+${colLetter}${fcfStructure.assetDisposal}`]];
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
          fcfSheet.getRange(`B${currentRow}`).formulas = [[`=NPV(Assumptions!${waccCellRef}/100,${undiscountedRange})`]];
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
      const numberFormat = '#,##0_);[Red](#,##0);"-"'; // Positive, negative (red brackets), zero as dash
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
    
    let debtInfo = `- Loan Issuance Fees: ${modelData.loanIssuanceFees || 1.5}% (Cell: ${this.cellTracker.getCellReference('loanIssuanceFees')})\n`;
    debtInfo += `- Interest Rate Type: ${modelData.interestRateType || 'fixed'} (Cell: ${this.cellTracker.getCellReference('interestRateType')})\n`;
    
    if (modelData.interestRateType === 'floating') {
      debtInfo += `- Base Rate: ${modelData.baseRate || 3.9}% (Cell: ${this.cellTracker.getCellReference('baseRate')})\n`;
      debtInfo += `- Credit Margin: ${modelData.creditMargin || 2.0}% (Cell: ${this.cellTracker.getCellReference('creditMargin')})\n`;
      debtInfo += `- Total Rate: Formula in (Cell: ${this.cellTracker.getCellReference('totalInterestRate')})`;
    } else {
      debtInfo += `- Fixed Rate: ${modelData.fixedRate || 5.5}% (Cell: ${this.cellTracker.getCellReference('fixedRate')})`;
    }
    
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
                  growthFormula = `=${prevCol}${currentRow}*(1+${periodRate}/100)`;
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
                  growthFormula = `=${prevCol}${currentRow}*(1+${periodRate}/100)`;
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
      
      // DEBT SERVICE SECTION (if applicable)
      const hasDebt = modelData.dealLTV && parseFloat(modelData.dealLTV) > 0;
      if (hasDebt) {
        plSheet.getRange(`A${currentRow}`).values = [['Interest Expense']];
        plSheet.getRange(`A${currentRow}`).format.font.bold = true;
        
        // Get debt financing and interest rate references
        const debtRef = this.cellTracker.getCellReference('debtFinancing');
        let interestRateRef;
        
        if (modelData.interestRateType === 'floating') {
          interestRateRef = this.cellTracker.getCellReference('totalInterestRate');
        } else {
          interestRateRef = this.cellTracker.getCellReference('fixedRate');
        }
        
        // Interest expense for each period
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          let interestFormula = '';
          
          // Adjust interest rate based on period type
          // Fix debt and interest rate references
          const debtCellRef = debtRef ? `Assumptions!${debtRef.split('!')[1]}` : 'Assumptions!B8';
          const rateCellRef = interestRateRef ? `Assumptions!${interestRateRef.split('!')[1]}` : 'Assumptions!B15';
          
          switch (modelData.modelPeriods) {
            case 'daily':
              interestFormula = `=-${debtCellRef}*${rateCellRef}/100/365`;
              break;
            case 'monthly':
              interestFormula = `=-${debtCellRef}*${rateCellRef}/12`;
              break;
            case 'quarterly':
              interestFormula = `=-${debtCellRef}*${rateCellRef}/4`;
              break;
            case 'yearly':
            default:
              interestFormula = `=-${debtCellRef}*${rateCellRef}/100`;
          }
          
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[interestFormula]];
        }
        currentRow += 2;
      }
      
      // Store interest expense row for Net Income calculation
      const interestExpenseRow = hasDebt ? currentRow - 2 : 0;
      
      // NET INCOME
      plSheet.getRange(`A${currentRow}`).values = [['Net Income']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#FFD700';
      
      // Find NOI row (it's 2 rows up from current, or 4 if we have debt)
      const ebitdaRowForNetIncome = hasDebt ? currentRow - 4 : currentRow - 2;
      
      // Net Income formulas for each period
      for (let col = 1; col <= periodColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (hasDebt && interestExpenseRow > 0) {
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
            [[`=${colLetter}${ebitdaRowForNetIncome}+${colLetter}${interestExpenseRow}`]];
        } else {
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
            [[`=${colLetter}${ebitdaRowForNetIncome}`]];
        }
      }
      
      // Format numbers for all data columns
      const dataRange = plSheet.getRange(`B5:${this.getColumnLetter(periodColumns)}${currentRow}`);
      dataRange.numberFormat = [['#,##0_);[Red](#,##0);"-"']];
      
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
        return `Q${(periodIndex % 4) + 1} ${date.getFullYear()}`;
      case 'yearly':
        date.setFullYear(date.getFullYear() + periodIndex);
        return date.getFullYear().toString();
      default:
        date.setMonth(date.getMonth() + periodIndex);
        return date.toLocaleDateString('en-US', { year: 'numeric', month: 'short' });
    }
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
      return `=${prevCellRef}*(1+${growthRateRef}${periodAdjustment}/100)`;
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
  generateRealFCFPrompt(modelData, plStructure, assumptionStructure) {
    console.log('🤖 Generating REAL FCF AI prompt with discovered cell references...');
    
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
${this.formatAssumptionStructureForPrompt(assumptionStructure)}

**REQUIRED FCF SHEET STRUCTURE:**

Create a comprehensive Free Cash Flow statement with the following structure:
- Column A: Line item names  
- Columns B through ${this.getColumnLetter(maxPeriods)}: Period data (${maxPeriods} periods total)
- Period Headers: ${periodHeaders.join(', ')}

**FCF CALCULATION METHODOLOGY:**

1. **OPERATING CASH FLOW SECTION:**
   - NOI: Reference the exact NOI row from P&L
   - Real estate model: No tax calculations required

2. **WORKING CAPITAL ADJUSTMENTS:**
   - Change in Working Capital (2% of Total Revenue change)
   - Calculate as: Current Period Revenue * 2% - Previous Period Revenue * 2%

3. **INITIAL INVESTMENTS:**
   - Reference capital expense items from Assumptions if any exist
   - Apply growth rates if specified

4. **UNLEVERED FREE CASH FLOW:**
   - = NOPAT - Change in Working Capital - Initial Investments

5. **FINANCING CASH FLOWS:**
   - Interest Payments: Reference from P&L Interest Expense line
   - Principal Repayments (if applicable)

6. **LEVERED FREE CASH FLOW:**
   - = Unlevered FCF - Interest Payments - Principal Repayments

7. **CUMULATIVE METRICS:**
   - Cumulative FCF
   - Undiscounted NPV (simple sum of FCF)
   - Discounted NPV using WACC from Assumptions
   - IRR calculation including initial investment

**CRITICAL REQUIREMENTS:**

1. **Use EXACT cell references** from the discovered P&L and Assumptions structures above
2. **Reference format**: Use 'P&L Statement'!B15 or 'Assumptions'!B23 format  
3. **Handle missing data**: If a P&L line item is not found, use conservative estimates
4. **Most Important**: ALWAYS reference the Net Income line from P&L as the starting point for FCF calculations
5. **Period consistency**: Ensure all ${maxPeriods} periods are calculated
6. **Formula accuracy**: All formulas must be valid Excel syntax

**NET INCOME PRIORITY:**
The most critical value to extract from the P&L is the Net Income for each period, located at: ${plStructure.lineItems?.netIncome?.cellRef || 'Not Found - Please locate manually'}

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
    
    if (plStructure.lineItems.interestExpense) {
      output += `- Interest Expense: Row ${plStructure.lineItems.interestExpense.row}, Range B${plStructure.lineItems.interestExpense.row}:${this.getColumnLetter(maxPeriods)}${plStructure.lineItems.interestExpense.row}\n`;
    }
    
    if (plStructure.lineItems.netIncome) {
      output += `- 🎯 NET INCOME (CRITICAL): Row ${plStructure.lineItems.netIncome.row}, Range B${plStructure.lineItems.netIncome.row}:${this.getColumnLetter(maxPeriods)}${plStructure.lineItems.netIncome.row}\n`;
    }
    
    if (plStructure.lineItems.totalCapEx) {
      output += `- Total CapEx: Row ${plStructure.lineItems.totalCapEx.row}, Range B${plStructure.lineItems.totalCapEx.row}:${this.getColumnLetter(maxPeriods)}${plStructure.lineItems.totalCapEx.row}\n`;
    }

    output += `\n**REFERENCE FORMAT:** Use 'P&L Statement'!B[row]:[column][row] for ranges\n`;
    output += `**EXAMPLE:** 'P&L Statement'!B${plStructure.lineItems.netIncome?.row || '15'}:${this.getColumnLetter(maxPeriods)}${plStructure.lineItems.netIncome?.row || '15'} for Net Income across all periods\n\n`;

    return output;
  }

  // Format Assumptions structure for AI prompt  
  formatAssumptionStructureForPrompt(assumptionStructure) {
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
      sheet.getRange(`B${row}`).format.numberFormat = [['#,##0_);[Red](#,##0)']];
      row++;
      
      sheet.getRange(`A${row}`).values = [['Equity Investment:']];
      sheet.getRange(`B${row}`).values = [[equityContribution]];
      sheet.getRange(`B${row}`).format.numberFormat = [['#,##0_);[Red](#,##0)']];
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

    prompt += `\n\n**REQUIRED P&L STRUCTURE WITH PERIOD 0:**
You MUST create a P&L Statement with this EXACT structure:

**CRITICAL: Include Period 0 (Initial Investment) before operating periods**
- Period 0: "Initial Investment" 
- Period 1: First operating period (${modelData.projectStartDate})
- Period 2+: Subsequent operating periods

**P&L STRUCTURE:**
1. REVENUE section (Period 0 = 0, then operating periods with growth)
2. Total Revenue  
3. OPERATING EXPENSES section (Period 0 = 0, then operating periods with growth)
4. Total Operating Expenses  
5. NOI (Total Revenue - Total Operating Expenses)
6. Interest Expense (if debt exists)
7. Net Income (NOI - Interest Expense)

**CRITICAL REQUIREMENTS:**
- Real estate model: No depreciation or tax calculations required
- ALL formulas must reference exact cells provided above
- Return complete P&L structure in JSON format with exact formulas

**JSON FORMAT REQUIRED:**
{
  "plStructure": {
    "revenueItems": [{"name": "...", "formulas": ["=cell1", "=cell2", ...]}],
    "totalRevenue": {"formula": "=SUM(...)"},
    "operatingExpenses": [{"name": "...", "formulas": ["=cell1", "=cell2", ...]}],
    "totalOpEx": {"formula": "=SUM(...)"},
    "noi": {"formula": "=TotalRevenue-TotalOpEx"},
    "interestExpense": {"formula": "=..."},
    "netIncome": {"formula": "=NOI-InterestExpense"}
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
        await this.createPLSheetFromAI(modelData, plStructure.plStructure);
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

      let currentRow = 1;

      // TITLE
      plSheet.getRange('A1').values = [['Profit & Loss Statement']];
      plSheet.getRange('A1').format.font.bold = true;
      plSheet.getRange('A1').format.font.size = 16;
      plSheet.getRange('A1').format.fill.color = '#1f4e79';
      plSheet.getRange('A1').format.font.color = 'white';
      currentRow = 3;

      // TIME PERIOD HEADERS - Include Period 0 for Initial Investment
      const headers = [''];
      headers.push('Initial Investment'); // Period 0
      const startDate = new Date(modelData.projectStartDate);
      for (let i = 0; i < periodColumns; i++) {
        headers.push(this.formatPeriodHeader(startDate, i, modelData.modelPeriods));
      }
      const totalColumns = periodColumns + 1; // +1 for Initial Investment period

      const headerRange = plSheet.getRange(`A${currentRow}:${this.getColumnLetter(totalColumns)}${currentRow}`);
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = '#d9d9d9';
      currentRow++;

      // REVENUE SECTION
      plSheet.getRange(`A${currentRow}`).values = [['REVENUE']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#e8f5e8';
      currentRow++;

      // Add revenue items with growth
      if (modelData.revenueItems && modelData.revenueItems.length > 0) {
        modelData.revenueItems.forEach((item, index) => {
          plSheet.getRange(`A${currentRow}`).values = [[item.name]];
          
          const valueRef = this.cellTracker.getCellReference(`revenue_${index}`);
          const growthRateRef = this.cellTracker.getCellReference(`revenue_${index}_growth_rate`);
          
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            if (col === 1) {
              // Period 0 (Initial Investment): No revenue
              plSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
            } else if (col === 2) {
              // First operating period: base value
              plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${valueRef}`]];
            } else {
              // Subsequent periods: apply growth
              if (growthRateRef) {
                const prevColLetter = this.getColumnLetter(col - 1);
                if (modelData.modelPeriods === 'monthly') {
                  plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef}/12)`]];
                } else if (modelData.modelPeriods === 'quarterly') {
                  plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef}/4)`]];
                } else {
                  plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef}/100)`]];
                }
              } else {
                const prevColLetter = this.getColumnLetter(col - 1);
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}`]];
              }
            }
          }
          currentRow++;
        });
      }

      // Total Revenue
      plSheet.getRange(`A${currentRow}`).values = [['Total Revenue']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      const totalRevenueRow = currentRow;
      for (let col = 1; col <= totalColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        const revenueStartRow = totalRevenueRow - modelData.revenueItems.length;
        const revenueEndRow = totalRevenueRow - 1;
        plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=SUM(${colLetter}${revenueStartRow + 1}:${colLetter}${revenueEndRow})`]];
      }
      currentRow += 2;

      // OPERATING EXPENSES SECTION
      plSheet.getRange(`A${currentRow}`).values = [['OPERATING EXPENSES']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#ffe6e6';
      currentRow++;

      // Add operating expense items
      if (modelData.operatingExpenses && modelData.operatingExpenses.length > 0) {
        modelData.operatingExpenses.forEach((item, index) => {
          plSheet.getRange(`A${currentRow}`).values = [[item.name]];
          
          const valueRef = this.cellTracker.getCellReference(`opex_${index}`);
          const growthRateRef = this.cellTracker.getCellReference(`opex_${index}_growth_rate`);
          
          for (let col = 1; col <= totalColumns; col++) {
            const colLetter = this.getColumnLetter(col);
            if (col === 1) {
              // Period 0 (Initial Investment): No operating expenses
              plSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
            } else if (col === 2) {
              // First operating period: base value
              plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=-${valueRef}`]];
            } else {
              // Subsequent periods: apply growth
              if (growthRateRef) {
                const prevColLetter = this.getColumnLetter(col - 1);
                if (modelData.modelPeriods === 'monthly') {
                  plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef}/12)`]];
                } else if (modelData.modelPeriods === 'quarterly') {
                  plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef}/4)`]];
                } else {
                  plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}*(1+${growthRateRef}/100)`]];
                }
              } else {
                const prevColLetter = this.getColumnLetter(col - 1);
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevColLetter}${currentRow}`]];
              }
            }
          }
          currentRow++;
        });
      }

      // Total Operating Expenses
      plSheet.getRange(`A${currentRow}`).values = [['Total Operating Expenses']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      const totalOpExRow = currentRow;
      for (let col = 1; col <= totalColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        const opexStartRow = totalOpExRow - modelData.operatingExpenses.length;
        const opexEndRow = totalOpExRow - 1;
        plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=SUM(${colLetter}${opexStartRow + 1}:${colLetter}${opexEndRow})`]];
      }
      currentRow++;

      // NOI
      plSheet.getRange(`A${currentRow}`).values = [['NOI']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#fff2cc';
      const ebitdaRow = currentRow;
      for (let col = 1; col <= totalColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${totalRevenueRow}+${colLetter}${totalOpExRow}`]];
      }
      currentRow += 2;

      // CAPEX SECTION
      if (modelData.capEx && modelData.capEx.length > 0) {
        plSheet.getRange(`A${currentRow}`).values = [['CAPITAL EXPENDITURES']];
        plSheet.getRange(`A${currentRow}`).format.font.bold = true;
        plSheet.getRange(`A${currentRow}`).format.fill.color = '#f2f2f2';
        currentRow++;

        const capexStartRow = currentRow;
      
        // Add individual CapEx items
        modelData.capEx.forEach((item, index) => {
          plSheet.getRange(`A${currentRow}`).values = [[item.name || `CapEx ${index + 1}`]];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#ffeaa7';
      for (let col = 1; col <= totalColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (col === 1) {
          // Period 0: Show total initial capital investments
          if (modelData.capitalExpenses && modelData.capitalExpenses.length > 0) {
            let totalCapExFormula = '';
            modelData.capitalExpenses.forEach((item, index) => {
              const valueRef = this.cellTracker.getCellReference(`capex_${index}`);
              if (valueRef) {
                if (index === 0) {
                  totalCapExFormula = `-${valueRef}`;
                } else {
                  totalCapExFormula += `-${valueRef}`;
                }
              }
            });
            if (totalCapExFormula) {
              plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[totalCapExFormula]];
            } else {
              plSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
            }
          } else {
            plSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
          }
        } else {
          // Operating periods: No initial investments
          plSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
        }
      }
      currentRow++;
      
      // Real estate model: No depreciation calculations required
      // NOI is the final metric for real estate models

      // Interest Expense (if debt exists)
      let interestExpenseRow = null;
      if (modelData.dealLTV && parseFloat(modelData.dealLTV) > 0) {
        plSheet.getRange(`A${currentRow}`).values = [['Interest Expense']];
        interestExpenseRow = currentRow;
        // Simplified interest calculation for now
        for (let col = 1; col <= totalColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          plSheet.getRange(`${colLetter}${currentRow}`).values = [[0]]; // Placeholder
        }
        currentRow++;
      }

      // EBT
      plSheet.getRange(`A${currentRow}`).values = [['EBT (Earnings Before Tax)']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      const ebtRow = currentRow;
      for (let col = 1; col <= totalColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (interestExpenseRow) {
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${ebitRow}-${colLetter}${interestExpenseRow}`]];
        } else {
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${ebitRow}`]];
        }
      }
      currentRow++;

      // Tax Expense (25% default)
      plSheet.getRange(`A${currentRow}`).values = [['Tax Expense (25%)']];
      const taxExpenseRow = currentRow;
      for (let col = 1; col <= totalColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${ebtRow}*0.25`]];
      }
      currentRow++;

      // Net Income
      plSheet.getRange(`A${currentRow}`).values = [['Net Income']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#c5e0b4';
      plSheet.getRange(`A${currentRow}`).format.borders.getItem('EdgeTop').style = 'Double';
      plSheet.getRange(`A${currentRow}`).format.borders.getItem('EdgeTop').weight = 'Thick';
      for (let col = 1; col <= totalColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${ebtRow}-${colLetter}${taxExpenseRow}`]];
        plSheet.getRange(`${colLetter}${currentRow}`).format.borders.getItem('EdgeTop').style = 'Double';
        plSheet.getRange(`${colLetter}${currentRow}`).format.borders.getItem('EdgeTop').weight = 'Thick';
      }

      // Format all data cells with red brackets for negatives and dash for empty cells
      const dataRange = plSheet.getRange(`B5:${this.getColumnLetter(totalColumns)}${currentRow}`);
      const numberFormat = '#,##0_);[Red](#,##0);"-"'; // Positive, negative (red brackets), zero as dash
      dataRange.numberFormat = [[numberFormat]];
      
      // Auto-resize columns
      plSheet.getUsedRange().format.autofitColumns();
      
      console.log('✅ Enhanced P&L Statement with depreciation created successfully!');
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