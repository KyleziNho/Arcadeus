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
    this.plCellTracker = new CellTracker(); // Track P&L cell references
    this.currentWorkbook = null;
  }

  async generateModel(modelData) {
    try {
      console.log('🚀 Starting fresh model generation...');
      console.log('📊 Model data:', modelData);
      
      // Reset cell trackers
      this.cellTracker = new CellTracker();
      this.plCellTracker = new CellTracker();
      
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
      
      // Add growth rates for revenue items
      currentRow++;
      sheet.getRange(`A${currentRow}`).values = [['Revenue Growth Rates']];
      sheet.getRange(`A${currentRow}`).format.font.bold = true;
      sheet.getRange(`A${currentRow}`).format.font.italic = true;
      currentRow++;
      
      data.revenueItems.forEach((item, index) => {
        const itemName = item.name || `Revenue Item ${index + 1}`;
        sheet.getRange(`A${currentRow}`).values = [[`${itemName} - Growth Type`]];
        sheet.getRange(`B${currentRow}`).values = [[item.growthType || 'None']];
        this.cellTracker.recordCell(`revenue_${index}_growth_type`, 'Assumptions', `B${currentRow}`);
        currentRow++;
        
        if (item.growthType === 'annual' && item.annualGrowthRate) {
          sheet.getRange(`A${currentRow}`).values = [[`${itemName} - Annual Growth Rate (%)`]];
          sheet.getRange(`B${currentRow}`).values = [[item.annualGrowthRate]];
          this.cellTracker.recordCell(`revenue_${index}_growth_rate`, 'Assumptions', `B${currentRow}`);
          currentRow++;
        }
      });
      
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
      
      // Add growth rates for operating expenses
      currentRow++;
      sheet.getRange(`A${currentRow}`).values = [['Operating Expense Growth Rates']];
      sheet.getRange(`A${currentRow}`).format.font.bold = true;
      sheet.getRange(`A${currentRow}`).format.font.italic = true;
      currentRow++;
      
      data.operatingExpenses.forEach((item, index) => {
        const itemName = item.name || `OpEx Item ${index + 1}`;
        sheet.getRange(`A${currentRow}`).values = [[`${itemName} - Growth Type`]];
        sheet.getRange(`B${currentRow}`).values = [[item.growthType || 'None']];
        this.cellTracker.recordCell(`opex_${index}_growth_type`, 'Assumptions', `B${currentRow}`);
        currentRow++;
        
        if (item.growthType === 'annual' && item.annualGrowthRate) {
          sheet.getRange(`A${currentRow}`).values = [[`${itemName} - Annual Growth Rate (%)`]];
          sheet.getRange(`B${currentRow}`).values = [[item.annualGrowthRate]];
          this.cellTracker.recordCell(`opex_${index}_growth_rate`, 'Assumptions', `B${currentRow}`);
          currentRow++;
        }
      });
      
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
      
      // Add growth rates for capital expenses
      currentRow++;
      sheet.getRange(`A${currentRow}`).values = [['Capital Expense Growth Rates']];
      sheet.getRange(`A${currentRow}`).format.font.bold = true;
      sheet.getRange(`A${currentRow}`).format.font.italic = true;
      currentRow++;
      
      data.capitalExpenses.forEach((item, index) => {
        const itemName = item.name || `CapEx Item ${index + 1}`;
        sheet.getRange(`A${currentRow}`).values = [[`${itemName} - Growth Type`]];
        sheet.getRange(`B${currentRow}`).values = [[item.growthType || 'None']];
        this.cellTracker.recordCell(`capex_${index}_growth_type`, 'Assumptions', `B${currentRow}`);
        currentRow++;
        
        if (item.growthType === 'annual' && item.annualGrowthRate) {
          sheet.getRange(`A${currentRow}`).values = [[`${itemName} - Annual Growth Rate (%)`]];
          sheet.getRange(`B${currentRow}`).values = [[item.annualGrowthRate]];
          this.cellTracker.recordCell(`capex_${index}_growth_rate`, 'Assumptions', `B${currentRow}`);
          currentRow++;
        }
      });
      
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

4. **EBITDA Calculation:**
   - EBITDA = Total Revenue + Total Operating Expenses (expenses are negative)

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
   - Net Income = EBITDA - CapEx - Interest Expense

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
      const growthTypeRef = this.cellTracker.getCellReference(`revenue_${index}_growth_type`);
      const growthRateRef = this.cellTracker.getCellReference(`revenue_${index}_growth_rate`);
      
      output += `\n- ${item.name || `Revenue Item ${index + 1}`}:\n`;
      output += `  * Base Value: ${valueRef}\n`;
      output += `  * Growth Type: ${item.growthType || 'none'}\n`;
      if (item.growthType === 'annual' && item.annualGrowthRate) {
        output += `  * Annual Growth Rate: ${item.annualGrowthRate}%\n`;
      }
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
      
      output += `\n- ${item.name || `OpEx Item ${index + 1}`}:\n`;
      output += `  * Base Value: ${valueRef}\n`;
      output += `  * Growth Type: ${item.growthType || 'none'}\n`;
      if (item.growthType === 'annual' && item.annualGrowthRate) {
        output += `  * Annual Growth Rate: ${item.annualGrowthRate}%\n`;
      }
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
      
      output += `\n- ${item.name || `CapEx Item ${index + 1}`}:\n`;
      output += `  * Base Value: ${valueRef}\n`;
      output += `  * Growth Type: ${item.growthType || 'none'}\n`;
      if (item.growthType === 'annual' && item.annualGrowthRate) {
        output += `  * Annual Growth Rate: ${item.annualGrowthRate}%\n`;
      }
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
  
  // Generate P&L with formulas (without AI for now)
  async generatePLWithAI(modelData) {
    try {
      console.log('📈 Generating P&L Statement...');
      
      // Create the actual P&L sheet with formulas
      await this.createPLSheet(modelData);
      
      return { success: true, message: 'P&L Statement generated successfully!' };
      
    } catch (error) {
      console.error('❌ Error generating P&L:', error);
      return { success: false, error: error.message };
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
      
      // Step 4: Create FCF sheet that shows the AI prompt (for now)
      // In production, this would send to OpenAI and parse the response
      await this.createAIFCFSheet(modelData, fcfPrompt);
      
      console.log('📋 REAL FCF AI Prompt for OpenAI:');
      console.log('='.repeat(100));
      console.log(fcfPrompt);
      console.log('='.repeat(100));
      
      return { success: true, message: 'FCF AI prompt generated with REAL P&L cell references!' };
      
    } catch (error) {
      console.error('❌ Error generating FCF:', error);
      return { success: false, error: error.message };
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
      const periodColumns = Math.min(periods, 36); // Show up to 36 periods
      
      let currentRow = 1;
      
      // TITLE
      plSheet.getRange('A1').values = [['P&L Statement']];
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
          for (let col = 1; col <= periodColumns; col++) {
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
              // Growth formula for subsequent periods
              const prevCol = this.getColumnLetter(col - 1);
              const growthRate = item.annualGrowthRate || 0;
              const periodAdjustment = this.getPeriodAdjustment(modelData.modelPeriods);
              
              if (growthRate > 0) {
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
                  [[`=${prevCol}${currentRow}*(1+${growthRate}${periodAdjustment}/100)`]];
              } else {
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
          for (let col = 1; col <= periodColumns; col++) {
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
              // Growth formula for subsequent periods
              const prevCol = this.getColumnLetter(col - 1);
              const growthRate = item.annualGrowthRate || 0;
              const periodAdjustment = this.getPeriodAdjustment(modelData.modelPeriods);
              
              if (growthRate > 0) {
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
                  [[`=${prevCol}${currentRow}*(1+${growthRate}${periodAdjustment}/100)`]];
              } else {
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
      
      // EBITDA
      plSheet.getRange(`A${currentRow}`).values = [['EBITDA']];
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
      
      // Format numbers
      const dataRange = plSheet.getRange(`B5:${this.getColumnLetter(periodColumns)}${currentRow}`);
      dataRange.numberFormat = [['#,##0;[Red](#,##0)']];
      
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
- Capital Expenditures: Use CapEx from assumptions
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
Row 15: EBITDA from P&L
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
    
    output += `**CAPITAL EXPENDITURE ASSUMPTIONS:**\n`;
    if (modelData.capitalExpenses && modelData.capitalExpenses.length > 0) {
      modelData.capitalExpenses.forEach((item, index) => {
        output += `- ${item.name}: ${this.cellTracker.getCellReference(`capex_${index}`)}\n`;
        output += `  * Growth Type: ${item.growthType || 'none'}\n`;
        if (item.growthType === 'annual' && item.annualGrowthRate) {
          output += `  * Annual Growth Rate: ${item.annualGrowthRate}%\n`;
        }
      });
    } else {
      output += `- No capital expenditure items specified\n`;
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
    output += `- EBITDA: ${this.plCellTracker.getCellReference('ebitda')}\n`;
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
      output = 'No capital expenditures specified.';
    }
    return output;
  }
  
  // Create AI-generated FCF sheet (placeholder for now)
  async createAIFCFSheet(modelData, aiPrompt) {
    return Excel.run(async (context) => {
      console.log('💰 Creating Free Cash Flow sheet...');
      
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
      const periodColumns = Math.min(periods, 36);
      
      let currentRow = 1;
      
      // TITLE
      fcfSheet.getRange('A1').values = [['Free Cash Flow Statement']];
      fcfSheet.getRange('A1').format.font.bold = true;
      fcfSheet.getRange('A1').format.font.size = 16;
      currentRow = 3;
      
      // TIME PERIOD HEADERS
      const headers = [''];
      const startDate = new Date(modelData.projectStartDate);
      for (let i = 0; i < periodColumns; i++) {
        headers.push(this.formatPeriodHeader(startDate, i, modelData.modelPeriods));
      }
      
      const headerRange = fcfSheet.getRange(`A${currentRow}:${this.getColumnLetter(periodColumns)}${currentRow}`);
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = '#e0e0e0';
      currentRow += 2;
      
      // OPERATING CASH FLOW SECTION
      fcfSheet.getRange(`A${currentRow}`).values = [['OPERATING CASH FLOW']];
      fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
      fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#87CEEB';
      currentRow++;
      
      // EBITDA (from P&L)
      fcfSheet.getRange(`A${currentRow}`).values = [['EBITDA']];
      const ebitdaRef = this.plCellTracker.getCellReference('ebitda');
      if (ebitdaRef) {
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          const plCol = this.getColumnLetter(col + 1); // P&L starts from column B
          fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`='P&L Statement'!${plCol}${ebitdaRef.split('!')[1].match(/\\d+/)[0]}`]];
        }
      }
      currentRow++;
      
      // Tax (25% of EBITDA)
      fcfSheet.getRange(`A${currentRow}`).values = [['Less: Tax (25%)']];
      for (let col = 1; col <= periodColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=-${colLetter}${currentRow-1}*0.25`]];
      }
      currentRow++;
      
      // NOPAT
      fcfSheet.getRange(`A${currentRow}`).values = [['NOPAT']];
      fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
      for (let col = 1; col <= periodColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${currentRow-2}+${colLetter}${currentRow-1}`]];
      }
      currentRow += 2;
      
      // Working Capital Change
      fcfSheet.getRange(`A${currentRow}`).values = [['Less: Change in Working Capital']];
      const totalRevenueRef = this.plCellTracker.getCellReference('total_revenue');
      if (totalRevenueRef) {
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          const plCol = this.getColumnLetter(col + 1);
          if (col === 1) {
            // First period - just 2% of revenue
            fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=-'P&L Statement'!${plCol}${totalRevenueRef.split('!')[1].match(/\\d+/)[0]}*0.02`]];
          } else {
            // Change from previous period
            const prevPlCol = this.getColumnLetter(col);
            fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=-('P&L Statement'!${plCol}${totalRevenueRef.split('!')[1].match(/\\d+/)[0]}*0.02-'P&L Statement'!${prevPlCol}${totalRevenueRef.split('!')[1].match(/\\d+/)[0]}*0.02)`]];
          }
        }
      }
      currentRow++;
      
      // Capital Expenditures
      fcfSheet.getRange(`A${currentRow}`).values = [['Less: Capital Expenditures']];
      if (modelData.capitalExpenses && modelData.capitalExpenses.length > 0) {
        // Sum all CapEx items
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          let capexFormula = '0';
          modelData.capitalExpenses.forEach((item, index) => {
            const capexRef = this.cellTracker.getCellReference(`capex_${index}`);
            if (capexRef) {
              if (capexFormula === '0') {
                capexFormula = `-${capexRef}`;
              } else {
                capexFormula += `-${capexRef}`;
              }
            }
          });
          fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[capexFormula]];
        }
      } else {
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          fcfSheet.getRange(`${colLetter}${currentRow}`).values = [[0]];
        }
      }
      currentRow += 2;
      
      // Unlevered Free Cash Flow
      fcfSheet.getRange(`A${currentRow}`).values = [['Unlevered Free Cash Flow']];
      fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
      fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#98FB98';
      for (let col = 1; col <= periodColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${currentRow-5}+${colLetter}${currentRow-3}+${colLetter}${currentRow-1}`]];
      }
      currentRow += 2;
      
      // Interest Payments
      fcfSheet.getRange(`A${currentRow}`).values = [['Less: Interest Payments']];
      const interestRef = this.plCellTracker.getCellReference('interest_expense');
      if (interestRef) {
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          const plCol = this.getColumnLetter(col + 1);
          fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`='P&L Statement'!${plCol}${interestRef.split('!')[1].match(/\\d+/)[0]}`]];
        }
      }
      currentRow++;
      
      // Levered Free Cash Flow
      fcfSheet.getRange(`A${currentRow}`).values = [['Levered Free Cash Flow']];
      fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
      fcfSheet.getRange(`A${currentRow}`).format.fill.color = '#FFD700';
      for (let col = 1; col <= periodColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${currentRow-3}+${colLetter}${currentRow-1}`]];
      }
      currentRow += 2;
      
      // Cumulative FCF
      fcfSheet.getRange(`A${currentRow}`).values = [['Cumulative FCF']];
      fcfSheet.getRange(`A${currentRow}`).format.font.bold = true;
      for (let col = 1; col <= periodColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (col === 1) {
          fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${colLetter}${currentRow-2}`]];
        } else {
          const prevCol = this.getColumnLetter(col - 1);
          fcfSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=${prevCol}${currentRow}+${colLetter}${currentRow-2}`]];
        }
      }
      
      // Format numbers
      const dataRange = fcfSheet.getRange(`B5:${this.getColumnLetter(periodColumns)}${currentRow}`);
      dataRange.numberFormat = [['#,##0;[Red](#,##0)']];
      
      // Auto-fit columns
      fcfSheet.getRange(`A:${this.getColumnLetter(periodColumns)}`).format.autofitColumns();
      
      await context.sync();
      // Add AI prompt information
      fcfSheet.getRange('A1').values = [['Free Cash Flow - AI Generated']];
      fcfSheet.getRange('A1').format.font.bold = true;
      fcfSheet.getRange('A1').format.font.size = 16;
      fcfSheet.getRange('A1').format.fill.color = '#FFD700';
      
      fcfSheet.getRange('A3').values = [['This FCF should be generated by sending the prompt below to OpenAI:']];
      fcfSheet.getRange('A3').format.font.bold = true;
      fcfSheet.getRange('A3').format.fill.color = '#FFFF00';
      
      // Add truncated prompt for display
      const promptPreview = aiPrompt.substring(0, 1000) + '...';
      fcfSheet.getRange('A5').values = [[promptPreview]];
      fcfSheet.getRange('A5:A20').merge();
      fcfSheet.getRange('A5').format.wrapText = true;
      
      fcfSheet.getRange('A22').values = [['Key P&L References Available:']];
      fcfSheet.getRange('A22').format.font.bold = true;
      
      let refRow = 23;
      fcfSheet.getRange(`A${refRow}`).values = [[`EBITDA: ${this.plCellTracker.getCellReference('ebitda')}`]];
      refRow++;
      fcfSheet.getRange(`A${refRow}`).values = [[`Net Income: ${this.plCellTracker.getCellReference('net_income')}`]];
      refRow++;
      fcfSheet.getRange(`A${refRow}`).values = [[`Interest Expense: ${this.plCellTracker.getCellReference('interest_expense')}`]];
      refRow++;
      fcfSheet.getRange(`A${refRow}`).values = [[`Total Revenue: ${this.plCellTracker.getCellReference('total_revenue')}`]];
      
      fcfSheet.getRange('A:A').format.autofitColumns();
      
      console.log('✅ FCF AI prompt sheet created successfully');
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
      
      // Calculate number of periods
      const periods = this.calculatePeriods(modelData.projectStartDate, modelData.projectEndDate, modelData.modelPeriods);
      const periodColumns = Math.min(periods, 36); // Show more periods
      
      // HEADER
      plSheet.getRange('A1').values = [['P&L Statement']];
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
            for (let col = 1; col <= periodColumns; col++) {
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
      
      // Store total revenue row for EBITDA calculation
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
            for (let col = 1; col <= periodColumns; col++) {
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
      
      // Store total opex row for EBITDA calculation
      const totalOpexRow = opexCount > 0 ? currentRow - 2 : 0;
      
      // EBITDA CALCULATION
      plSheet.getRange(`A${currentRow}`).values = [['EBITDA']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#98FB98';
      
      // EBITDA formulas for each period
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
              interestFormula = `=-${debtCellRef}*${rateCellRef}/100/12`;
              break;
            case 'quarterly':
              interestFormula = `=-${debtCellRef}*${rateCellRef}/100/4`;
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
      
      // Find EBITDA row (it's 2 rows up from current, or 4 if we have debt)
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
      dataRange.numberFormat = [['#,##0.00;[Red](#,##0.00)']];
      
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
        return Math.min(diffDays, 365);
      case 'monthly':
        return Math.min(Math.ceil(diffDays / 30), 60);
      case 'quarterly':
        return Math.min(Math.ceil(diffDays / 90), 20);
      case 'yearly':
        return Math.min(Math.ceil(diffDays / 365), 10);
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
              if (lowerValue.includes('ebitda') || (lowerValue.includes('ebit') && lowerValue.includes('da'))) {
                structure.lineItems.ebitda = { row: row + 1, startCol: 'B', cellRef: `B${row + 1}` };
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
   - EBITDA: Reference the exact EBITDA row from P&L
   - Less: Tax (25% of EBITDA)
   - = NOPAT (Net Operating Profit After Tax)

2. **WORKING CAPITAL ADJUSTMENTS:**
   - Change in Working Capital (2% of Total Revenue change)
   - Calculate as: Current Period Revenue * 2% - Previous Period Revenue * 2%

3. **CAPITAL EXPENDITURES:**
   - Reference capital expense items from Assumptions if any exist
   - Apply growth rates if specified

4. **UNLEVERED FREE CASH FLOW:**
   - = NOPAT - Change in Working Capital - Capital Expenditures

5. **FINANCING CASH FLOWS:**
   - Interest Payments: Reference from P&L Interest Expense line
   - Principal Repayments (if applicable)

6. **LEVERED FREE CASH FLOW:**
   - = Unlevered FCF - Interest Payments - Principal Repayments

7. **CUMULATIVE METRICS:**
   - Cumulative FCF
   - IRR calculation base

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
    
    if (plStructure.lineItems.ebitda) {
      output += `- EBITDA: Row ${plStructure.lineItems.ebitda.row}, Range B${plStructure.lineItems.ebitda.row}:${this.getColumnLetter(maxPeriods)}${plStructure.lineItems.ebitda.row}\n`;
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
}

// Export for use in main application
window.ExcelGenerator = ExcelGenerator;
window.CellTracker = CellTracker;