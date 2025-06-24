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
      
      // Step 2: Create P&L sheet with dynamic formulas
      await this.createPLSheet(modelData);
      
      console.log('✅ Model generation completed successfully!');
      this.cellTracker.printAllCells();
      
      // Log the AI prompt for debugging
      console.log('📋 AI Prompt that would be used:');
      console.log(this.generateAIPrompt(modelData));
      
      return { success: true, message: 'Model with P&L created successfully!' };
      
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

  // Generate OpenAI prompt with all assumptions and cell references
  generateAIPrompt(modelData) {
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
    
    const prompt = `You are an analyst at a top investment bank. You have been given the following M&A model assumptions with their Excel cell references:

**HIGH-LEVEL PARAMETERS:**
- Currency: ${modelData.currency} (Cell: ${this.cellTracker.getCellReference('currency')})
- Project Start Date: ${modelData.projectStartDate} (Cell: ${this.cellTracker.getCellReference('projectStartDate')})
- Project End Date: ${modelData.projectEndDate} (Cell: ${this.cellTracker.getCellReference('projectEndDate')})
- Model Periods: ${modelData.modelPeriods} (Cell: ${this.cellTracker.getCellReference('modelPeriods')})

**DEAL ASSUMPTIONS:**
- Deal Name: ${modelData.dealName} (Cell: ${this.cellTracker.getCellReference('dealName')})
- Deal Value: ${modelData.dealValue} (Cell: ${this.cellTracker.getCellReference('dealValue')})
- Transaction Fee %: ${modelData.transactionFee} (Cell: ${this.cellTracker.getCellReference('transactionFee')})
- Deal LTV %: ${modelData.dealLTV} (Cell: ${this.cellTracker.getCellReference('dealLTV')})
- Equity Contribution: (Cell: ${this.cellTracker.getCellReference('equityContribution')})
- Debt Financing: (Cell: ${this.cellTracker.getCellReference('debtFinancing')})

**REVENUE ITEMS:**
${this.formatRevenueItems(modelData)}

**OPERATING EXPENSES:**
${this.formatOpexItems(modelData)}

**CAPITAL EXPENSES:**
${this.formatCapexItems(modelData)}

**EXIT ASSUMPTIONS:**
- Disposal Cost %: ${modelData.disposalCost} (Cell: ${this.cellTracker.getCellReference('disposalCost')})
- Terminal Cap Rate %: ${modelData.terminalCapRate} (Cell: ${this.cellTracker.getCellReference('terminalCapRate')})

**DEBT MODEL:**
${this.formatDebtModel(modelData)}

**INSTRUCTIONS:**
Please generate a detailed P&L Statement structure with the following requirements:
1. Time periods should be ${modelData.modelPeriods} from ${modelData.projectStartDate} to ${modelData.projectEndDate}
2. All values should be Excel formulas referencing the assumption cells above
3. Growth rates are annual, so adjust them for the period:
   - Daily: divide annual rate by 365
   - Monthly: divide annual rate by 12
   - Quarterly: divide annual rate by 4
4. Include proper debt service calculations based on the debt model
5. Show Revenue, OpEx, EBITDA, Interest Expense, and Net Income
6. All formulas should reference the specific cells provided above

Please provide the P&L structure with exact Excel formulas for each line item.`;

    console.log('📝 Generated prompt:', prompt);
    return prompt;
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
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=Assumptions.${valueRef.split('!')[1]}`]];
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
        
        // Sum formulas for each period
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
            [[`=SUM(${colLetter}${revenueStartRow}:${colLetter}${currentRow - 1})`]];
        }
        currentRow += 2;
      }
      
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
                plSheet.getRange(`${colLetter}${currentRow}`).formulas = [[`=-Assumptions.${valueRef.split('!')[1]}`]];
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
        
        // Sum formulas for each period
        for (let col = 1; col <= periodColumns; col++) {
          const colLetter = this.getColumnLetter(col);
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
            [[`=SUM(${colLetter}${opexStartRow}:${colLetter}${currentRow - 1})`]];
        }
        currentRow += 2;
      }
      
      // EBITDA CALCULATION
      plSheet.getRange(`A${currentRow}`).values = [['EBITDA']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#98FB98';
      
      // Find the total revenue and total opex rows
      const totalRevenueRow = revenueStartRow + revenueCount + 1;
      const totalOpexRow = opexStartRow + opexCount + 1;
      
      // EBITDA formulas for each period
      for (let col = 1; col <= periodColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
          [[`=${colLetter}${totalRevenueRow}+${colLetter}${totalOpexRow}`]];
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
          const debtCellRef = debtRef ? `Assumptions.${debtRef.split('!')[1]}` : 'Assumptions.B8';
          const rateCellRef = interestRateRef ? `Assumptions.${interestRateRef.split('!')[1]}` : 'Assumptions.B15';
          
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
      
      // NET INCOME
      plSheet.getRange(`A${currentRow}`).values = [['Net Income']];
      plSheet.getRange(`A${currentRow}`).format.font.bold = true;
      plSheet.getRange(`A${currentRow}`).format.fill.color = '#FFD700';
      
      const ebitdaRow = hasDebt ? currentRow - 4 : currentRow - 2;
      const interestRow = hasDebt ? currentRow - 2 : 0;
      
      // Net Income formulas for each period
      for (let col = 1; col <= periodColumns; col++) {
        const colLetter = this.getColumnLetter(col);
        if (hasDebt) {
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
            [[`=${colLetter}${ebitdaRow}+${colLetter}${interestRow}`]];
        } else {
          plSheet.getRange(`${colLetter}${currentRow}`).formulas = 
            [[`=${colLetter}${ebitdaRow}`]];
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
}

// Export for use in main application
window.ExcelGenerator = ExcelGenerator;
window.CellTracker = CellTracker;