/* global Office, Excel */

interface DealAssumptions {
  dealName: string;
  dealSize: number;
  ltv: number;
  holdingPeriod: number;
  revenueGrowth: number;
  exitMultiple: number;
  selectedRange?: string;
  rangeData?: any[][];
}

interface ChatMessage {
  role: 'user' | 'assistant';
  content: string;
}

class MAModelingAddin {
  private chatMessages: ChatMessage[] = [];
  private selectedRange: string | null = null;
  private uploadedFiles: File[] = [];

  constructor() {
    // Initialize when Office is ready
    Office.onReady(() => {
      this.initializeAddin();
    });
  }

  private initializeAddin() {
    // Set up event listeners
    document.getElementById('selectRangeBtn')?.addEventListener('click', () => this.selectAssumptionsRange());
    document.getElementById('generateModelBtn')?.addEventListener('click', () => this.generateModel());
    document.getElementById('validateModelBtn')?.addEventListener('click', () => this.validateModel());
    document.getElementById('sendChatBtn')?.addEventListener('click', () => this.sendChatMessage());
    
    // Allow Enter key in chat input
    document.getElementById('chatInput')?.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        this.sendChatMessage();
      }
    });

    // File upload event listeners
    this.initializeFileUpload();

    // API key is already configured
    console.log('OpenAI API key configured');
  }

  private initializeFileUpload() {
    const dropzone = document.getElementById('fileDropzone');
    const fileInput = document.getElementById('fileInput') as HTMLInputElement;
    const uploadLink = document.querySelector('.upload-link');

    // Dropzone click handler
    dropzone?.addEventListener('click', () => {
      fileInput?.click();
    });

    // Upload link click handler
    uploadLink?.addEventListener('click', (e) => {
      e.stopPropagation();
      fileInput?.click();
    });

    // File input change handler
    fileInput?.addEventListener('change', (e) => {
      const files = (e.target as HTMLInputElement).files;
      if (files) {
        this.handleFileSelection(Array.from(files));
      }
    });

    // Drag and drop handlers
    dropzone?.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });

    dropzone?.addEventListener('dragleave', () => {
      dropzone.classList.remove('dragover');
    });

    dropzone?.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      const files = Array.from(e.dataTransfer?.files || []);
      this.handleFileSelection(files);
    });
  }

  private handleFileSelection(files: File[]) {
    // Filter valid files (PDF and CSV only)
    const validFiles = files.filter(file => {
      const isValidType = file.type === 'application/pdf' || file.type === 'text/csv' || file.name.endsWith('.csv');
      const isValidSize = file.size <= 10 * 1024 * 1024; // 10MB limit
      return isValidType && isValidSize;
    });

    // Check total file limit
    if (this.uploadedFiles.length + validFiles.length > 4) {
      this.addChatMessage('assistant', 'Maximum 4 files allowed. Please remove some files first.');
      return;
    }

    // Add files to uploaded list
    this.uploadedFiles.push(...validFiles);
    this.updateFileDisplay();

    if (validFiles.length > 0) {
      this.addChatMessage('assistant', `Successfully uploaded ${validFiles.length} file(s). You can now ask me to analyze them and fill out your assumptions template!`);
    }
  }

  private updateFileDisplay() {
    const uploadedFilesDiv = document.getElementById('uploadedFiles');
    const fileList = document.getElementById('fileList');

    if (this.uploadedFiles.length === 0) {
      uploadedFilesDiv!.style.display = 'none';
      return;
    }

    uploadedFilesDiv!.style.display = 'block';
    fileList!.innerHTML = '';

    this.uploadedFiles.forEach((file, index) => {
      const fileItem = document.createElement('div');
      fileItem.className = 'file-item';
      fileItem.innerHTML = `
        <div class="file-info">
          <svg class="file-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14,2 14,8 20,8"></polyline>
          </svg>
          <div>
            <div class="file-name">${file.name}</div>
            <div class="file-size">${this.formatFileSize(file.size)}</div>
          </div>
        </div>
        <button class="remove-file" data-index="${index}">Remove</button>
      `;

      const removeBtn = fileItem.querySelector('.remove-file');
      removeBtn?.addEventListener('click', () => this.removeFile(index));

      fileList!.appendChild(fileItem);
    });
  }

  private removeFile(index: number) {
    this.uploadedFiles.splice(index, 1);
    this.updateFileDisplay();
  }

  private formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  private async selectAssumptionsRange() {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(['address', 'values', 'text']);
        
        await context.sync();
        
        this.selectedRange = range.address;
        const rangeData = range.values;
        
        // Update UI
        const statusElement = document.getElementById('rangeStatus');
        if (statusElement) {
          statusElement.textContent = `Selected: ${range.address} (${rangeData.length} rows)`;
        }
        
        // Try to parse assumptions from selected range
        this.parseAssumptionsFromRange(rangeData);
      });
    } catch (error) {
      console.error('Error selecting range:', error);
      this.showStatus('Error selecting range. Please try again.');
    }
  }

  private parseAssumptionsFromRange(rangeData: any[][]) {
    // Smart parsing of assumptions from Excel range
    const assumptions: any = {};
    
    for (let i = 0; i < rangeData.length; i++) {
      const row = rangeData[i];
      if (row.length >= 2) {
        const label = String(row[0]).toLowerCase();
        const value = row[1];
        
        // Map common assumption labels to our fields
        if (label.includes('deal') && label.includes('size')) {
          assumptions.dealSize = parseFloat(value) || 0;
        } else if (label.includes('ltv') || label.includes('leverage')) {
          assumptions.ltv = parseFloat(value) * (value <= 1 ? 100 : 1); // Convert decimal to percentage
        } else if (label.includes('holding') || label.includes('period')) {
          assumptions.holdingPeriod = parseFloat(value) || 0;
        } else if (label.includes('revenue') && label.includes('growth')) {
          assumptions.revenueGrowth = parseFloat(value) * (value <= 1 ? 100 : 1);
        } else if (label.includes('exit') && label.includes('multiple')) {
          assumptions.exitMultiple = parseFloat(value) || 0;
        }
      }
    }
    
    // Update form fields with parsed values
    this.updateFormWithAssumptions(assumptions);
  }

  private updateFormWithAssumptions(assumptions: any) {
    const fields = ['dealSize', 'ltv', 'holdingPeriod', 'revenueGrowth', 'exitMultiple'];
    
    fields.forEach(field => {
      const element = document.getElementById(field) as HTMLInputElement;
      if (element && assumptions[field] !== undefined) {
        element.value = assumptions[field].toString();
      }
    });
  }

  private collectAssumptions(): DealAssumptions {
    const getValue = (id: string): string => {
      const element = document.getElementById(id) as HTMLInputElement;
      return element?.value || '';
    };

    return {
      dealName: getValue('dealName') || 'Sample Deal',
      dealSize: parseFloat(getValue('dealSize')) || 50,
      ltv: parseFloat(getValue('ltv')) || 70,
      holdingPeriod: parseFloat(getValue('holdingPeriod')) || 60,
      revenueGrowth: parseFloat(getValue('revenueGrowth')) || 15,
      exitMultiple: parseFloat(getValue('exitMultiple')) || 12,
      selectedRange: this.selectedRange || undefined
    };
  }

  private async generateModel() {
    const assumptions = this.collectAssumptions();
    
    this.showLoading(true);
    this.showStatus('Generating M&A model sheets...');
    
    try {
      await Excel.run(async (context) => {
        // Create or clear sheets
        await this.createModelSheets(context, assumptions);
        
        // Generate cash flows and calculate metrics
        const metrics = await this.calculateMetrics(assumptions);
        
        // Populate sheets with data and formulas
        await this.populateAssumptionsSheet(context, assumptions);
        await this.populateNPVSheet(context, assumptions, metrics);
        await this.populatePLSheet(context, assumptions);
        await this.populateCFSheet(context, assumptions);
        
        await context.sync();
        
        this.showStatus('✅ M&A model generated successfully!');
        this.addChatMessage('assistant', `M&A model for "${assumptions.dealName}" has been generated with the following key metrics:\n• Levered IRR: ${(metrics.leveredIRR * 100).toFixed(1)}%\n• MOIC: ${metrics.moic.toFixed(1)}x\n• Deal Size: $${assumptions.dealSize}M`);
      });
    } catch (error) {
      console.error('Error generating model:', error);
      this.showStatus('❌ Error generating model. Please check your inputs.');
    } finally {
      this.showLoading(false);
    }
  }

  private async createModelSheets(context: Excel.RequestContext, assumptions: DealAssumptions) {
    const workbook = context.workbook;
    const sheetNames = ['Assumptions', 'P&L', 'CF Statement', 'NPV', 'Outputs'];
    
    // Create sheets if they don't exist
    for (const sheetName of sheetNames) {
      try {
        let sheet = workbook.worksheets.getItem(sheetName);
        sheet.load('name');
        await context.sync();
        
        // Clear existing content
        const usedRange = sheet.getUsedRange();
        usedRange.clear();
      } catch {
        // Sheet doesn't exist, create it
        workbook.worksheets.add(sheetName);
      }
    }
  }

  private async populateAssumptionsSheet(context: Excel.RequestContext, assumptions: DealAssumptions) {
    const sheet = context.workbook.worksheets.getItem('Assumptions');
    
    // Header
    sheet.getRange('A1:F1').merge();
    sheet.getRange('A1').values = [[`${assumptions.dealName} - Key Assumptions`]];
    sheet.getRange('A1').format.font.bold = true;
    sheet.getRange('A1').format.font.size = 16;
    
    // Deal parameters
    const data = [
      ['', ''],
      ['Deal Information', ''],
      ['Deal Size ($M)', assumptions.dealSize],
      ['LTV (%)', assumptions.ltv],
      ['Holding Period (Months)', assumptions.holdingPeriod],
      ['Revenue Growth (% annually)', assumptions.revenueGrowth],
      ['Exit Multiple (x EBITDA)', assumptions.exitMultiple],
      ['', ''],
      ['Calculated Values', ''],
      ['Debt Amount ($M)', `=C4*C5/100`],
      ['Equity Amount ($M)', `=C4*(1-C5/100)`],
      ['Exit Year', `=C6/12`]
    ];
    
    sheet.getRange('B3:C14').values = data;
    
    // Format headers
    sheet.getRange('B4').format.fill.color = '#1F4E79';
    sheet.getRange('B4').format.font.color = 'white';
    sheet.getRange('B4').format.font.bold = true;
    
    sheet.getRange('B10').format.fill.color = '#1F4E79';
    sheet.getRange('B10').format.font.color = 'white';
    sheet.getRange('B10').format.font.bold = true;
    
    // Format input cells
    const inputRanges = ['C4:C8'];
    inputRanges.forEach(range => {
      sheet.getRange(range).format.fill.color = '#FDE9D9';
    });
    
    // Format calculated cells
    const calcRanges = ['C11:C13'];
    calcRanges.forEach(range => {
      sheet.getRange(range).format.fill.color = '#E6E6FA';
    });
  }

  private async populateNPVSheet(context: Excel.RequestContext, assumptions: DealAssumptions, metrics: any) {
    const sheet = context.workbook.worksheets.getItem('NPV');
    
    // Header
    sheet.getRange('A1').values = [['NPV Analysis & Returns']];
    sheet.getRange('A1').format.font.bold = true;
    sheet.getRange('A1').format.font.size = 16;
    
    // Key metrics section
    const metricsData = [
      ['Key Metrics', 'Value'],
      ['Levered IRR', metrics.leveredIRR],
      ['Unlevered IRR', metrics.unleveredIRR],
      ['MOIC', metrics.moic],
      ['Levered NPV ($M)', metrics.leveredNPV / 1000000],
      ['Unlevered NPV ($M)', metrics.unleveredNPV / 1000000]
    ];
    
    sheet.getRange('A3:B8').values = metricsData;
    
    // Format metrics
    sheet.getRange('A3:B3').format.fill.color = '#1F4E79';
    sheet.getRange('A3:B3').format.font.color = 'white';
    sheet.getRange('A3:B3').format.font.bold = true;
    
    // Format IRR cells as percentages
    sheet.getRange('B4:B5').numberFormat = [['0.00%']];
    
    // Format MOIC with x
    sheet.getRange('B6').numberFormat = [['0.0"x"']];
    
    // Format NPV as currency
    sheet.getRange('B7:B8').numberFormat = [['$#,##0.0']];
  }

  private async populatePLSheet(context: Excel.RequestContext, assumptions: DealAssumptions) {
    const sheet = context.workbook.worksheets.getItem('P&L');
    
    // Create a simplified P&L with formulas
    const months = assumptions.holdingPeriod;
    const headers = ['P&L Item', 'Formula/Notes'];
    
    // Add month columns
    for (let i = 0; i <= months; i++) {
      headers.push(`M${i}`);
    }
    
    sheet.getRange('A1').values = [headers];
    
    // Basic P&L structure
    const plData = [
      ['Revenue', 'Growing at specified rate'],
      ['Operating Expenses', 'Fixed + variable'],
      ['EBITDA', '=Revenue - OpEx'],
      ['Interest Expense', 'From debt model'],
      ['Net Income', '=EBITDA - Interest']
    ];
    
    sheet.getRange('A2:B6').values = plData;
    
    // Format headers
    sheet.getRange('A1:' + this.getColumnLetter(headers.length) + '1').format.fill.color = '#1F4E79';
    sheet.getRange('A1:' + this.getColumnLetter(headers.length) + '1').format.font.color = 'white';
    sheet.getRange('A1:' + this.getColumnLetter(headers.length) + '1').format.font.bold = true;
  }

  private async populateCFSheet(context: Excel.RequestContext, assumptions: DealAssumptions) {
    const sheet = context.workbook.worksheets.getItem('CF Statement');
    
    // Basic cash flow structure
    sheet.getRange('A1').values = [['Cash Flow Statement']];
    sheet.getRange('A1').format.font.bold = true;
    sheet.getRange('A1').format.font.size = 16;
    
    const cfData = [
      ['', ''],
      ['Operating Cash Flows', ''],
      ['EBITDA', 'Link to P&L'],
      ['Working Capital Changes', 'Assumption'],
      ['Operating CF', '=EBITDA + WC Changes'],
      ['', ''],
      ['Investment Cash Flows', ''],
      ['Initial Investment', -assumptions.dealSize * 1000000],
      ['Exit Proceeds', 'Final period only'],
      ['', ''],
      ['Financing Cash Flows', ''],
      ['Debt Drawn', 'Initial period'],
      ['Interest Payments', 'Monthly'],
      ['Principal Repayments', 'Monthly'],
      ['', ''],
      ['Net Cash Flow', 'Sum of all sections']
    ];
    
    sheet.getRange('A3:B18').values = cfData;
  }

  private async calculateMetrics(assumptions: DealAssumptions): Promise<any> {
    // Use the same calculation logic from the web app
    const cashFlows = this.generateCashFlows(assumptions);
    
    // Use fallback calculation for now
    return this.calculateMetricsFallback(cashFlows);
  }

  private generateCashFlows(assumptions: DealAssumptions): number[] {
    const periods = assumptions.holdingPeriod;
    const dealSize = assumptions.dealSize * 1000000;
    const equity = dealSize * (1 - assumptions.ltv / 100);
    const cashFlows: number[] = [];
    
    // Initial investment
    cashFlows.push(-equity);
    
    // Operating periods
    const monthlyEBITDA = dealSize * 0.2 / 12; // Assume 20% EBITDA margin
    for (let i = 1; i <= periods; i++) {
      const growthFactor = Math.pow(1 + assumptions.revenueGrowth / 100, (i - 1) / 12);
      let cf = monthlyEBITDA * growthFactor;
      
      // Add exit value in final period
      if (i === periods) {
        const annualEBITDA = monthlyEBITDA * 12 * growthFactor;
        const exitValue = annualEBITDA * assumptions.exitMultiple;
        cf += exitValue;
      }
      
      cashFlows.push(cf);
    }
    
    return cashFlows;
  }

  private calculateMetricsFallback(cashFlows: number[]): any {
    // Simple IRR approximation
    let irr = 0.1;
    for (let i = 0; i < 100; i++) {
      let npv = 0;
      let dnpv = 0;
      
      for (let j = 0; j < cashFlows.length; j++) {
        const period = j / 12;
        npv += cashFlows[j] / Math.pow(1 + irr, period);
        dnpv -= (period * cashFlows[j]) / Math.pow(1 + irr, period + 1);
      }
      
      if (Math.abs(npv) < 0.01) break;
      if (Math.abs(dnpv) < 0.01) break;
      
      irr = irr - npv / dnpv;
    }
    
    const moic = cashFlows[cashFlows.length - 1] / Math.abs(cashFlows[0]);
    const npv8 = cashFlows.reduce((sum, cf, i) => sum + cf / Math.pow(1.08, i / 12), 0);
    
    return {
      leveredIRR: irr,
      unleveredIRR: irr * 0.9, // Simplified
      moic: moic,
      leveredNPV: npv8,
      unleveredNPV: npv8 * 1.1
    };
  }

  private async validateModel() {
    this.addChatMessage('assistant', 'Model validation feature coming soon! This will check all formulas and cross-references for accuracy.');
  }

  private async sendChatMessage() {
    const input = document.getElementById('chatInput') as HTMLInputElement;
    const message = input.value.trim();
    
    if (!message) return;
    
    // Add user message
    this.addChatMessage('user', message);
    input.value = '';
    
    // Process with AI
    await this.processWithAI(message);
  }

  private async processWithAI(message: string) {
    try {
      const context = await this.getExcelContext();
      
      // Process uploaded files
      const fileContents = await this.processUploadedFiles();
      
      const response = await fetch('/.netlify/functions/chat', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          message: message,
          excelContext: context,
          fileContents: fileContents
        })
      });

      const data = await response.json();
      
      // Check if response includes commands
      if (data.commands && Array.isArray(data.commands)) {
        this.addChatMessage('assistant', data.response || 'Processing your request...');
        
        // Execute commands
        for (const command of data.commands) {
          await this.executeCommand(command);
        }
      } else {
        // Regular text response
        const responseText = data.response || data.error || 'No response received';
        this.addChatMessage('assistant', responseText);
        
        // Legacy formula detection for backwards compatibility
        if (responseText && typeof responseText === 'string' && responseText.includes('=') && responseText.includes('cell')) {
          this.offerToImplementFormula(responseText);
        }
      }
      
    } catch (error) {
      console.error('AI processing error:', error);
      
      // Try to get more specific error info
      let errorMessage = 'Sorry, I encountered an error. Please try again.';
      if (error instanceof Error) {
        errorMessage = `Error: ${error.message}`;
      }
      
      this.addChatMessage('assistant', errorMessage);
    }
  }

  private async processUploadedFiles(): Promise<string[]> {
    const fileContents: string[] = [];
    
    for (const file of this.uploadedFiles) {
      try {
        let content = '';
        
        if (file.type === 'text/csv' || file.name.endsWith('.csv')) {
          content = await this.readTextFile(file);
          // Limit CSV content to first 5000 characters to stay within token limits
          content = content.substring(0, 5000);
        } else if (file.type === 'application/pdf') {
          // For PDF files, we'll send the filename and indicate it needs server-side processing
          content = `[PDF FILE: ${file.name} - ${this.formatFileSize(file.size)}]`;
        }
        
        fileContents.push(`File: ${file.name}\nContent: ${content}`);
      } catch (error) {
        console.error(`Error processing file ${file.name}:`, error);
        fileContents.push(`File: ${file.name}\nError: Could not read file`);
      }
    }
    
    return fileContents;
  }

  private readTextFile(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target?.result as string);
      reader.onerror = reject;
      reader.readAsText(file);
    });
  }

  private async getExcelContext(): Promise<string> {
    try {
      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const selectedRange = context.workbook.getSelectedRange();
        const usedRange = worksheet.getUsedRange();
        
        worksheet.load(['name']);
        selectedRange.load(['address', 'values', 'formulas']);
        usedRange.load(['address', 'values', 'formulas']);
        
        await context.sync();
        
        return JSON.stringify({
          worksheetName: worksheet.name,
          selectedRange: {
            address: selectedRange.address,
            values: selectedRange.values,
            formulas: selectedRange.formulas
          },
          usedRange: {
            address: usedRange.address,
            values: usedRange.values.slice(0, 10), // Limit to first 10 rows
            formulas: usedRange.formulas.slice(0, 10)
          }
        });
      });
    } catch {
      return JSON.stringify({ error: 'Could not read Excel context' });
    }
  }

  private offerToImplementFormula(aiResponse: string) {
    if (confirm('Would you like me to implement the suggested formula in your selected cell?')) {
      this.implementSuggestedFormula(aiResponse);
    }
  }

  private async implementSuggestedFormula(aiResponse: string) {
    // Extract formula from AI response (simple pattern matching)
    const formulaMatch = aiResponse.match(/=([^"]+)/);
    if (!formulaMatch) return;
    
    const formula = '=' + formulaMatch[1].trim();
    
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.formulas = [[formula]];
        await context.sync();
        
        this.addChatMessage('assistant', `✅ Formula implemented: ${formula}`);
      });
    } catch (error) {
      this.addChatMessage('assistant', `❌ Error implementing formula: ${error}`);
    }
  }

  private addChatMessage(role: 'user' | 'assistant', content: string) {
    this.chatMessages.push({ role, content });
    
    const messagesDiv = document.getElementById('chatMessages');
    if (messagesDiv) {
      const messageBubble = document.createElement('div');
      messageBubble.className = `message-bubble ${role}-bubble`;
      
      const messageContent = document.createElement('div');
      messageContent.className = 'message-content';
      
      const messageText = document.createElement('div');
      messageText.className = 'message-text';
      messageText.textContent = content;
      
      messageContent.appendChild(messageText);
      messageBubble.appendChild(messageContent);
      messagesDiv.appendChild(messageBubble);
      
      messagesDiv.scrollTop = messagesDiv.scrollHeight;
    }
  }

  private showLoading(show: boolean) {
    const loading = document.getElementById('loading');
    if (loading) {
      loading.style.display = show ? 'block' : 'none';
    }
  }

  private showStatus(message: string) {
    console.log('Status:', message);
    // Could show in a status bar
  }

  private async executeCommand(command: any) {
    try {
      await Excel.run(async (context) => {
        switch (command.action) {
          case 'setValue':
            await this.setValueCommand(context, command.cell, command.value);
            break;
          case 'addToCell':
            await this.addToCellCommand(context, command.cell, command.value);
            break;
          case 'setFormula':
            await this.setFormulaCommand(context, command.cell, command.formula);
            break;
          case 'formatCell':
            await this.formatCellCommand(context, command.cell, command.format);
            break;
          case 'generateAssumptionsTemplate':
            await this.generateAssumptionsTemplate(context);
            break;
          case 'fillAssumptionsData':
            await this.fillAssumptionsData(context, command.data);
            break;
          default:
            console.warn('Unknown command action:', command.action);
        }
        
        await context.sync();
        
        if (command.action === 'generateAssumptionsTemplate') {
          this.addChatMessage('assistant', `✅ Generated M&A assumptions template`);
        } else if (command.action === 'fillAssumptionsData') {
          this.addChatMessage('assistant', `✅ Filled assumptions template with sample data`);
        } else {
          this.addChatMessage('assistant', `✅ Executed: ${command.action} on ${command.cell}`);
        }
      });
    } catch (error) {
      console.error('Command execution error:', error);
      this.addChatMessage('assistant', `❌ Error executing ${command.action}: ${error}`);
    }
  }

  private async setValueCommand(context: Excel.RequestContext, cellAddress: string, value: any) {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(cellAddress);
    range.values = [[value]];
  }

  private async addToCellCommand(context: Excel.RequestContext, cellAddress: string, addValue: number) {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(cellAddress);
    range.load(['values']);
    await context.sync();
    
    const currentValue = parseFloat(range.values[0][0]) || 0;
    const newValue = currentValue + addValue;
    range.values = [[newValue]];
  }

  private async setFormulaCommand(context: Excel.RequestContext, cellAddress: string, formula: string) {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(cellAddress);
    range.formulas = [[formula]];
  }

  private async formatCellCommand(context: Excel.RequestContext, cellAddress: string, format: any) {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(cellAddress);
    
    if (format.bold !== undefined) {
      range.format.font.bold = format.bold;
    }
    if (format.color) {
      range.format.font.color = format.color;
    }
    if (format.backgroundColor) {
      range.format.fill.color = format.backgroundColor;
    }
    if (format.numberFormat) {
      range.numberFormat = [[format.numberFormat]];
    }
  }

  private async generateAssumptionsTemplate(context: Excel.RequestContext) {
    try {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Step 1: Add basic content first
      const templateData = [
        ["Sample Company Ltd. - Key Assumptions", ""],
        ["Deal type", ""],
        ["Sector/Industry", ""],
        ["Geography", ""],
        ["Business Model", ""],
        ["Ownership Structure", ""],
        ["Purchase Price ($M)", ""],
        ["", ""],
        ["Acquisition Assumptions", ""],
        ["Acquisition date", ""],
        ["Holding Period (Months)", ""],
        ["Currency", ""],
        ["Transaction Fees", ""],
        ["Acquisition LTV", ""],
        ["Equity Contribution", "=B7*(1-B14)"],
        ["Debt Financing", "=B7*B14"],
        ["Debt Issuance Fees", ""],
        ["Interest Rate Margin", ""],
        ["", ""],
        ["Cost Items", ""],
        ["Staff expenses", ""],
        ["Salary Growth (p.a.)", ""],
        ["Cost Item 1", ""],
        ["Cost Item 2", ""],
        ["Cost Item 3", ""],
        ["Cost Item 4", ""],
        ["Cost Item 5", ""],
        ["Cost Item 6", ""],
        ["", ""],
        ["Exit Assumptions", ""],
        ["Disposal Costs", ""],
        ["Terminal Equity Multiple", ""],
        ["Terminal EBITDA", ""],
        ["Sale Price", ""]
      ];
      
      // Add all data at once
      worksheet.getRange("A1:B34").values = templateData;
      await context.sync();
      
      // Step 2: Apply formatting
      // Set Times New Roman font for all cells
      worksheet.getRange("A1:B34").format.font.name = "Times New Roman";
      worksheet.getRange("A1:B34").format.font.size = 12;
      
      // Title formatting - merge cells and format
      worksheet.getRange("A1:B1").merge();
      worksheet.getRange("A1").format.font.bold = true;
      worksheet.getRange("A1").format.fill.color = "#D9D9D9";
      
      // Section headers - merge cells and format
      worksheet.getRange("A9:B9").merge();
      worksheet.getRange("A9").format.fill.color = "#1F4E79";
      worksheet.getRange("A9").format.font.color = "white";
      worksheet.getRange("A9").format.font.bold = true;
      
      worksheet.getRange("A20:B20").merge();
      worksheet.getRange("A20").format.fill.color = "#1F4E79";
      worksheet.getRange("A20").format.font.color = "white";
      worksheet.getRange("A20").format.font.bold = true;
      
      worksheet.getRange("A30:B30").merge();
      worksheet.getRange("A30").format.fill.color = "#1F4E79";
      worksheet.getRange("A30").format.font.color = "white";
      worksheet.getRange("A30").format.font.bold = true;
      
      await context.sync();
      
      // Step 3: Set column widths and row heights
      worksheet.getRange("A:A").format.columnWidth = 200;
      worksheet.getRange("B:B").format.columnWidth = 150;
      
      // Set row heights for empty rows before section headers
      worksheet.getRange("8:8").format.rowHeight = 5; // Before "Acquisition Assumptions"
      worksheet.getRange("19:19").format.rowHeight = 5; // Before "Cost Items"
      worksheet.getRange("29:29").format.rowHeight = 5; // Before "Exit Assumptions"
      
      await context.sync();
      
    } catch (error) {
      console.error("Template generation error:", error);
      throw error;
    }
  }

  private async fillAssumptionsData(context: Excel.RequestContext, data: any) {
    try {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Map the data to the corresponding Excel cells
      const cellValues = [
        ["B2", data.dealType || "Business Acquisition"],
        ["B3", data.sector || "Technology"],
        ["B4", data.geography || "United States"],
        ["B5", data.businessModel || "SaaS"],
        ["B6", data.ownership || "Private Equity"],
        ["B7", data.purchasePrice || 100],
        ["B10", data.acquisitionDate || "31/03/2025"],
        ["B11", data.holdingPeriod || 60],
        ["B12", data.currency || "USD"],
        ["B13", data.transactionFees || "1.5%"],
        ["B14", data.acquisitionLTV || "75%"],
        ["B17", data.debtIssuanceFees || "1.0%"],
        ["B18", data.interestRateMargin || "3.5%"],
        ["B21", data.staffExpenses || 5000000],
        ["B22", data.salaryGrowth || "3.0%"],
        ["B23", data.costItem1 || 2000000],
        ["B24", data.costItem2 || 800000],
        ["B25", data.costItem3 || 1200000],
        ["B26", data.costItem4 || 400000],
        ["B27", data.costItem5 || 600000],
        ["B28", data.costItem6 || 300000],
        ["B31", data.disposalCosts || "0.5%"],
        ["B32", data.terminalEquityMultiple || 12.5],
        ["B33", data.terminalEBITDA || 15000000],
        ["B34", data.salePrice || 187500000]
      ];
      
      // Fill all data at once using bulk operations
      for (const [cell, value] of cellValues) {
        worksheet.getRange(cell).values = [[value]];
      }
      
      await context.sync();
      
    } catch (error) {
      console.error("Data filling error:", error);
      throw error;
    }
  }

  private getColumnLetter(col: number): string {
    let letter = '';
    while (col > 0) {
      const mod = (col - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      col = Math.floor((col - 1) / 26);
    }
    return letter;
  }
}

// Initialize the add-in
new MAModelingAddin();