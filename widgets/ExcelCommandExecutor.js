class ExcelCommandExecutor {
  constructor() {
    this.commandHistory = [];
    this.currentIndex = -1;
    this.maxHistorySize = 50;
    this.isExecuting = false;
  }

  async executeCommand(command) {
    if (this.isExecuting) {
      console.warn('Command execution in progress, please wait');
      return { success: false, error: 'Another command is being executed' };
    }

    this.isExecuting = true;

    try {
      // Store the current state before executing
      const beforeState = await this.captureState(command.affectedRanges);
      
      // Execute the command
      const result = await this.runCommand(command);
      
      if (result.success) {
        // Store the after state
        const afterState = await this.captureState(command.affectedRanges);
        
        // Add to history
        this.addToHistory({
          id: Date.now().toString(),
          command: command,
          beforeState: beforeState,
          afterState: afterState,
          timestamp: new Date().toISOString(),
          description: command.description || 'Excel modification'
        });
      }
      
      return result;
      
    } catch (error) {
      console.error('Command execution error:', error);
      return { success: false, error: error.message };
    } finally {
      this.isExecuting = false;
    }
  }

  async runCommand(command) {
    if (typeof Excel === 'undefined') {
      return { success: false, error: 'Excel not available' };
    }

    try {
      return await Excel.run(async (context) => {
        let result = { success: true };
        
        switch (command.type) {
          case 'setValue':
            result = await this.setValue(context, command.params);
            break;
            
          case 'setFormula':
            result = await this.setFormula(context, command.params);
            break;
            
          case 'setFormat':
            result = await this.setFormat(context, command.params);
            break;
            
          case 'insertRows':
            result = await this.insertRows(context, command.params);
            break;
            
          case 'insertColumns':
            result = await this.insertColumns(context, command.params);
            break;
            
          case 'deleteRows':
            result = await this.deleteRows(context, command.params);
            break;
            
          case 'deleteColumns':
            result = await this.deleteColumns(context, command.params);
            break;
            
          case 'createSheet':
            result = await this.createSheet(context, command.params);
            break;
            
          case 'createTable':
            result = await this.createTable(context, command.params);
            break;
            
          case 'createChart':
            result = await this.createChart(context, command.params);
            break;
            
          case 'batchUpdate':
            result = await this.batchUpdate(context, command.params);
            break;
            
          default:
            result = { success: false, error: `Unknown command type: ${command.type}` };
        }
        
        await context.sync();
        return result;
      });
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  async setValue(context, params) {
    const { worksheet, range, values } = params;
    const ws = worksheet ? 
      context.workbook.worksheets.getItem(worksheet) : 
      context.workbook.worksheets.getActiveWorksheet();
    
    const targetRange = ws.getRange(range);
    targetRange.values = values;
    
    return { success: true };
  }

  async setFormula(context, params) {
    const { worksheet, range, formulas } = params;
    const ws = worksheet ? 
      context.workbook.worksheets.getItem(worksheet) : 
      context.workbook.worksheets.getActiveWorksheet();
    
    const targetRange = ws.getRange(range);
    targetRange.formulas = formulas;
    
    return { success: true };
  }

  async setFormat(context, params) {
    const { worksheet, range, format } = params;
    const ws = worksheet ? 
      context.workbook.worksheets.getItem(worksheet) : 
      context.workbook.worksheets.getActiveWorksheet();
    
    const targetRange = ws.getRange(range);
    
    if (format.numberFormat) {
      targetRange.numberFormat = format.numberFormat;
    }
    
    if (format.font) {
      Object.assign(targetRange.format.font, format.font);
    }
    
    if (format.fill) {
      targetRange.format.fill.color = format.fill.color;
    }
    
    if (format.borders) {
      // Apply borders
      ['top', 'bottom', 'left', 'right'].forEach(side => {
        if (format.borders[side]) {
          targetRange.format.borders.getItem(side).style = format.borders[side].style;
          targetRange.format.borders.getItem(side).color = format.borders[side].color;
        }
      });
    }
    
    return { success: true };
  }

  async insertRows(context, params) {
    const { worksheet, startRow, count } = params;
    const ws = worksheet ? 
      context.workbook.worksheets.getItem(worksheet) : 
      context.workbook.worksheets.getActiveWorksheet();
    
    const range = ws.getRangeByIndexes(startRow, 0, count, 1);
    range.insert(Excel.InsertShiftDirection.down);
    
    return { success: true };
  }

  async insertColumns(context, params) {
    const { worksheet, startColumn, count } = params;
    const ws = worksheet ? 
      context.workbook.worksheets.getItem(worksheet) : 
      context.workbook.worksheets.getActiveWorksheet();
    
    const range = ws.getRangeByIndexes(0, startColumn, 1, count);
    range.insert(Excel.InsertShiftDirection.right);
    
    return { success: true };
  }

  async deleteRows(context, params) {
    const { worksheet, startRow, count } = params;
    const ws = worksheet ? 
      context.workbook.worksheets.getItem(worksheet) : 
      context.workbook.worksheets.getActiveWorksheet();
    
    const range = ws.getRangeByIndexes(startRow, 0, count, 1);
    range.delete(Excel.DeleteShiftDirection.up);
    
    return { success: true };
  }

  async deleteColumns(context, params) {
    const { worksheet, startColumn, count } = params;
    const ws = worksheet ? 
      context.workbook.worksheets.getItem(worksheet) : 
      context.workbook.worksheets.getActiveWorksheet();
    
    const range = ws.getRangeByIndexes(0, startColumn, 1, count);
    range.delete(Excel.DeleteShiftDirection.left);
    
    return { success: true };
  }

  async createSheet(context, params) {
    const { name, position } = params;
    const newSheet = context.workbook.worksheets.add(name);
    
    if (position !== undefined) {
      newSheet.position = position;
    }
    
    newSheet.activate();
    
    return { success: true, sheetName: name };
  }

  async createTable(context, params) {
    const { worksheet, range, name, hasHeaders, style } = params;
    const ws = worksheet ? 
      context.workbook.worksheets.getItem(worksheet) : 
      context.workbook.worksheets.getActiveWorksheet();
    
    const targetRange = ws.getRange(range);
    const table = ws.tables.add(targetRange, hasHeaders !== false);
    
    if (name) table.name = name;
    if (style) table.style = style;
    
    return { success: true, tableName: table.name };
  }

  async createChart(context, params) {
    const { worksheet, sourceData, chartType, name, position } = params;
    const ws = worksheet ? 
      context.workbook.worksheets.getItem(worksheet) : 
      context.workbook.worksheets.getActiveWorksheet();
    
    const dataRange = ws.getRange(sourceData);
    const chart = ws.charts.add(
      chartType || Excel.ChartType.columnClustered,
      dataRange,
      Excel.ChartSeriesBy.auto
    );
    
    if (name) chart.name = name;
    if (position) {
      chart.top = position.top || 0;
      chart.left = position.left || 0;
      chart.height = position.height || 300;
      chart.width = position.width || 400;
    }
    
    return { success: true, chartName: chart.name };
  }

  async batchUpdate(context, params) {
    const { updates } = params;
    const results = [];
    
    for (const update of updates) {
      const result = await this.runCommand(update);
      results.push(result);
      
      if (!result.success) {
        return { 
          success: false, 
          error: `Batch update failed at step ${results.length}: ${result.error}`,
          completedSteps: results.length - 1
        };
      }
    }
    
    return { success: true, results: results };
  }

  async captureState(ranges) {
    if (typeof Excel === 'undefined' || !ranges || ranges.length === 0) {
      return null;
    }

    try {
      return await Excel.run(async (context) => {
        const state = {};
        
        for (const rangeInfo of ranges) {
          const ws = rangeInfo.worksheet ? 
            context.workbook.worksheets.getItem(rangeInfo.worksheet) : 
            context.workbook.worksheets.getActiveWorksheet();
          
          const range = ws.getRange(rangeInfo.address);
          range.load(['values', 'formulas', 'numberFormat', 'format/*']);
          
          await context.sync();
          
          const key = `${rangeInfo.worksheet || 'active'}!${rangeInfo.address}`;
          state[key] = {
            values: range.values,
            formulas: range.formulas,
            numberFormat: range.numberFormat,
            format: {
              font: {
                bold: range.format.font.bold,
                color: range.format.font.color,
                italic: range.format.font.italic,
                size: range.format.font.size,
                name: range.format.font.name
              },
              fill: {
                color: range.format.fill.color
              }
            }
          };
        }
        
        return state;
      });
    } catch (error) {
      console.error('Error capturing state:', error);
      return null;
    }
  }

  async undo() {
    if (this.currentIndex < 0) {
      console.log('Nothing to undo');
      return { success: false, error: 'No actions to undo' };
    }

    const historyItem = this.commandHistory[this.currentIndex];
    
    try {
      // Restore the before state
      await this.restoreState(historyItem.beforeState);
      this.currentIndex--;
      
      return { 
        success: true, 
        description: `Undone: ${historyItem.description}` 
      };
    } catch (error) {
      console.error('Undo failed:', error);
      return { success: false, error: error.message };
    }
  }

  async redo() {
    if (this.currentIndex >= this.commandHistory.length - 1) {
      console.log('Nothing to redo');
      return { success: false, error: 'No actions to redo' };
    }

    this.currentIndex++;
    const historyItem = this.commandHistory[this.currentIndex];
    
    try {
      // Restore the after state
      await this.restoreState(historyItem.afterState);
      
      return { 
        success: true, 
        description: `Redone: ${historyItem.description}` 
      };
    } catch (error) {
      console.error('Redo failed:', error);
      this.currentIndex--;
      return { success: false, error: error.message };
    }
  }

  async restoreState(state) {
    if (!state || typeof Excel === 'undefined') {
      return;
    }

    await Excel.run(async (context) => {
      for (const [key, data] of Object.entries(state)) {
        const [worksheet, address] = key.split('!');
        const ws = worksheet === 'active' ? 
          context.workbook.worksheets.getActiveWorksheet() : 
          context.workbook.worksheets.getItem(worksheet);
        
        const range = ws.getRange(address);
        
        if (data.values) range.values = data.values;
        if (data.formulas) range.formulas = data.formulas;
        if (data.numberFormat) range.numberFormat = data.numberFormat;
        
        if (data.format) {
          if (data.format.font) {
            Object.assign(range.format.font, data.format.font);
          }
          if (data.format.fill && data.format.fill.color) {
            range.format.fill.color = data.format.fill.color;
          }
        }
      }
      
      await context.sync();
    });
  }

  addToHistory(item) {
    // Remove any items after current index (for branching history)
    this.commandHistory = this.commandHistory.slice(0, this.currentIndex + 1);
    
    // Add new item
    this.commandHistory.push(item);
    this.currentIndex++;
    
    // Limit history size
    if (this.commandHistory.length > this.maxHistorySize) {
      this.commandHistory.shift();
      this.currentIndex--;
    }
  }

  getHistory() {
    return this.commandHistory.map((item, index) => ({
      id: item.id,
      description: item.description,
      timestamp: item.timestamp,
      isCurrent: index === this.currentIndex,
      canUndo: index <= this.currentIndex,
      canRedo: index > this.currentIndex
    }));
  }

  clearHistory() {
    this.commandHistory = [];
    this.currentIndex = -1;
  }
}

window.ExcelCommandExecutor = ExcelCommandExecutor;