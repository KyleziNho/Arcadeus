class ExcelContextReader {
  constructor() {
    this.lastContext = null;
    this.updateInterval = null;
    this.selectionChangeHandler = null;
  }

  async initialize() {
    if (typeof Excel === 'undefined') {
      console.warn('Excel API not available');
      return false;
    }

    try {
      await Excel.run(async (context) => {
        // Register event handlers for selection changes
        context.workbook.onSelectionChanged.add(this.handleSelectionChange.bind(this));
        await context.sync();
      });

      console.log('Excel Context Reader initialized');
      return true;
    } catch (error) {
      console.error('Failed to initialize Excel Context Reader:', error);
      return false;
    }
  }

  async handleSelectionChange(event) {
    console.log('Selection changed:', event.address);
    
    // Emit custom event that chat can listen to
    const customEvent = new CustomEvent('excelSelectionChanged', {
      detail: { address: event.address }
    });
    window.dispatchEvent(customEvent);
  }

  async getFullContext() {
    if (typeof Excel === 'undefined') {
      return { error: 'Excel not available' };
    }

    try {
      return await Excel.run(async (context) => {
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;
        const activeWorksheet = worksheets.getActiveWorksheet();
        const selectedRange = workbook.getSelectedRange();
        
        // Load all worksheets
        worksheets.load('items');
        await context.sync();
        
        const allSheets = [];
        
        // Get data from each worksheet
        for (let i = 0; i < worksheets.items.length; i++) {
          const worksheet = worksheets.items[i];
          worksheet.load(['name', 'visibility']);
          
          // Get used range
          const usedRange = worksheet.getUsedRangeOrNullObject();
          usedRange.load(['address', 'values', 'formulas', 'format', 'rowCount', 'columnCount']);
          
          await context.sync();
          
          if (!usedRange.isNullObject) {
            allSheets.push({
              name: worksheet.name,
              isActive: worksheet.name === activeWorksheet.name,
              visibility: worksheet.visibility,
              data: {
                address: usedRange.address,
                rowCount: usedRange.rowCount,
                columnCount: usedRange.columnCount,
                values: usedRange.values,
                formulas: usedRange.formulas
              }
            });
          }
        }
        
        // Get selected range details
        selectedRange.load(['address', 'values', 'formulas', 'format/*']);
        activeWorksheet.load('name');
        
        await context.sync();
        
        // Get named ranges
        const namedRanges = workbook.names;
        namedRanges.load('items');
        await context.sync();
        
        const namedRangesList = [];
        for (let i = 0; i < namedRanges.items.length; i++) {
          const namedRange = namedRanges.items[i];
          namedRange.load(['name', 'formula', 'comment']);
          await context.sync();
          
          namedRangesList.push({
            name: namedRange.name,
            formula: namedRange.formula,
            comment: namedRange.comment
          });
        }
        
        // Get workbook properties
        workbook.load(['name']);
        const properties = workbook.properties;
        properties.load(['title', 'subject', 'author', 'comments', 'company']);
        
        await context.sync();
        
        const fullContext = {
          workbook: {
            name: workbook.name,
            properties: {
              title: properties.title,
              subject: properties.subject,
              author: properties.author,
              comments: properties.comments,
              company: properties.company
            }
          },
          sheets: allSheets,
          selection: {
            worksheet: activeWorksheet.name,
            address: selectedRange.address,
            values: selectedRange.values,
            formulas: selectedRange.formulas,
            format: {
              fill: selectedRange.format.fill,
              font: selectedRange.format.font,
              borders: selectedRange.format.borders
            }
          },
          namedRanges: namedRangesList,
          timestamp: new Date().toISOString()
        };
        
        this.lastContext = fullContext;
        return fullContext;
      });
    } catch (error) {
      console.error('Error getting Excel context:', error);
      return {
        error: error.message,
        lastContext: this.lastContext
      };
    }
  }

  async getSelectedCellContext() {
    if (typeof Excel === 'undefined') {
      return { error: 'Excel not available' };
    }

    try {
      return await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        
        selectedRange.load(['address', 'values', 'formulas', 'format/*', 'rowIndex', 'columnIndex']);
        worksheet.load('name');
        
        // Get surrounding cells for context
        const expandedRange = selectedRange.getResizedRange(2, 2);
        expandedRange.load(['values', 'formulas']);
        
        await context.sync();
        
        return {
          cell: {
            address: selectedRange.address,
            value: selectedRange.values[0][0],
            formula: selectedRange.formulas[0][0],
            rowIndex: selectedRange.rowIndex,
            columnIndex: selectedRange.columnIndex,
            format: {
              numberFormat: selectedRange.format.numberFormat,
              font: selectedRange.format.font,
              fill: selectedRange.format.fill
            }
          },
          worksheet: worksheet.name,
          surroundingCells: {
            values: expandedRange.values,
            formulas: expandedRange.formulas
          }
        };
      });
    } catch (error) {
      console.error('Error getting selected cell context:', error);
      return { error: error.message };
    }
  }

  async getWorksheetData(worksheetName) {
    if (typeof Excel === 'undefined') {
      return { error: 'Excel not available' };
    }

    try {
      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem(worksheetName);
        const usedRange = worksheet.getUsedRange();
        
        usedRange.load(['address', 'values', 'formulas', 'format/*']);
        
        // Get charts
        const charts = worksheet.charts;
        charts.load('items');
        
        // Get tables
        const tables = worksheet.tables;
        tables.load('items');
        
        await context.sync();
        
        const chartsData = [];
        for (let i = 0; i < charts.items.length; i++) {
          const chart = charts.items[i];
          chart.load(['name', 'title', 'chartType', 'height', 'width']);
          await context.sync();
          
          chartsData.push({
            name: chart.name,
            title: chart.title,
            type: chart.chartType,
            dimensions: {
              height: chart.height,
              width: chart.width
            }
          });
        }
        
        const tablesData = [];
        for (let i = 0; i < tables.items.length; i++) {
          const table = tables.items[i];
          table.load(['name', 'style']);
          const tableRange = table.getRange();
          tableRange.load('address');
          await context.sync();
          
          tablesData.push({
            name: table.name,
            style: table.style,
            range: tableRange.address
          });
        }
        
        return {
          worksheet: worksheetName,
          data: {
            address: usedRange.address,
            values: usedRange.values,
            formulas: usedRange.formulas
          },
          charts: chartsData,
          tables: tablesData
        };
      });
    } catch (error) {
      console.error('Error getting worksheet data:', error);
      return { error: error.message };
    }
  }

  async watchForChanges(callback, interval = 1000) {
    if (this.updateInterval) {
      clearInterval(this.updateInterval);
    }

    this.updateInterval = setInterval(async () => {
      const context = await this.getFullContext();
      if (context && !context.error) {
        callback(context);
      }
    }, interval);
  }

  stopWatching() {
    if (this.updateInterval) {
      clearInterval(this.updateInterval);
      this.updateInterval = null;
    }
  }

  async getCellDependents(cellAddress) {
    if (typeof Excel === 'undefined') {
      return { error: 'Excel not available' };
    }

    try {
      return await Excel.run(async (context) => {
        const range = context.workbook.worksheets.getActiveWorksheet().getRange(cellAddress);
        const dependents = range.getDirectDependents();
        
        dependents.load('areas');
        await context.sync();
        
        const dependentCells = [];
        for (let i = 0; i < dependents.areas.items.length; i++) {
          const area = dependents.areas.items[i];
          area.load(['address', 'values', 'formulas']);
          await context.sync();
          
          dependentCells.push({
            address: area.address,
            values: area.values,
            formulas: area.formulas
          });
        }
        
        return {
          cell: cellAddress,
          dependents: dependentCells
        };
      });
    } catch (error) {
      console.error('Error getting cell dependents:', error);
      return { error: error.message };
    }
  }

  async getCellPrecedents(cellAddress) {
    if (typeof Excel === 'undefined') {
      return { error: 'Excel not available' };
    }

    try {
      return await Excel.run(async (context) => {
        const range = context.workbook.worksheets.getActiveWorksheet().getRange(cellAddress);
        const precedents = range.getDirectPrecedents();
        
        precedents.load('areas');
        await context.sync();
        
        const precedentCells = [];
        for (let i = 0; i < precedents.areas.items.length; i++) {
          const area = precedents.areas.items[i];
          area.load(['address', 'values', 'formulas']);
          await context.sync();
          
          precedentCells.push({
            address: area.address,
            values: area.values,
            formulas: area.formulas
          });
        }
        
        return {
          cell: cellAddress,
          precedents: precedentCells
        };
      });
    } catch (error) {
      console.error('Error getting cell precedents:', error);
      return { error: error.message };
    }
  }
}

window.ExcelContextReader = ExcelContextReader;