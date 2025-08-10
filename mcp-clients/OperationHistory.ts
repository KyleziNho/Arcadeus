/**
 * OperationHistory.ts
 * Manages operation history for undo/redo functionality in Excel operations
 */

export interface Operation {
  id: string;
  type: string;
  timestamp: Date;
  description: string;
  params: Record<string, any>;
  status: 'pending' | 'completed' | 'failed' | 'undone' | 'redone';
  error?: string;
  inverse?: {
    type: string;
    params: Record<string, any>;
  };
  beforeState?: any;
  afterState?: any;
  affectedCells?: string[];
  userId?: string;
}

export interface OperationBatch {
  id: string;
  timestamp: Date;
  description: string;
  operations: Operation[];
  status: 'completed' | 'undone';
}

export interface RecentChange {
  operation: string;
  timestamp: Date;
  messageId?: string;
  description?: string;
}

export class OperationHistory {
  private operations: Operation[] = [];
  private batches: OperationBatch[] = [];
  private undoneOperations: Operation[] = [];
  private maxHistorySize: number = 500;
  private maxBatchSize: number = 100;
  private currentBatch: Operation[] | null = null;
  private batchDescription: string = '';

  /**
   * Record a new operation
   */
  recordOperation(operation: Omit<Operation, 'id' | 'timestamp'>): Operation {
    const fullOperation: Operation = {
      ...operation,
      id: this.generateOperationId(),
      timestamp: new Date()
    };

    // Add to current batch if one is active
    if (this.currentBatch) {
      this.currentBatch.push(fullOperation);
    } else {
      this.operations.push(fullOperation);
    }

    // Clear redo history when new operation is recorded
    this.undoneOperations = [];

    // Maintain history size limit
    this.limitHistorySize();

    console.log('📝 Recorded operation:', fullOperation);
    return fullOperation;
  }

  /**
   * Start a batch of operations
   */
  startBatch(description: string): void {
    if (this.currentBatch) {
      console.warn('⚠️ Starting new batch while another is active');
      this.endBatch();
    }

    this.currentBatch = [];
    this.batchDescription = description;
    console.log(`📦 Started batch: ${description}`);
  }

  /**
   * End the current batch
   */
  endBatch(): string | null {
    if (!this.currentBatch) {
      console.warn('⚠️ No active batch to end');
      return null;
    }

    if (this.currentBatch.length === 0) {
      console.warn('⚠️ Ending empty batch');
      this.currentBatch = null;
      return null;
    }

    const batch: OperationBatch = {
      id: this.generateBatchId(),
      timestamp: new Date(),
      description: this.batchDescription,
      operations: [...this.currentBatch],
      status: 'completed'
    };

    this.batches.push(batch);
    this.operations.push(...this.currentBatch);

    console.log(`✅ Ended batch: ${batch.description} (${batch.operations.length} operations)`);
    
    this.currentBatch = null;
    this.batchDescription = '';

    return batch.id;
  }

  /**
   * Get the last operation that can be undone
   */
  getLastOperation(): Operation | null {
    // Find the last completed operation that hasn't been undone
    for (let i = this.operations.length - 1; i >= 0; i--) {
      const op = this.operations[i];
      if (op.status === 'completed' && !this.isOperationUndone(op.id)) {
        return op;
      }
    }
    return null;
  }

  /**
   * Get the last undone operation that can be redone
   */
  getLastUndoneOperation(): Operation | null {
    if (this.undoneOperations.length === 0) return null;
    return this.undoneOperations[this.undoneOperations.length - 1];
  }

  /**
   * Mark an operation as undone
   */
  markAsUndone(operationId: string): boolean {
    const operation = this.findOperation(operationId);
    if (!operation) {
      console.error(`❌ Operation not found: ${operationId}`);
      return false;
    }

    operation.status = 'undone';
    this.undoneOperations.push(operation);
    
    console.log(`↩️ Marked operation as undone: ${operation.description}`);
    return true;
  }

  /**
   * Mark an operation as redone
   */
  markAsRedone(operationId: string): boolean {
    const operationIndex = this.undoneOperations.findIndex(op => op.id === operationId);
    if (operationIndex === -1) {
      console.error(`❌ Undone operation not found: ${operationId}`);
      return false;
    }

    const operation = this.undoneOperations[operationIndex];
    operation.status = 'redone';
    
    // Remove from undone operations
    this.undoneOperations.splice(operationIndex, 1);
    
    console.log(`↪️ Marked operation as redone: ${operation.description}`);
    return true;
  }

  /**
   * Undo a batch of operations
   */
  async undoBatch(batchId: string): Promise<boolean> {
    const batch = this.batches.find(b => b.id === batchId);
    if (!batch || batch.status === 'undone') {
      return false;
    }

    // Undo operations in reverse order
    for (let i = batch.operations.length - 1; i >= 0; i--) {
      const operation = batch.operations[i];
      if (operation.inverse) {
        try {
          // Execute inverse operation
          await this.executeInverse(operation);
          this.markAsUndone(operation.id);
        } catch (error) {
          console.error(`Failed to undo operation ${operation.id}:`, error);
          return false;
        }
      }
    }

    batch.status = 'undone';
    console.log(`↩️ Undid batch: ${batch.description}`);
    return true;
  }

  /**
   * Get recent changes for context
   */
  getRecentChanges(limit: number = 10): RecentChange[] {
    return this.operations
      .filter(op => op.status === 'completed')
      .slice(-limit)
      .map(op => ({
        operation: op.type,
        timestamp: op.timestamp,
        description: op.description
      }));
  }

  /**
   * Get all operations
   */
  getAllOperations(): Operation[] {
    return [...this.operations];
  }

  /**
   * Get operations by type
   */
  getOperationsByType(type: string): Operation[] {
    return this.operations.filter(op => op.type === type);
  }

  /**
   * Get operations in date range
   */
  getOperationsInRange(startDate: Date, endDate: Date): Operation[] {
    return this.operations.filter(op => 
      op.timestamp >= startDate && op.timestamp <= endDate
    );
  }

  /**
   * Search operations by description or params
   */
  searchOperations(query: string): Operation[] {
    const queryLower = query.toLowerCase();
    return this.operations.filter(op =>
      op.description.toLowerCase().includes(queryLower) ||
      JSON.stringify(op.params).toLowerCase().includes(queryLower)
    );
  }

  /**
   * Get operation statistics
   */
  getStatistics(): {
    totalOperations: number;
    completedOperations: number;
    failedOperations: number;
    undoneOperations: number;
    averageOperationTime?: number;
    mostCommonType: string;
    operationsByHour: { [hour: number]: number };
  } {
    const completed = this.operations.filter(op => op.status === 'completed');
    const failed = this.operations.filter(op => op.status === 'failed');
    const undone = this.undoneOperations;

    // Count operations by type
    const typeCount: { [type: string]: number } = {};
    this.operations.forEach(op => {
      typeCount[op.type] = (typeCount[op.type] || 0) + 1;
    });
    
    const mostCommonType = Object.entries(typeCount)
      .sort(([,a], [,b]) => b - a)[0]?.[0] || 'none';

    // Group by hour
    const operationsByHour: { [hour: number]: number } = {};
    this.operations.forEach(op => {
      const hour = op.timestamp.getHours();
      operationsByHour[hour] = (operationsByHour[hour] || 0) + 1;
    });

    return {
      totalOperations: this.operations.length,
      completedOperations: completed.length,
      failedOperations: failed.length,
      undoneOperations: undone.length,
      mostCommonType,
      operationsByHour
    };
  }

  /**
   * Create inverse operation for undo functionality
   */
  createInverseOperation(operation: Operation): Operation['inverse'] | null {
    switch (operation.type) {
      case 'write-value':
        return {
          type: 'write-value',
          params: {
            ...operation.params,
            data: operation.beforeState?.value || '',
            originalData: operation.params.data // For redo
          }
        };

      case 'format-cells':
        return {
          type: 'format-cells',
          params: {
            ...operation.params,
            format: operation.beforeState?.format || {},
            originalFormat: operation.params.format
          }
        };

      case 'delete-range':
        return {
          type: 'write-value',
          params: {
            range: operation.params.range,
            worksheet: operation.params.worksheet,
            data: operation.beforeState?.values || ''
          }
        };

      case 'insert-cells':
        return {
          type: 'delete-range',
          params: {
            range: operation.params.range,
            worksheet: operation.params.worksheet
          }
        };

      default:
        console.warn(`⚠️ No inverse operation defined for type: ${operation.type}`);
        return null;
    }
  }

  /**
   * Capture before state for operation
   */
  async captureBeforeState(operation: Operation, excelAPI?: any): Promise<void> {
    if (!excelAPI || typeof Excel === 'undefined') return;

    try {
      await Excel.run(async (context) => {
        const worksheet = operation.params.worksheet 
          ? context.workbook.worksheets.getItem(operation.params.worksheet)
          : context.workbook.worksheets.getActiveWorksheet();
        
        const range = worksheet.getRange(operation.params.range);
        range.load(['values', 'formulas', 'format']);
        
        await context.sync();
        
        operation.beforeState = {
          values: range.values,
          formulas: range.formulas,
          format: range.format
        };
      });
    } catch (error) {
      console.warn('Failed to capture before state:', error);
    }
  }

  /**
   * Clear operation history
   */
  clear(): void {
    this.operations = [];
    this.batches = [];
    this.undoneOperations = [];
    this.currentBatch = null;
    this.batchDescription = '';
    console.log('🗑️ Cleared operation history');
  }

  /**
   * Export operation history
   */
  export(): string {
    return JSON.stringify({
      operations: this.operations,
      batches: this.batches,
      undoneOperations: this.undoneOperations,
      exportTime: new Date().toISOString(),
      version: '1.0.0'
    }, null, 2);
  }

  /**
   * Import operation history
   */
  import(historyData: string): void {
    try {
      const data = JSON.parse(historyData);
      
      this.operations = data.operations || [];
      this.batches = data.batches || [];
      this.undoneOperations = data.undoneOperations || [];
      
      // Convert timestamp strings back to Date objects
      this.operations = this.operations.map(op => ({
        ...op,
        timestamp: new Date(op.timestamp)
      }));
      
      this.batches = this.batches.map(batch => ({
        ...batch,
        timestamp: new Date(batch.timestamp),
        operations: batch.operations.map(op => ({
          ...op,
          timestamp: new Date(op.timestamp)
        }))
      }));
      
      console.log('📥 Imported operation history');
    } catch (error) {
      console.error('❌ Failed to import history:', error);
      throw new Error('Invalid history data format');
    }
  }

  /**
   * Check if operation can be undone
   */
  canUndo(operationId?: string): boolean {
    if (operationId) {
      const operation = this.findOperation(operationId);
      return operation?.status === 'completed' && !this.isOperationUndone(operationId);
    }
    
    return this.getLastOperation() !== null;
  }

  /**
   * Check if operation can be redone
   */
  canRedo(): boolean {
    return this.undoneOperations.length > 0;
  }

  // Private helper methods

  private generateOperationId(): string {
    return `op_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  private generateBatchId(): string {
    return `batch_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  private findOperation(operationId: string): Operation | null {
    return this.operations.find(op => op.id === operationId) || null;
  }

  private isOperationUndone(operationId: string): boolean {
    return this.undoneOperations.some(op => op.id === operationId);
  }

  private limitHistorySize(): void {
    if (this.operations.length > this.maxHistorySize) {
      const toRemove = this.operations.length - this.maxHistorySize;
      this.operations.splice(0, toRemove);
    }

    if (this.batches.length > this.maxBatchSize) {
      const toRemove = this.batches.length - this.maxBatchSize;
      this.batches.splice(0, toRemove);
    }
  }

  private async executeInverse(operation: Operation): Promise<void> {
    // This would be implemented to actually execute the inverse operation
    // For now, it's a placeholder that would integrate with the Excel MCP server
    console.log(`🔄 Executing inverse of operation: ${operation.description}`);
  }

  /**
   * Get undo/redo suggestions based on recent operations
   */
  getSuggestions(): {
    canUndo: boolean;
    undoDescription?: string;
    canRedo: boolean;
    redoDescription?: string;
    batchAvailable: boolean;
  } {
    const lastOp = this.getLastOperation();
    const lastUndone = this.getLastUndoneOperation();
    
    return {
      canUndo: lastOp !== null,
      undoDescription: lastOp?.description,
      canRedo: lastUndone !== null,
      redoDescription: lastUndone?.description,
      batchAvailable: this.currentBatch !== null
    };
  }
}

export default OperationHistory;