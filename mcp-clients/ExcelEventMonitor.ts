/**
 * ExcelEventMonitor.ts
 * Monitors Excel events and provides real-time notifications to MCP system
 */

export interface ExcelEvent {
  id: string;
  type: string;
  timestamp: Date;
  source: string;
  data: any;
  affectedCells?: string[];
  worksheet?: string;
  user?: string;
}

export interface EventSubscription {
  id: string;
  eventTypes: string[];
  callback: (event: ExcelEvent) => Promise<void>;
  filter?: (event: ExcelEvent) => boolean;
  priority: 'low' | 'normal' | 'high';
}

export interface EventMonitorConfig {
  enabledEvents: string[];
  debounceMs: number;
  maxEventsPerSecond: number;
  persistEvents: boolean;
  logLevel: 'none' | 'error' | 'warn' | 'info' | 'debug';
}

export class ExcelEventMonitor {
  private config: EventMonitorConfig;
  private subscriptions: Map<string, EventSubscription> = new Map();
  private eventHistory: ExcelEvent[] = [];
  private maxHistorySize: number = 1000;
  private isInitialized: boolean = false;
  private eventQueue: ExcelEvent[] = [];
  private isProcessingQueue: boolean = false;
  private debounceTimers: Map<string, NodeJS.Timeout> = new Map();
  private rateLimitCounter: number = 0;
  private rateLimitWindow: number = 1000; // 1 second

  constructor(config?: Partial<EventMonitorConfig>) {
    this.config = {
      enabledEvents: [
        'worksheet-changed',
        'selection-changed',
        'format-changed',
        'formula-changed',
        'worksheet-added',
        'worksheet-deleted',
        'workbook-opened',
        'calculation-completed'
      ],
      debounceMs: 100,
      maxEventsPerSecond: 50,
      persistEvents: true,
      logLevel: 'info',
      ...config
    };

    this.setupRateLimitReset();
  }

  /**
   * Initialize Excel event monitoring
   */
  async initialize(): Promise<void> {
    if (this.isInitialized) return;

    console.log('🎯 Initializing Excel Event Monitor...');

    if (typeof Excel === 'undefined') {
      console.warn('⚠️ Excel API not available, event monitoring disabled');
      return;
    }

    try {
      await this.setupExcelEventListeners();
      this.isInitialized = true;
      console.log('✅ Excel Event Monitor initialized');
    } catch (error) {
      console.error('❌ Failed to initialize Excel Event Monitor:', error);
      throw error;
    }
  }

  /**
   * Set up Excel event listeners
   */
  private async setupExcelEventListeners(): Promise<void> {
    await Excel.run(async (context) => {
      const workbook = context.workbook;

      // Selection change events
      if (this.isEventEnabled('selection-changed')) {
        workbook.onSelectionChanged.add(async (event) => {
          await this.emitEvent({
            type: 'selection-changed',
            source: 'excel-api',
            data: {
              address: event.address,
              worksheet: event.worksheetId
            },
            affectedCells: [event.address],
            worksheet: event.worksheetId
          });
        });
      }

      // Worksheet collection events
      if (this.isEventEnabled('worksheet-added') || this.isEventEnabled('worksheet-deleted')) {
        workbook.worksheets.onAdded.add(async (event) => {
          await this.emitEvent({
            type: 'worksheet-added',
            source: 'excel-api',
            data: {
              worksheetId: event.worksheetId
            },
            worksheet: event.worksheetId
          });
        });

        workbook.worksheets.onDeleted.add(async (event) => {
          await this.emitEvent({
            type: 'worksheet-deleted',
            source: 'excel-api',
            data: {
              worksheetId: event.worksheetId
            },
            worksheet: event.worksheetId
          });
        });
      }

      // Worksheet data change events
      if (this.isEventEnabled('worksheet-changed')) {
        workbook.worksheets.onChanged.add(async (event) => {
          await this.emitEvent({
            type: 'worksheet-changed',
            source: 'excel-api',
            data: {
              address: event.address,
              changeType: event.changeType,
              details: event.details,
              worksheet: event.worksheetId
            },
            affectedCells: event.address ? [event.address] : undefined,
            worksheet: event.worksheetId
          });
        });
      }

      // Formula change events
      if (this.isEventEnabled('formula-changed')) {
        workbook.worksheets.onFormulaChanged.add(async (event) => {
          await this.emitEvent({
            type: 'formula-changed',
            source: 'excel-api',
            data: {
              address: event.address,
              formulaDetails: event.formulaDetails,
              worksheet: event.worksheetId
            },
            affectedCells: event.address ? [event.address] : undefined,
            worksheet: event.worksheetId
          });
        });
      }

      // Format change events (if supported)
      if (this.isEventEnabled('format-changed')) {
        try {
          workbook.worksheets.onFormatChanged.add(async (event: any) => {
            await this.emitEvent({
              type: 'format-changed',
              source: 'excel-api',
              data: {
                address: event.address,
                formatType: event.formatType,
                worksheet: event.worksheetId
              },
              affectedCells: event.address ? [event.address] : undefined,
              worksheet: event.worksheetId
            });
          });
        } catch (error) {
          this.log('warn', 'Format change events not supported in this Excel version');
        }
      }

      // Calculation events
      if (this.isEventEnabled('calculation-completed')) {
        workbook.onActivated.add(async (event) => {
          await this.emitEvent({
            type: 'workbook-activated',
            source: 'excel-api',
            data: {
              workbookId: event.source
            }
          });
        });
      }

      await context.sync();
      this.log('info', 'Excel event listeners registered');
    });
  }

  /**
   * Subscribe to events
   */
  subscribe(
    eventTypes: string | string[],
    callback: (event: ExcelEvent) => Promise<void>,
    options?: {
      filter?: (event: ExcelEvent) => boolean;
      priority?: 'low' | 'normal' | 'high';
    }
  ): string {
    const subscriptionId = this.generateSubscriptionId();
    const types = Array.isArray(eventTypes) ? eventTypes : [eventTypes];
    
    const subscription: EventSubscription = {
      id: subscriptionId,
      eventTypes: types,
      callback,
      filter: options?.filter,
      priority: options?.priority || 'normal'
    };

    this.subscriptions.set(subscriptionId, subscription);
    
    this.log('info', `Subscribed to events: ${types.join(', ')}`);
    return subscriptionId;
  }

  /**
   * Unsubscribe from events
   */
  unsubscribe(subscriptionId: string): boolean {
    const success = this.subscriptions.delete(subscriptionId);
    if (success) {
      this.log('info', `Unsubscribed: ${subscriptionId}`);
    }
    return success;
  }

  /**
   * Emit an event
   */
  private async emitEvent(eventData: Omit<ExcelEvent, 'id' | 'timestamp'>): Promise<void> {
    const event: ExcelEvent = {
      ...eventData,
      id: this.generateEventId(),
      timestamp: new Date()
    };

    // Check rate limiting
    if (!this.isWithinRateLimit()) {
      this.log('warn', 'Event rate limit exceeded, dropping event');
      return;
    }

    // Add to history if persistence is enabled
    if (this.config.persistEvents) {
      this.addToHistory(event);
    }

    // Add to processing queue
    this.eventQueue.push(event);
    this.processEventQueue();

    this.log('debug', `Emitted event: ${event.type}`);
  }

  /**
   * Process event queue
   */
  private async processEventQueue(): Promise<void> {
    if (this.isProcessingQueue || this.eventQueue.length === 0) {
      return;
    }

    this.isProcessingQueue = true;

    try {
      while (this.eventQueue.length > 0) {
        const event = this.eventQueue.shift()!;
        await this.processEvent(event);
      }
    } finally {
      this.isProcessingQueue = false;
    }
  }

  /**
   * Process individual event
   */
  private async processEvent(event: ExcelEvent): Promise<void> {
    // Get subscriptions that match this event type
    const matchingSubscriptions = Array.from(this.subscriptions.values())
      .filter(sub => sub.eventTypes.includes(event.type))
      .filter(sub => !sub.filter || sub.filter(event));

    // Sort by priority
    matchingSubscriptions.sort((a, b) => {
      const priorityOrder = { 'high': 3, 'normal': 2, 'low': 1 };
      return priorityOrder[b.priority] - priorityOrder[a.priority];
    });

    // Execute callbacks
    for (const subscription of matchingSubscriptions) {
      try {
        const debounceKey = `${subscription.id}-${event.type}`;
        
        // Clear existing debounce timer
        const existingTimer = this.debounceTimers.get(debounceKey);
        if (existingTimer) {
          clearTimeout(existingTimer);
        }

        // Set new debounce timer
        const timer = setTimeout(async () => {
          try {
            await subscription.callback(event);
          } catch (error) {
            this.log('error', `Subscription callback failed: ${error}`);
          } finally {
            this.debounceTimers.delete(debounceKey);
          }
        }, this.config.debounceMs);

        this.debounceTimers.set(debounceKey, timer);
        
      } catch (error) {
        this.log('error', `Failed to process event subscription: ${error}`);
      }
    }
  }

  /**
   * Get event history
   */
  getEventHistory(
    eventType?: string,
    limit?: number,
    since?: Date
  ): ExcelEvent[] {
    let events = [...this.eventHistory];

    if (eventType) {
      events = events.filter(e => e.type === eventType);
    }

    if (since) {
      events = events.filter(e => e.timestamp >= since);
    }

    if (limit) {
      events = events.slice(-limit);
    }

    return events.reverse(); // Most recent first
  }

  /**
   * Get event statistics
   */
  getStatistics(): {
    totalEvents: number;
    eventsByType: { [type: string]: number };
    eventsInLastHour: number;
    averageEventsPerMinute: number;
    activeSubscriptions: number;
  } {
    const oneHourAgo = new Date(Date.now() - 60 * 60 * 1000);
    const eventsInLastHour = this.eventHistory.filter(
      e => e.timestamp >= oneHourAgo
    ).length;

    const eventsByType: { [type: string]: number } = {};
    this.eventHistory.forEach(event => {
      eventsByType[event.type] = (eventsByType[event.type] || 0) + 1;
    });

    const oldestEvent = this.eventHistory[0];
    const averageEventsPerMinute = oldestEvent 
      ? this.eventHistory.length / ((Date.now() - oldestEvent.timestamp.getTime()) / (1000 * 60))
      : 0;

    return {
      totalEvents: this.eventHistory.length,
      eventsByType,
      eventsInLastHour,
      averageEventsPerMinute,
      activeSubscriptions: this.subscriptions.size
    };
  }

  /**
   * Monitor specific cell or range
   */
  monitorRange(
    range: string,
    worksheet?: string,
    callback?: (event: ExcelEvent) => Promise<void>
  ): string {
    return this.subscribe(
      ['worksheet-changed', 'formula-changed', 'format-changed'],
      async (event) => {
        // Check if event affects the monitored range
        if (this.doesEventAffectRange(event, range, worksheet)) {
          if (callback) {
            await callback(event);
          }
          this.log('info', `Monitored range ${range} affected by ${event.type}`);
        }
      },
      {
        filter: (event) => this.doesEventAffectRange(event, range, worksheet),
        priority: 'high'
      }
    );
  }

  /**
   * Stop monitoring and cleanup
   */
  async cleanup(): Promise<void> {
    // Clear all debounce timers
    for (const timer of this.debounceTimers.values()) {
      clearTimeout(timer);
    }
    this.debounceTimers.clear();

    // Clear subscriptions
    this.subscriptions.clear();

    // Clear event queue
    this.eventQueue = [];

    this.isInitialized = false;
    this.log('info', 'Excel Event Monitor cleaned up');
  }

  // Private helper methods

  private isEventEnabled(eventType: string): boolean {
    return this.config.enabledEvents.includes(eventType);
  }

  private isWithinRateLimit(): boolean {
    this.rateLimitCounter++;
    return this.rateLimitCounter <= this.config.maxEventsPerSecond;
  }

  private setupRateLimitReset(): void {
    setInterval(() => {
      this.rateLimitCounter = 0;
    }, this.rateLimitWindow);
  }

  private addToHistory(event: ExcelEvent): void {
    this.eventHistory.push(event);
    
    // Maintain history size limit
    if (this.eventHistory.length > this.maxHistorySize) {
      this.eventHistory.splice(0, this.eventHistory.length - this.maxHistorySize);
    }
  }

  private doesEventAffectRange(event: ExcelEvent, range: string, worksheet?: string): boolean {
    // Check worksheet match
    if (worksheet && event.worksheet && event.worksheet !== worksheet) {
      return false;
    }

    // Check if event affects the range
    if (event.affectedCells) {
      return event.affectedCells.some(cell => this.isRangeOverlap(cell, range));
    }

    return false;
  }

  private isRangeOverlap(range1: string, range2: string): boolean {
    // Simple range overlap check - in production, use proper range comparison
    return range1.includes(range2) || range2.includes(range1);
  }

  private generateEventId(): string {
    return `event_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  private generateSubscriptionId(): string {
    return `sub_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  private log(level: string, message: string): void {
    const levels = ['none', 'error', 'warn', 'info', 'debug'];
    const configLevel = levels.indexOf(this.config.logLevel);
    const messageLevel = levels.indexOf(level);

    if (messageLevel <= configLevel) {
      console.log(`[ExcelEventMonitor] ${message}`);
    }
  }

  /**
   * Export event monitor state
   */
  export(): string {
    return JSON.stringify({
      config: this.config,
      eventHistory: this.eventHistory,
      statistics: this.getStatistics(),
      exportTime: new Date().toISOString()
    }, null, 2);
  }

  /**
   * Get real-time event stream (for debugging)
   */
  getEventStream(): {
    subscribe: (callback: (event: ExcelEvent) => void) => string;
    unsubscribe: (id: string) => void;
  } {
    const streamSubscriptions = new Map<string, (event: ExcelEvent) => void>();

    return {
      subscribe: (callback: (event: ExcelEvent) => void) => {
        const id = this.generateSubscriptionId();
        streamSubscriptions.set(id, callback);

        // Subscribe to all events
        this.subscribe('*', async (event) => {
          const streamCallback = streamSubscriptions.get(id);
          if (streamCallback) {
            streamCallback(event);
          }
        });

        return id;
      },
      unsubscribe: (id: string) => {
        streamSubscriptions.delete(id);
        this.unsubscribe(id);
      }
    };
  }
}

export default ExcelEventMonitor;