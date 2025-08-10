/**
 * ContextStore.ts
 * Manages persistent storage of conversation context, user preferences, and Excel state
 */

import {
  ConversationContext,
  ConversationMessage,
  ExcelState
} from '../mcp-types/interfaces';

export interface StoredContext {
  id: string;
  sessionId: string;
  timestamp: Date;
  context: ConversationContext;
  messages: ConversationMessage[];
  excelState: ExcelState;
  version: string;
}

export interface ContextStoreConfig {
  maxStoredSessions: number;
  maxMessagesPerSession: number;
  compressionEnabled: boolean;
  autoSave: boolean;
  encryptionEnabled: boolean;
}

export interface StorageBackend {
  get(key: string): Promise<string | null>;
  set(key: string, value: string): Promise<void>;
  delete(key: string): Promise<void>;
  list(prefix?: string): Promise<string[]>;
  clear(): Promise<void>;
}

export class ContextStore {
  private config: ContextStoreConfig;
  private storage: StorageBackend;
  private memoryCache: Map<string, StoredContext> = new Map();
  private saveQueue: Set<string> = new Set();
  private isProcessingQueue: boolean = false;

  constructor(config?: Partial<ContextStoreConfig>, customStorage?: StorageBackend) {
    this.config = {
      maxStoredSessions: 50,
      maxMessagesPerSession: 1000,
      compressionEnabled: true,
      autoSave: true,
      encryptionEnabled: false,
      ...config
    };

    this.storage = customStorage || this.createDefaultStorage();
    this.initializeStore();
  }

  /**
   * Initialize the context store
   */
  private async initializeStore(): Promise<void> {
    console.log('🏪 Initializing ContextStore...');
    
    try {
      // Load recent sessions into memory cache
      await this.loadRecentSessions();
      
      // Clean up old sessions
      await this.cleanupOldSessions();
      
      console.log('✅ ContextStore initialized');
    } catch (error) {
      console.error('❌ Failed to initialize ContextStore:', error);
    }
  }

  /**
   * Create default storage backend (localStorage)
   */
  private createDefaultStorage(): StorageBackend {
    return {
      async get(key: string): Promise<string | null> {
        return localStorage.getItem(key);
      },
      
      async set(key: string, value: string): Promise<void> {
        localStorage.setItem(key, value);
      },
      
      async delete(key: string): Promise<void> {
        localStorage.removeItem(key);
      },
      
      async list(prefix?: string): Promise<string[]> {
        const keys: string[] = [];
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (key && (!prefix || key.startsWith(prefix))) {
            keys.push(key);
          }
        }
        return keys;
      },
      
      async clear(): Promise<void> {
        localStorage.clear();
      }
    };
  }

  /**
   * Store conversation context
   */
  async storeContext(
    sessionId: string,
    context: ConversationContext,
    messages: ConversationMessage[],
    excelState: ExcelState
  ): Promise<void> {
    const contextId = `context_${sessionId}`;
    
    const storedContext: StoredContext = {
      id: contextId,
      sessionId,
      timestamp: new Date(),
      context,
      messages: this.limitMessages(messages),
      excelState,
      version: '1.0.0'
    };

    // Store in memory cache
    this.memoryCache.set(contextId, storedContext);

    // Queue for persistent storage if auto-save enabled
    if (this.config.autoSave) {
      this.saveQueue.add(contextId);
      this.processSaveQueue();
    }

    console.log(`💾 Stored context for session ${sessionId}`);
  }

  /**
   * Load conversation context
   */
  async loadContext(sessionId: string): Promise<StoredContext | null> {
    const contextId = `context_${sessionId}`;
    
    // Check memory cache first
    if (this.memoryCache.has(contextId)) {
      return this.memoryCache.get(contextId)!;
    }

    // Load from persistent storage
    try {
      const stored = await this.storage.get(contextId);
      if (stored) {
        const parsedContext = this.deserializeContext(stored);
        this.memoryCache.set(contextId, parsedContext);
        return parsedContext;
      }
    } catch (error) {
      console.error(`❌ Failed to load context for ${sessionId}:`, error);
    }

    return null;
  }

  /**
   * Update stored context
   */
  async updateContext(
    sessionId: string,
    updates: Partial<StoredContext>
  ): Promise<void> {
    const contextId = `context_${sessionId}`;
    const existing = await this.loadContext(sessionId);
    
    if (existing) {
      const updated: StoredContext = {
        ...existing,
        ...updates,
        timestamp: new Date()
      };
      
      this.memoryCache.set(contextId, updated);
      
      if (this.config.autoSave) {
        this.saveQueue.add(contextId);
        this.processSaveQueue();
      }
    }
  }

  /**
   * Delete stored context
   */
  async deleteContext(sessionId: string): Promise<void> {
    const contextId = `context_${sessionId}`;
    
    // Remove from memory cache
    this.memoryCache.delete(contextId);
    
    // Remove from persistent storage
    await this.storage.delete(contextId);
    
    console.log(`🗑️ Deleted context for session ${sessionId}`);
  }

  /**
   * List all stored sessions
   */
  async listSessions(): Promise<Array<{id: string; sessionId: string; timestamp: Date}>> {
    const sessions: Array<{id: string; sessionId: string; timestamp: Date}> = [];
    
    // Get from memory cache
    for (const [id, context] of this.memoryCache) {
      sessions.push({
        id,
        sessionId: context.sessionId,
        timestamp: context.timestamp
      });
    }
    
    // Get from persistent storage (keys not in memory)
    try {
      const keys = await this.storage.list('context_');
      for (const key of keys) {
        if (!this.memoryCache.has(key)) {
          const stored = await this.storage.get(key);
          if (stored) {
            try {
              const context = this.deserializeContext(stored);
              sessions.push({
                id: key,
                sessionId: context.sessionId,
                timestamp: context.timestamp
              });
            } catch (error) {
              console.warn(`Failed to parse stored context ${key}:`, error);
            }
          }
        }
      }
    } catch (error) {
      console.error('Failed to list sessions:', error);
    }

    return sessions.sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime());
  }

  /**
   * Search stored contexts by content
   */
  async searchContexts(query: string): Promise<StoredContext[]> {
    const results: StoredContext[] = [];
    const queryLower = query.toLowerCase();
    
    // Search memory cache
    for (const context of this.memoryCache.values()) {
      if (this.contextMatchesQuery(context, queryLower)) {
        results.push(context);
      }
    }
    
    // Search persistent storage
    try {
      const keys = await this.storage.list('context_');
      for (const key of keys) {
        if (!this.memoryCache.has(key)) {
          const stored = await this.storage.get(key);
          if (stored) {
            try {
              const context = this.deserializeContext(stored);
              if (this.contextMatchesQuery(context, queryLower)) {
                results.push(context);
              }
            } catch (error) {
              console.warn(`Failed to search context ${key}:`, error);
            }
          }
        }
      }
    } catch (error) {
      console.error('Failed to search contexts:', error);
    }

    return results.sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime());
  }

  /**
   * Get storage statistics
   */
  async getStatistics(): Promise<{
    totalSessions: number;
    totalMessages: number;
    storageSize: number;
    cacheSize: number;
    oldestSession: Date | null;
    newestSession: Date | null;
  }> {
    const sessions = await this.listSessions();
    
    let totalMessages = 0;
    let storageSize = 0;
    
    // Calculate from memory cache
    for (const context of this.memoryCache.values()) {
      totalMessages += context.messages.length;
      storageSize += JSON.stringify(context).length;
    }
    
    return {
      totalSessions: sessions.length,
      totalMessages,
      storageSize,
      cacheSize: this.memoryCache.size,
      oldestSession: sessions.length > 0 ? sessions[sessions.length - 1].timestamp : null,
      newestSession: sessions.length > 0 ? sessions[0].timestamp : null
    };
  }

  /**
   * Export all stored contexts
   */
  async exportAll(): Promise<string> {
    const sessions = await this.listSessions();
    const exportData: StoredContext[] = [];
    
    for (const session of sessions) {
      const context = await this.loadContext(session.sessionId);
      if (context) {
        exportData.push(context);
      }
    }
    
    return JSON.stringify({
      exportTime: new Date().toISOString(),
      version: '1.0.0',
      totalSessions: exportData.length,
      sessions: exportData
    }, null, 2);
  }

  /**
   * Import contexts from export data
   */
  async importAll(exportData: string): Promise<number> {
    try {
      const data = JSON.parse(exportData);
      let importedCount = 0;
      
      if (data.sessions && Array.isArray(data.sessions)) {
        for (const session of data.sessions) {
          try {
            await this.storeContext(
              session.sessionId,
              session.context,
              session.messages,
              session.excelState
            );
            importedCount++;
          } catch (error) {
            console.warn(`Failed to import session ${session.sessionId}:`, error);
          }
        }
      }
      
      console.log(`📥 Imported ${importedCount} sessions`);
      return importedCount;
    } catch (error) {
      console.error('Failed to import contexts:', error);
      throw new Error('Invalid export data format');
    }
  }

  /**
   * Manually save all pending contexts
   */
  async saveAll(): Promise<void> {
    console.log('💾 Manually saving all pending contexts...');
    
    for (const [contextId, context] of this.memoryCache) {
      try {
        const serialized = this.serializeContext(context);
        await this.storage.set(contextId, serialized);
      } catch (error) {
        console.error(`Failed to save context ${contextId}:`, error);
      }
    }
    
    this.saveQueue.clear();
    console.log('✅ All contexts saved');
  }

  /**
   * Clear all stored contexts
   */
  async clearAll(): Promise<void> {
    console.log('🗑️ Clearing all stored contexts...');
    
    // Clear memory cache
    this.memoryCache.clear();
    
    // Clear persistent storage
    const keys = await this.storage.list('context_');
    for (const key of keys) {
      await this.storage.delete(key);
    }
    
    this.saveQueue.clear();
    console.log('✅ All contexts cleared');
  }

  /**
   * Process save queue
   */
  private async processSaveQueue(): Promise<void> {
    if (this.isProcessingQueue || this.saveQueue.size === 0) {
      return;
    }
    
    this.isProcessingQueue = true;
    
    try {
      const contextsToSave = Array.from(this.saveQueue);
      this.saveQueue.clear();
      
      for (const contextId of contextsToSave) {
        const context = this.memoryCache.get(contextId);
        if (context) {
          try {
            const serialized = this.serializeContext(context);
            await this.storage.set(contextId, serialized);
          } catch (error) {
            console.error(`Failed to save context ${contextId}:`, error);
            // Re-queue for retry
            this.saveQueue.add(contextId);
          }
        }
      }
    } finally {
      this.isProcessingQueue = false;
    }
    
    // Process any new items that were added during processing
    if (this.saveQueue.size > 0) {
      setTimeout(() => this.processSaveQueue(), 1000);
    }
  }

  /**
   * Load recent sessions into memory cache
   */
  private async loadRecentSessions(): Promise<void> {
    try {
      const keys = await this.storage.list('context_');
      const recentKeys = keys.slice(0, 10); // Load 10 most recent
      
      for (const key of recentKeys) {
        if (!this.memoryCache.has(key)) {
          const stored = await this.storage.get(key);
          if (stored) {
            try {
              const context = this.deserializeContext(stored);
              this.memoryCache.set(key, context);
            } catch (error) {
              console.warn(`Failed to load context ${key}:`, error);
            }
          }
        }
      }
    } catch (error) {
      console.error('Failed to load recent sessions:', error);
    }
  }

  /**
   * Clean up old sessions
   */
  private async cleanupOldSessions(): Promise<void> {
    try {
      const sessions = await this.listSessions();
      
      if (sessions.length > this.config.maxStoredSessions) {
        const sessionsToDelete = sessions.slice(this.config.maxStoredSessions);
        
        for (const session of sessionsToDelete) {
          await this.deleteContext(session.sessionId);
        }
        
        console.log(`🧹 Cleaned up ${sessionsToDelete.length} old sessions`);
      }
    } catch (error) {
      console.error('Failed to cleanup old sessions:', error);
    }
  }

  /**
   * Limit messages to configured maximum
   */
  private limitMessages(messages: ConversationMessage[]): ConversationMessage[] {
    if (messages.length <= this.config.maxMessagesPerSession) {
      return messages;
    }
    
    return messages.slice(-this.config.maxMessagesPerSession);
  }

  /**
   * Serialize context for storage
   */
  private serializeContext(context: StoredContext): string {
    const data = JSON.stringify(context);
    
    if (this.config.compressionEnabled) {
      // Simple compression - in production, use proper compression library
      return data;
    }
    
    return data;
  }

  /**
   * Deserialize context from storage
   */
  private deserializeContext(data: string): StoredContext {
    const context = JSON.parse(data) as StoredContext;
    
    // Convert timestamp string back to Date
    context.timestamp = new Date(context.timestamp);
    
    // Convert message timestamps back to Date objects
    context.messages = context.messages.map(msg => ({
      ...msg,
      timestamp: new Date(msg.timestamp)
    }));
    
    return context;
  }

  /**
   * Check if context matches search query
   */
  private contextMatchesQuery(context: StoredContext, query: string): boolean {
    // Search in messages
    for (const message of context.messages) {
      if (message.content.toLowerCase().includes(query)) {
        return true;
      }
    }
    
    // Search in Excel state
    if (context.excelState.activeWorksheet.toLowerCase().includes(query)) {
      return true;
    }
    
    return false;
  }
}

export default ContextStore;