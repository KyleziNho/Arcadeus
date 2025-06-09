/**
 * ExtractionHistory.js - Track and manage extraction history
 * Provides undo/redo functionality and extraction session management
 */

class ExtractionHistory {
  constructor() {
    this.maxHistorySize = 50; // Maximum number of extraction sessions to keep
    this.history = [];
    this.currentIndex = -1;
    this.sessionId = null;
    this.storageKey = 'ma_extraction_history';
    
    this.init();
  }

  init() {
    this.loadHistory();
    console.log('✅ ExtractionHistory initialized with', this.history.length, 'sessions');
  }

  /**
   * Save a new extraction session
   */
  saveSession(sessionData) {
    const session = {
      id: this.generateSessionId(),
      timestamp: new Date().toISOString(),
      files: sessionData.files.map(f => ({
        name: f.name,
        type: f.type,
        size: f.size,
        lastModified: f.lastModified
      })),
      extractedData: sessionData.extractedData,
      appliedData: sessionData.appliedData || null,
      extractionResults: sessionData.extractionResults || null,
      metadata: {
        userAgent: navigator.userAgent,
        timestamp: Date.now(),
        sessionDuration: sessionData.duration || 0,
        extractionMethod: sessionData.method || 'unknown',
        version: '1.0'
      }
    };

    // Remove any sessions after current index (for undo/redo)
    if (this.currentIndex < this.history.length - 1) {
      this.history = this.history.slice(0, this.currentIndex + 1);
    }

    // Add new session
    this.history.push(session);
    this.currentIndex = this.history.length - 1;
    this.sessionId = session.id;

    // Trim history if it exceeds max size
    if (this.history.length > this.maxHistorySize) {
      this.history = this.history.slice(-this.maxHistorySize);
      this.currentIndex = this.history.length - 1;
    }

    this.persistHistory();
    this.notifyHistoryChange();

    console.log('💾 Saved extraction session:', session.id);
    return session.id;
  }

  /**
   * Get current extraction session
   */
  getCurrentSession() {
    if (this.currentIndex >= 0 && this.currentIndex < this.history.length) {
      return this.history[this.currentIndex];
    }
    return null;
  }

  /**
   * Go back to previous extraction session
   */
  undo() {
    if (this.canUndo()) {
      this.currentIndex--;
      const session = this.getCurrentSession();
      this.sessionId = session?.id || null;
      this.notifyHistoryChange();
      console.log('↶ Undid to session:', this.sessionId);
      return session;
    }
    return null;
  }

  /**
   * Go forward to next extraction session
   */
  redo() {
    if (this.canRedo()) {
      this.currentIndex++;
      const session = this.getCurrentSession();
      this.sessionId = session?.id || null;
      this.notifyHistoryChange();
      console.log('↷ Redid to session:', this.sessionId);
      return session;
    }
    return null;
  }

  /**
   * Check if undo is possible
   */
  canUndo() {
    return this.currentIndex > 0;
  }

  /**
   * Check if redo is possible
   */
  canRedo() {
    return this.currentIndex < this.history.length - 1;
  }

  /**
   * Restore a specific session by ID
   */
  restoreSession(sessionId) {
    const sessionIndex = this.history.findIndex(s => s.id === sessionId);
    if (sessionIndex >= 0) {
      this.currentIndex = sessionIndex;
      this.sessionId = sessionId;
      const session = this.getCurrentSession();
      this.notifyHistoryChange();
      console.log('🔄 Restored session:', sessionId);
      return session;
    }
    return null;
  }

  /**
   * Delete a specific session
   */
  deleteSession(sessionId) {
    const sessionIndex = this.history.findIndex(s => s.id === sessionId);
    if (sessionIndex >= 0) {
      this.history.splice(sessionIndex, 1);
      
      // Adjust current index if necessary
      if (this.currentIndex >= sessionIndex) {
        this.currentIndex = Math.max(0, this.currentIndex - 1);
      }
      
      // Update session ID if we deleted the current session
      if (this.sessionId === sessionId) {
        const newSession = this.getCurrentSession();
        this.sessionId = newSession?.id || null;
      }
      
      this.persistHistory();
      this.notifyHistoryChange();
      console.log('🗑️ Deleted session:', sessionId);
      return true;
    }
    return false;
  }

  /**
   * Get all extraction sessions
   */
  getAllSessions() {
    return this.history.map(session => ({
      ...session,
      isCurrent: session.id === this.sessionId
    }));
  }

  /**
   * Get sessions filtered by criteria
   */
  getFilteredSessions(options = {}) {
    const {
      startDate = null,
      endDate = null,
      fileNames = null,
      minConfidence = null,
      extractionMethod = null,
      limit = null
    } = options;

    let filtered = [...this.history];

    // Date range filter
    if (startDate) {
      filtered = filtered.filter(s => new Date(s.timestamp) >= new Date(startDate));
    }
    if (endDate) {
      filtered = filtered.filter(s => new Date(s.timestamp) <= new Date(endDate));
    }

    // File name filter
    if (fileNames && fileNames.length > 0) {
      filtered = filtered.filter(s => 
        s.files.some(f => 
          fileNames.some(name => f.name.toLowerCase().includes(name.toLowerCase()))
        )
      );
    }

    // Confidence filter
    if (minConfidence !== null) {
      filtered = filtered.filter(s => {
        const avgConfidence = this.calculateAverageConfidence(s.extractedData);
        return avgConfidence >= minConfidence;
      });
    }

    // Extraction method filter
    if (extractionMethod) {
      filtered = filtered.filter(s => 
        s.metadata.extractionMethod === extractionMethod
      );
    }

    // Sort by timestamp (newest first)
    filtered.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

    // Limit results
    if (limit && limit > 0) {
      filtered = filtered.slice(0, limit);
    }

    return filtered.map(session => ({
      ...session,
      isCurrent: session.id === this.sessionId
    }));
  }

  /**
   * Get extraction statistics
   */
  getStatistics() {
    const stats = {
      totalSessions: this.history.length,
      totalFiles: 0,
      averageConfidence: 0,
      extractionMethods: {},
      fileTypes: {},
      successRate: 0,
      timeRange: null
    };

    if (this.history.length === 0) {
      return stats;
    }

    let totalConfidence = 0;
    let successfulExtractions = 0;
    const timestamps = this.history.map(s => new Date(s.timestamp));

    for (const session of this.history) {
      // Count files
      stats.totalFiles += session.files.length;

      // Track extraction methods
      const method = session.metadata.extractionMethod || 'unknown';
      stats.extractionMethods[method] = (stats.extractionMethods[method] || 0) + 1;

      // Track file types
      for (const file of session.files) {
        const type = file.type || 'unknown';
        stats.fileTypes[type] = (stats.fileTypes[type] || 0) + 1;
      }

      // Calculate confidence
      const avgConfidence = this.calculateAverageConfidence(session.extractedData);
      totalConfidence += avgConfidence;
      
      if (avgConfidence > 0.5) {
        successfulExtractions++;
      }
    }

    stats.averageConfidence = totalConfidence / this.history.length;
    stats.successRate = (successfulExtractions / this.history.length) * 100;
    stats.timeRange = {
      earliest: new Date(Math.min(...timestamps)).toISOString(),
      latest: new Date(Math.max(...timestamps)).toISOString()
    };

    return stats;
  }

  /**
   * Export extraction history
   */
  exportHistory(format = 'json', options = {}) {
    const {
      includeFileContents = false,
      sessionsToExport = 'all'
    } = options;

    let sessions = this.history;
    
    if (sessionsToExport === 'current' && this.sessionId) {
      sessions = this.history.filter(s => s.id === this.sessionId);
    }

    const exportData = {
      metadata: {
        exportDate: new Date().toISOString(),
        version: '1.0',
        totalSessions: sessions.length,
        includeFileContents
      },
      sessions: sessions.map(session => {
        const exportSession = { ...session };
        
        if (!includeFileContents) {
          // Remove file content data to reduce size
          delete exportSession.files;
          exportSession.fileMetadata = session.files.map(f => ({
            name: f.name,
            type: f.type,
            size: f.size
          }));
        }
        
        return exportSession;
      })
    };

    if (format === 'json') {
      return JSON.stringify(exportData, null, 2);
    } else if (format === 'csv') {
      return this.convertToCSV(exportData);
    }

    return exportData;
  }

  /**
   * Import extraction history
   */
  importHistory(importData, options = {}) {
    const {
      mergeWithExisting = true,
      replaceExisting = false
    } = options;

    try {
      let data;
      if (typeof importData === 'string') {
        data = JSON.parse(importData);
      } else {
        data = importData;
      }

      if (!data.sessions || !Array.isArray(data.sessions)) {
        throw new Error('Invalid import data format');
      }

      const validSessions = data.sessions.filter(session => 
        session.id && session.timestamp && session.extractedData
      );

      if (replaceExisting) {
        this.history = validSessions;
        this.currentIndex = this.history.length - 1;
      } else if (mergeWithExisting) {
        // Merge, avoiding duplicates
        const existingIds = new Set(this.history.map(s => s.id));
        const newSessions = validSessions.filter(s => !existingIds.has(s.id));
        
        this.history = [...this.history, ...newSessions];
        this.history.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
        
        // Trim if necessary
        if (this.history.length > this.maxHistorySize) {
          this.history = this.history.slice(-this.maxHistorySize);
        }
        
        this.currentIndex = this.history.length - 1;
      }

      this.persistHistory();
      this.notifyHistoryChange();

      console.log('📥 Imported', validSessions.length, 'extraction sessions');
      return {
        imported: validSessions.length,
        total: this.history.length
      };

    } catch (error) {
      console.error('Failed to import extraction history:', error);
      throw error;
    }
  }

  /**
   * Clear all history
   */
  clearHistory() {
    this.history = [];
    this.currentIndex = -1;
    this.sessionId = null;
    this.persistHistory();
    this.notifyHistoryChange();
    console.log('🧹 Cleared extraction history');
  }

  /**
   * Calculate average confidence for session data
   */
  calculateAverageConfidence(extractedData) {
    if (!extractedData || typeof extractedData !== 'object') {
      return 0;
    }

    let totalConfidence = 0;
    let fieldCount = 0;

    for (const [key, value] of Object.entries(extractedData)) {
      if (key.startsWith('_')) continue; // Skip metadata
      
      if (value && typeof value === 'object' && value.confidence !== undefined) {
        totalConfidence += value.confidence;
        fieldCount++;
      }
    }

    return fieldCount > 0 ? totalConfidence / fieldCount : 0;
  }

  /**
   * Generate unique session ID
   */
  generateSessionId() {
    return 'session_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
  }

  /**
   * Load history from localStorage
   */
  loadHistory() {
    try {
      const stored = localStorage.getItem(this.storageKey);
      if (stored) {
        const data = JSON.parse(stored);
        this.history = data.history || [];
        this.currentIndex = data.currentIndex || -1;
        this.sessionId = data.sessionId || null;
        
        // Validate and clean up history
        this.history = this.history.filter(session => 
          session.id && session.timestamp && session.extractedData
        );
        
        // Adjust current index if necessary
        if (this.currentIndex >= this.history.length) {
          this.currentIndex = this.history.length - 1;
        }
        
        // Validate session ID
        if (this.sessionId && !this.history.find(s => s.id === this.sessionId)) {
          this.sessionId = null;
        }
      }
    } catch (error) {
      console.warn('Failed to load extraction history:', error);
      this.history = [];
      this.currentIndex = -1;
      this.sessionId = null;
    }
  }

  /**
   * Persist history to localStorage
   */
  persistHistory() {
    try {
      const data = {
        history: this.history,
        currentIndex: this.currentIndex,
        sessionId: this.sessionId,
        lastUpdated: new Date().toISOString()
      };
      
      localStorage.setItem(this.storageKey, JSON.stringify(data));
    } catch (error) {
      console.warn('Failed to persist extraction history:', error);
    }
  }

  /**
   * Convert export data to CSV format
   */
  convertToCSV(exportData) {
    const headers = [
      'Session ID',
      'Timestamp',
      'File Count',
      'File Names',
      'Extraction Method',
      'Average Confidence',
      'Fields Extracted',
      'Duration (ms)'
    ];

    const rows = exportData.sessions.map(session => {
      const fileNames = (session.files || session.fileMetadata || [])
        .map(f => f.name)
        .join('; ');
      
      const avgConfidence = this.calculateAverageConfidence(session.extractedData);
      
      const fieldsExtracted = Object.keys(session.extractedData || {})
        .filter(key => !key.startsWith('_'))
        .length;

      return [
        session.id,
        session.timestamp,
        (session.files || session.fileMetadata || []).length,
        fileNames,
        session.metadata?.extractionMethod || 'unknown',
        (avgConfidence * 100).toFixed(1) + '%',
        fieldsExtracted,
        session.metadata?.sessionDuration || 0
      ];
    });

    const csvContent = [headers, ...rows]
      .map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(','))
      .join('\n');

    return csvContent;
  }

  /**
   * Notify listeners of history changes
   */
  notifyHistoryChange() {
    const event = new CustomEvent('extractionHistoryChanged', {
      detail: {
        history: this.getAllSessions(),
        currentSessionId: this.sessionId,
        canUndo: this.canUndo(),
        canRedo: this.canRedo(),
        statistics: this.getStatistics()
      }
    });
    
    document.dispatchEvent(event);
  }

  /**
   * Add event listener for history changes
   */
  onHistoryChange(callback) {
    document.addEventListener('extractionHistoryChanged', callback);
  }

  /**
   * Remove event listener for history changes
   */
  offHistoryChange(callback) {
    document.removeEventListener('extractionHistoryChanged', callback);
  }

  /**
   * Get history summary for UI display
   */
  getSummary() {
    const current = this.getCurrentSession();
    const stats = this.getStatistics();
    
    return {
      currentSession: current ? {
        id: current.id,
        timestamp: current.timestamp,
        fileCount: current.files.length,
        confidence: this.calculateAverageConfidence(current.extractedData)
      } : null,
      totalSessions: stats.totalSessions,
      canUndo: this.canUndo(),
      canRedo: this.canRedo(),
      averageConfidence: stats.averageConfidence,
      successRate: stats.successRate
    };
  }

  /**
   * Create a comparison between two sessions
   */
  compareSession(sessionId1, sessionId2) {
    const session1 = this.history.find(s => s.id === sessionId1);
    const session2 = this.history.find(s => s.id === sessionId2);
    
    if (!session1 || !session2) {
      return null;
    }
    
    const comparison = {
      sessions: [session1, session2],
      differences: {},
      confidence: {
        session1: this.calculateAverageConfidence(session1.extractedData),
        session2: this.calculateAverageConfidence(session2.extractedData)
      }
    };
    
    // Find field differences
    const allFields = new Set([
      ...Object.keys(session1.extractedData || {}),
      ...Object.keys(session2.extractedData || {})
    ]);
    
    for (const field of allFields) {
      if (field.startsWith('_')) continue;
      
      const value1 = session1.extractedData?.[field]?.value;
      const value2 = session2.extractedData?.[field]?.value;
      
      if (value1 !== value2) {
        comparison.differences[field] = {
          session1: value1,
          session2: value2,
          confidence1: session1.extractedData?.[field]?.confidence || 0,
          confidence2: session2.extractedData?.[field]?.confidence || 0
        };
      }
    }
    
    return comparison;
  }
}

// Export for use
window.ExtractionHistory = ExtractionHistory;