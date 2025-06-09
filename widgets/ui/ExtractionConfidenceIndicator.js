/**
 * ExtractionConfidenceIndicator.js - Visual confidence indicators for extracted data
 * Displays confidence levels with color-coded icons and tooltips
 */

class ExtractionConfidenceIndicator {
  constructor() {
    this.confidenceThresholds = {
      high: 0.8,
      medium: 0.5,
      low: 0.3
    };
    
    this.confidenceConfig = {
      high: {
        icon: '✓',
        color: '#22c55e',
        backgroundColor: '#dcfce7',
        borderColor: '#22c55e',
        label: 'High Confidence',
        description: 'Data extracted with high confidence from multiple sources'
      },
      medium: {
        icon: '⚠',
        color: '#f59e0b',
        backgroundColor: '#fef3c7',
        borderColor: '#f59e0b',
        label: 'Medium Confidence',
        description: 'Data extracted with moderate confidence, may need review'
      },
      low: {
        icon: '!',
        color: '#ef4444',
        backgroundColor: '#fee2e2',
        borderColor: '#ef4444',
        label: 'Low Confidence',
        description: 'Data extracted with low confidence, requires manual verification'
      },
      none: {
        icon: '?',
        color: '#6b7280',
        backgroundColor: '#f3f4f6',
        borderColor: '#6b7280',
        label: 'No Data',
        description: 'No data found or extraction failed'
      }
    };
    
    this.activeIndicators = new Map();
    this.init();
  }

  init() {
    this.injectStyles();
    console.log('✅ ExtractionConfidenceIndicator initialized');
  }

  /**
   * Inject CSS styles for confidence indicators
   */
  injectStyles() {
    if (document.getElementById('confidence-indicator-styles')) return;
    
    const styles = document.createElement('style');
    styles.id = 'confidence-indicator-styles';
    styles.textContent = `
      .confidence-indicator {
        position: relative;
        display: inline-flex;
        align-items: center;
        margin-left: 8px;
        z-index: 1000;
      }
      
      .confidence-badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 20px;
        height: 20px;
        border-radius: 50%;
        font-size: 12px;
        font-weight: bold;
        border: 2px solid;
        cursor: help;
        transition: all 0.2s ease;
        user-select: none;
      }
      
      .confidence-badge:hover {
        transform: scale(1.1);
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
      }
      
      .confidence-tooltip {
        position: absolute;
        bottom: 100%;
        left: 50%;
        transform: translateX(-50%);
        margin-bottom: 8px;
        padding: 8px 12px;
        background: #1f2937;
        color: white;
        border-radius: 6px;
        font-size: 12px;
        white-space: nowrap;
        opacity: 0;
        visibility: hidden;
        transition: all 0.2s ease;
        z-index: 1001;
        pointer-events: none;
      }
      
      .confidence-tooltip::after {
        content: '';
        position: absolute;
        top: 100%;
        left: 50%;
        transform: translateX(-50%);
        border: 5px solid transparent;
        border-top-color: #1f2937;
      }
      
      .confidence-indicator:hover .confidence-tooltip {
        opacity: 1;
        visibility: visible;
      }
      
      .confidence-details {
        display: flex;
        flex-direction: column;
        gap: 2px;
      }
      
      .confidence-percentage {
        font-weight: bold;
      }
      
      .confidence-source {
        font-size: 11px;
        opacity: 0.8;
      }
      
      .confidence-metrics {
        display: flex;
        align-items: center;
        gap: 4px;
        margin-top: 4px;
      }
      
      .confidence-bar {
        width: 60px;
        height: 4px;
        background: #e5e7eb;
        border-radius: 2px;
        overflow: hidden;
      }
      
      .confidence-fill {
        height: 100%;
        transition: width 0.3s ease;
        border-radius: 2px;
      }
      
      .field-updated {
        animation: confidenceHighlight 1s ease;
      }
      
      @keyframes confidenceHighlight {
        0% { box-shadow: 0 0 0 0 rgba(59, 130, 246, 0.7); }
        50% { box-shadow: 0 0 0 4px rgba(59, 130, 246, 0.4); }
        100% { box-shadow: 0 0 0 0 rgba(59, 130, 246, 0); }
      }
      
      .extraction-summary {
        margin-top: 12px;
        padding: 12px;
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        font-size: 12px;
      }
      
      .extraction-summary-title {
        font-weight: bold;
        margin-bottom: 8px;
        color: #374151;
      }
      
      .extraction-stats {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
        gap: 8px;
      }
      
      .extraction-stat {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 4px 8px;
        background: white;
        border-radius: 4px;
        border: 1px solid #e5e7eb;
      }
      
      .extraction-stat-value {
        font-weight: bold;
      }
    `;
    
    document.head.appendChild(styles);
  }

  /**
   * Add confidence indicator to a form field
   */
  addToField(fieldElement, data, options = {}) {
    if (!fieldElement || !data) return null;
    
    const {
      showTooltip = true,
      showProgressBar = false,
      position = 'after',
      customMessage = null
    } = options;
    
    // Remove existing indicator
    this.removeFromField(fieldElement);
    
    // Determine confidence level
    const confidenceLevel = this.getConfidenceLevel(data.confidence || 0);
    const config = this.confidenceConfig[confidenceLevel];
    
    // Create indicator element
    const indicator = this.createIndicatorElement(data, config, {
      showTooltip,
      showProgressBar,
      customMessage
    });
    
    // Position the indicator
    this.positionIndicator(fieldElement, indicator, position);
    
    // Track the indicator
    const fieldId = fieldElement.id || fieldElement.name || 'unknown';
    this.activeIndicators.set(fieldId, indicator);
    
    return indicator;
  }

  /**
   * Create the indicator DOM element
   */
  createIndicatorElement(data, config, options) {
    const indicator = document.createElement('div');
    indicator.className = 'confidence-indicator';
    
    // Create badge
    const badge = document.createElement('div');
    badge.className = 'confidence-badge';
    badge.style.color = config.color;
    badge.style.backgroundColor = config.backgroundColor;
    badge.style.borderColor = config.borderColor;
    badge.textContent = config.icon;
    
    indicator.appendChild(badge);
    
    // Create tooltip
    if (options.showTooltip) {
      const tooltip = this.createTooltip(data, config, options.customMessage);
      indicator.appendChild(tooltip);
    }
    
    // Add progress bar
    if (options.showProgressBar) {
      const progressBar = this.createProgressBar(data.confidence || 0, config);
      indicator.appendChild(progressBar);
    }
    
    return indicator;
  }

  /**
   * Create tooltip element
   */
  createTooltip(data, config, customMessage) {
    const tooltip = document.createElement('div');
    tooltip.className = 'confidence-tooltip';
    
    const details = document.createElement('div');
    details.className = 'confidence-details';
    
    // Confidence percentage
    const percentage = document.createElement('div');
    percentage.className = 'confidence-percentage';
    percentage.textContent = customMessage || `${config.label}: ${Math.round((data.confidence || 0) * 100)}%`;
    details.appendChild(percentage);
    
    // Source information
    if (data.source) {
      const source = document.createElement('div');
      source.className = 'confidence-source';
      source.textContent = `Source: ${this.formatSource(data.source)}`;
      details.appendChild(source);
    }
    
    // Description
    const description = document.createElement('div');
    description.textContent = config.description;
    details.appendChild(description);
    
    tooltip.appendChild(details);
    return tooltip;
  }

  /**
   * Create progress bar element
   */
  createProgressBar(confidence, config) {
    const metrics = document.createElement('div');
    metrics.className = 'confidence-metrics';
    
    const bar = document.createElement('div');
    bar.className = 'confidence-bar';
    
    const fill = document.createElement('div');
    fill.className = 'confidence-fill';
    fill.style.width = `${(confidence || 0) * 100}%`;
    fill.style.backgroundColor = config.color;
    
    bar.appendChild(fill);
    metrics.appendChild(bar);
    
    const percentage = document.createElement('span');
    percentage.textContent = `${Math.round((confidence || 0) * 100)}%`;
    percentage.style.fontSize = '10px';
    percentage.style.color = config.color;
    metrics.appendChild(percentage);
    
    return metrics;
  }

  /**
   * Position indicator relative to field
   */
  positionIndicator(fieldElement, indicator, position) {
    const container = fieldElement.parentElement;
    
    if (position === 'after') {
      // Insert after the field
      if (fieldElement.nextSibling) {
        container.insertBefore(indicator, fieldElement.nextSibling);
      } else {
        container.appendChild(indicator);
      }
    } else if (position === 'before') {
      // Insert before the field
      container.insertBefore(indicator, fieldElement);
    } else if (position === 'overlay') {
      // Overlay on the field
      fieldElement.style.position = 'relative';
      indicator.style.position = 'absolute';
      indicator.style.right = '8px';
      indicator.style.top = '50%';
      indicator.style.transform = 'translateY(-50%)';
      fieldElement.appendChild(indicator);
    }
  }

  /**
   * Remove confidence indicator from field
   */
  removeFromField(fieldElement) {
    const fieldId = fieldElement.id || fieldElement.name || 'unknown';
    const existingIndicator = this.activeIndicators.get(fieldId);
    
    if (existingIndicator && existingIndicator.parentElement) {
      existingIndicator.parentElement.removeChild(existingIndicator);
      this.activeIndicators.delete(fieldId);
    }
    
    // Also remove any indicators in the same container
    const container = fieldElement.parentElement;
    if (container) {
      const indicators = container.querySelectorAll('.confidence-indicator');
      indicators.forEach(indicator => {
        if (indicator.parentElement === container) {
          indicator.remove();
        }
      });
    }
  }

  /**
   * Update confidence for existing indicator
   */
  updateConfidence(fieldElement, newData) {
    this.addToField(fieldElement, newData);
  }

  /**
   * Get confidence level category
   */
  getConfidenceLevel(confidence) {
    if (confidence === 0 || confidence === null || confidence === undefined) {
      return 'none';
    } else if (confidence >= this.confidenceThresholds.high) {
      return 'high';
    } else if (confidence >= this.confidenceThresholds.medium) {
      return 'medium';
    } else {
      return 'low';
    }
  }

  /**
   * Format source for display
   */
  formatSource(source) {
    const sourceLabels = {
      'ai_extraction': 'AI Analysis',
      'pattern_matching': 'Pattern Recognition',
      'calculated': 'Calculated',
      'inferred': 'Inferred',
      'not_found': 'Not Found',
      'filename': 'File Name',
      'content_pattern': 'Document Content',
      'csv_pattern': 'CSV Data',
      'category_pattern': 'Category Match',
      'context_matching': 'Context Analysis',
      'inferred_earliest': 'Earliest Date',
      'inferred_latest': 'Latest Date'
    };
    
    return sourceLabels[source] || source.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
  }

  /**
   * Create extraction summary for a section
   */
  createExtractionSummary(sectionData, containerId) {
    const container = document.getElementById(containerId);
    if (!container) return null;
    
    // Remove existing summary
    const existing = container.querySelector('.extraction-summary');
    if (existing) existing.remove();
    
    const summary = document.createElement('div');
    summary.className = 'extraction-summary';
    
    const title = document.createElement('div');
    title.className = 'extraction-summary-title';
    title.textContent = 'Extraction Summary';
    summary.appendChild(title);
    
    const stats = this.calculateExtractionStats(sectionData);
    const statsGrid = document.createElement('div');
    statsGrid.className = 'extraction-stats';
    
    for (const [key, value] of Object.entries(stats)) {
      const stat = document.createElement('div');
      stat.className = 'extraction-stat';
      
      const label = document.createElement('span');
      label.textContent = this.formatStatLabel(key);
      
      const valueSpan = document.createElement('span');
      valueSpan.className = 'extraction-stat-value';
      valueSpan.textContent = value;
      
      stat.appendChild(label);
      stat.appendChild(valueSpan);
      statsGrid.appendChild(stat);
    }
    
    summary.appendChild(statsGrid);
    container.appendChild(summary);
    
    return summary;
  }

  /**
   * Calculate extraction statistics
   */
  calculateExtractionStats(data) {
    let totalFields = 0;
    let extractedFields = 0;
    let highConfidenceFields = 0;
    let calculatedFields = 0;
    
    for (const [key, value] of Object.entries(data)) {
      if (key.startsWith('_')) continue; // Skip metadata
      
      totalFields++;
      
      if (value && value.value !== null && value.value !== undefined) {
        extractedFields++;
        
        if (value.confidence >= this.confidenceThresholds.high) {
          highConfidenceFields++;
        }
        
        if (value.source === 'calculated') {
          calculatedFields++;
        }
      }
    }
    
    const extractionRate = totalFields > 0 ? Math.round((extractedFields / totalFields) * 100) : 0;
    
    return {
      totalFields,
      extractedFields,
      extractionRate: `${extractionRate}%`,
      highConfidenceFields,
      calculatedFields
    };
  }

  /**
   * Format statistic labels
   */
  formatStatLabel(key) {
    const labels = {
      totalFields: 'Total Fields',
      extractedFields: 'Extracted',
      extractionRate: 'Success Rate',
      highConfidenceFields: 'High Confidence',
      calculatedFields: 'Calculated'
    };
    
    return labels[key] || key;
  }

  /**
   * Show confidence legend
   */
  showLegend(containerId) {
    const container = document.getElementById(containerId);
    if (!container) return;
    
    const legend = document.createElement('div');
    legend.className = 'confidence-legend';
    legend.style.cssText = `
      display: flex;
      gap: 16px;
      margin: 12px 0;
      padding: 12px;
      background: #f8fafc;
      border-radius: 6px;
      font-size: 12px;
    `;
    
    for (const [level, config] of Object.entries(this.confidenceConfig)) {
      if (level === 'none') continue;
      
      const item = document.createElement('div');
      item.style.cssText = 'display: flex; align-items: center; gap: 6px;';
      
      const badge = document.createElement('div');
      badge.className = 'confidence-badge';
      badge.style.cssText = `
        width: 16px;
        height: 16px;
        font-size: 10px;
        color: ${config.color};
        background-color: ${config.backgroundColor};
        border: 1px solid ${config.borderColor};
      `;
      badge.textContent = config.icon;
      
      const label = document.createElement('span');
      label.textContent = config.label;
      
      item.appendChild(badge);
      item.appendChild(label);
      legend.appendChild(item);
    }
    
    container.appendChild(legend);
  }

  /**
   * Animate confidence indicator
   */
  animateIndicator(fieldElement, type = 'pulse') {
    const fieldId = fieldElement.id || fieldElement.name || 'unknown';
    const indicator = this.activeIndicators.get(fieldId);
    
    if (!indicator) return;
    
    const badge = indicator.querySelector('.confidence-badge');
    if (!badge) return;
    
    if (type === 'pulse') {
      badge.style.animation = 'confidenceHighlight 1s ease';
      setTimeout(() => {
        badge.style.animation = '';
      }, 1000);
    }
  }

  /**
   * Remove all indicators
   */
  clearAllIndicators() {
    this.activeIndicators.forEach(indicator => {
      if (indicator.parentElement) {
        indicator.parentElement.removeChild(indicator);
      }
    });
    this.activeIndicators.clear();
  }

  /**
   * Get confidence statistics across all active indicators
   */
  getOverallStats() {
    const stats = {
      total: this.activeIndicators.size,
      high: 0,
      medium: 0,
      low: 0,
      none: 0
    };
    
    this.activeIndicators.forEach(indicator => {
      const badge = indicator.querySelector('.confidence-badge');
      if (badge) {
        const level = this.getConfidenceLevelFromBadge(badge);
        stats[level]++;
      }
    });
    
    return stats;
  }

  /**
   * Get confidence level from badge element
   */
  getConfidenceLevelFromBadge(badge) {
    const icon = badge.textContent;
    
    for (const [level, config] of Object.entries(this.confidenceConfig)) {
      if (config.icon === icon) {
        return level;
      }
    }
    
    return 'none';
  }
}

// Export for use
window.ExtractionConfidenceIndicator = ExtractionConfidenceIndicator;