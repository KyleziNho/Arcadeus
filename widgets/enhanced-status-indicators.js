/**
 * Enhanced Status Indicators for Multi-Agent Processing
 * Professional progress visualization for M&A analysis workflow
 */

class EnhancedStatusIndicators {
  constructor() {
    this.currentStatus = null;
    this.statusElement = null;
    this.animationInterval = null;
    
    // Listen for multi-agent progress events
    window.addEventListener('multiAgentProgress', (event) => {
      this.updateProgress(event.detail);
    });
    
    console.log('📊 Enhanced Status Indicators initialized');
  }

  /**
   * Show multi-agent processing progress with professional styling
   */
  showMultiAgentProgress(queryId, stage, progress) {
    this.ensureStatusElement();
    
    const stageInfo = this.getStageInfo(stage);
    this.currentStatus = {
      queryId,
      stage,
      progress,
      stageInfo,
      startTime: this.currentStatus?.startTime || Date.now()
    };
    
    this.renderProgressIndicator();
  }

  /**
   * Get stage information with icons and descriptions
   */
  getStageInfo(stage) {
    const stages = {
      initializing: {
        icon: '🎯',
        title: 'Initializing Analysis',
        description: 'Preparing multi-agent system',
        color: '#3498db'
      },
      analyzing_query: {
        icon: '🧠',
        title: 'Analyzing Query',
        description: 'Understanding your request and routing to appropriate agents',
        color: '#9b59b6'
      },
      fetching_structure: {
        icon: '📊',
        title: 'Reading Excel Structure',
        description: 'Analyzing workbook structure and extracting key metrics',
        color: '#2ecc71'
      },
      agent_excelStructure: {
        icon: '🏗️',
        title: 'Excel Structure Agent',
        description: 'Analyzing formulas, dependencies, and data organization',
        color: '#f39c12'
      },
      agent_financialAnalysis: {
        icon: '💰',
        title: 'Financial Analysis Agent',
        description: 'Performing M&A financial analysis with investment banking expertise',
        color: '#e74c3c'
      },
      agent_dataValidation: {
        icon: '✅',
        title: 'Data Validation Agent',
        description: 'Checking model consistency and data quality',
        color: '#1abc9c'
      },
      synthesizing: {
        icon: '🎭',
        title: 'Synthesizing Response',
        description: 'Combining insights from all agents into professional analysis',
        color: '#34495e'
      },
      completed: {
        icon: '🎉',
        title: 'Analysis Complete',
        description: 'Professional M&A analysis ready',
        color: '#27ae60'
      },
      error: {
        icon: '❌',
        title: 'Analysis Error',
        description: 'An error occurred during processing',
        color: '#e74c3c'
      }
    };
    
    return stages[stage] || stages.initializing;
  }

  /**
   * Update progress from event
   */
  updateProgress(detail) {
    const { queryId, stage, progress } = detail;
    this.showMultiAgentProgress(queryId, stage, progress);
  }

  /**
   * Render the progress indicator with professional styling
   */
  renderProgressIndicator() {
    if (!this.statusElement || !this.currentStatus) return;
    
    const { stage, progress, stageInfo, startTime } = this.currentStatus;
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    
    // Create progress bar HTML
    const progressBarHtml = this.createProgressBar(progress, stageInfo.color);
    
    // Create agent timeline
    const timelineHtml = this.createAgentTimeline(stage);
    
    // Create status display
    const statusHtml = `
      <div class="multi-agent-status">
        <div class="status-header">
          <div class="status-icon">${stageInfo.icon}</div>
          <div class="status-content">
            <div class="status-title">${stageInfo.title}</div>
            <div class="status-description">${stageInfo.description}</div>
          </div>
          <div class="status-timing">
            <div class="elapsed-time">${elapsed}s</div>
            <div class="progress-percentage">${progress >= 0 ? Math.round(progress) : 0}%</div>
          </div>
        </div>
        ${progressBarHtml}
        ${timelineHtml}
      </div>
    `;
    
    this.statusElement.innerHTML = statusHtml;
    
    // Auto-hide after completion
    if (stage === 'completed') {
      this.scheduleHide(5000); // Hide after 5 seconds
    } else if (stage === 'error') {
      this.scheduleHide(10000); // Hide after 10 seconds for errors
    }
  }

  /**
   * Create animated progress bar
   */
  createProgressBar(progress, color) {
    const safeProgress = Math.max(0, Math.min(100, progress));
    
    return `
      <div class="progress-container">
        <div class="progress-bar">
          <div class="progress-fill" style="width: ${safeProgress}%; background: ${color};">
            <div class="progress-shine"></div>
          </div>
        </div>
      </div>
    `;
  }

  /**
   * Create agent processing timeline
   */
  createAgentTimeline(currentStage) {
    const agents = [
      { key: 'analyzing_query', icon: '🧠', name: 'Query Analysis' },
      { key: 'fetching_structure', icon: '📊', name: 'Excel Reading' },
      { key: 'agent_excelStructure', icon: '🏗️', name: 'Structure Agent' },
      { key: 'agent_financialAnalysis', icon: '💰', name: 'Financial Agent' },
      { key: 'agent_dataValidation', icon: '✅', name: 'Validation Agent' },
      { key: 'synthesizing', icon: '🎭', name: 'Synthesis' }
    ];
    
    const timelineItems = agents.map(agent => {
      let status = 'pending';
      if (currentStage === agent.key) {
        status = 'active';
      } else if (this.isStageCompleted(agent.key, currentStage)) {
        status = 'completed';
      }
      
      return `
        <div class="timeline-item ${status}">
          <div class="timeline-icon">${agent.icon}</div>
          <div class="timeline-name">${agent.name}</div>
        </div>
      `;
    }).join('');
    
    return `
      <div class="agent-timeline">
        ${timelineItems}
      </div>
    `;
  }

  /**
   * Check if a stage has been completed
   */
  isStageCompleted(stageKey, currentStage) {
    const stageOrder = [
      'initializing',
      'analyzing_query',
      'fetching_structure',
      'agent_excelStructure',
      'agent_financialAnalysis', 
      'agent_dataValidation',
      'synthesizing',
      'completed'
    ];
    
    const stageIndex = stageOrder.indexOf(stageKey);
    const currentIndex = stageOrder.indexOf(currentStage);
    
    return stageIndex < currentIndex || currentStage === 'completed';
  }

  /**
   * Ensure status element exists
   */
  ensureStatusElement() {
    if (this.statusElement) return;
    
    // Try to find existing element
    this.statusElement = document.getElementById('multiAgentStatus');
    
    if (!this.statusElement) {
      // Create new status element
      this.statusElement = document.createElement('div');
      this.statusElement.id = 'multiAgentStatus';
      this.statusElement.className = 'multi-agent-status-container';
      
      // Insert into chat container
      const chatContainer = document.getElementById('chatMessages') || document.body;
      chatContainer.appendChild(this.statusElement);
    }
  }

  /**
   * Schedule automatic hide
   */
  scheduleHide(delay) {
    setTimeout(() => {
      if (this.statusElement) {
        this.statusElement.style.transition = 'opacity 1s ease-out';
        this.statusElement.style.opacity = '0';
        
        setTimeout(() => {
          if (this.statusElement) {
            this.statusElement.innerHTML = '';
            this.statusElement.style.opacity = '1';
          }
        }, 1000);
      }
    }, delay);
  }

  /**
   * Hide status indicators
   */
  hide() {
    if (this.statusElement) {
      this.statusElement.innerHTML = '';
    }
    this.currentStatus = null;
  }

  /**
   * Show error status
   */
  showError(error, queryId) {
    this.showMultiAgentProgress(queryId, 'error', -1);
    
    // Update description with specific error
    setTimeout(() => {
      if (this.statusElement) {
        const descElement = this.statusElement.querySelector('.status-description');
        if (descElement) {
          descElement.textContent = `Error: ${error.message || error}`;
        }
      }
    }, 100);
  }

  /**
   * Show success completion with results summary
   */
  showCompletion(results, processingTime) {
    if (!this.statusElement || !results.metadata) return;
    
    const { agentsUsed, structureAnalyzed, queryType } = results.metadata;
    
    // Update the completion description with results
    setTimeout(() => {
      const descElement = this.statusElement.querySelector('.status-description');
      if (descElement) {
        descElement.innerHTML = `
          Analyzed ${structureAnalyzed} sheets using ${agentsUsed?.length || 0} agents
          <br><small>Query type: ${queryType} • Processing time: ${processingTime.toFixed(0)}ms</small>
        `;
      }
    }, 100);
  }

  /**
   * Get current status
   */
  getCurrentStatus() {
    return this.currentStatus;
  }
}

// Export for global use
window.EnhancedStatusIndicators = EnhancedStatusIndicators;
window.enhancedStatusIndicators = new EnhancedStatusIndicators();

console.log('📊 Enhanced Status Indicators loaded with:');
console.log('  ✅ Multi-agent progress visualization');
console.log('  ✅ Professional timeline display');
console.log('  ✅ Animated progress bars');
console.log('  ✅ Real-time status updates');
console.log('  ✅ Auto-hide on completion');