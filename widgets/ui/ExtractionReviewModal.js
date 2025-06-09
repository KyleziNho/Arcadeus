/**
 * ExtractionReviewModal.js - Modal for reviewing and editing extracted data
 * Allows users to review, edit, and approve extracted data before applying to form
 */

class ExtractionReviewModal {
  constructor() {
    this.isOpen = false;
    this.currentData = null;
    this.onApprove = null;
    this.onReject = null;
    this.confidenceIndicator = null;
    this.editedValues = new Map();
    
    this.init();
  }

  init() {
    this.injectStyles();
    this.createModalStructure();
    this.bindEvents();
    console.log('✅ ExtractionReviewModal initialized');
  }

  /**
   * Inject CSS styles for the modal
   */
  injectStyles() {
    if (document.getElementById('extraction-review-modal-styles')) return;
    
    const styles = document.createElement('style');
    styles.id = 'extraction-review-modal-styles';
    styles.textContent = `
      .extraction-review-overlay {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.5);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 10000;
        opacity: 0;
        visibility: hidden;
        transition: all 0.3s ease;
      }
      
      .extraction-review-overlay.active {
        opacity: 1;
        visibility: visible;
      }
      
      .extraction-review-modal {
        background: white;
        border-radius: 12px;
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
        max-width: 900px;
        width: 90%;
        max-height: 80vh;
        overflow: hidden;
        transform: scale(0.95);
        transition: transform 0.3s ease;
      }
      
      .extraction-review-overlay.active .extraction-review-modal {
        transform: scale(1);
      }
      
      .extraction-review-header {
        padding: 20px 24px;
        border-bottom: 1px solid #e5e7eb;
        background: #f8fafc;
      }
      
      .extraction-review-title {
        font-size: 18px;
        font-weight: 600;
        color: #111827;
        margin: 0;
      }
      
      .extraction-review-subtitle {
        font-size: 14px;
        color: #6b7280;
        margin: 4px 0 0 0;
      }
      
      .extraction-review-body {
        padding: 0;
        overflow-y: auto;
        max-height: calc(80vh - 140px);
      }
      
      .extraction-review-section {
        border-bottom: 1px solid #f3f4f6;
      }
      
      .extraction-review-section:last-child {
        border-bottom: none;
      }
      
      .extraction-section-header {
        padding: 16px 24px;
        background: #f9fafb;
        border-bottom: 1px solid #f3f4f6;
        font-weight: 600;
        color: #374151;
        font-size: 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
      }
      
      .extraction-section-stats {
        font-size: 12px;
        color: #6b7280;
        font-weight: normal;
      }
      
      .extraction-fields {
        padding: 16px 24px;
      }
      
      .extraction-field {
        display: grid;
        grid-template-columns: 200px 1fr auto;
        gap: 16px;
        align-items: center;
        padding: 12px 0;
        border-bottom: 1px solid #f3f4f6;
      }
      
      .extraction-field:last-child {
        border-bottom: none;
      }
      
      .extraction-field-label {
        font-weight: 500;
        color: #374151;
        font-size: 14px;
      }
      
      .extraction-field-value {
        display: flex;
        align-items: center;
        gap: 8px;
      }
      
      .extraction-field-input {
        flex: 1;
        padding: 8px 12px;
        border: 1px solid #d1d5db;
        border-radius: 6px;
        font-size: 14px;
        transition: border-color 0.2s ease;
      }
      
      .extraction-field-input:focus {
        outline: none;
        border-color: #3b82f6;
        box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
      }
      
      .extraction-field-input.edited {
        border-color: #f59e0b;
        background-color: #fffbeb;
      }
      
      .extraction-field-original {
        font-size: 12px;
        color: #6b7280;
        font-style: italic;
      }
      
      .extraction-field-actions {
        display: flex;
        gap: 4px;
      }
      
      .extraction-field-btn {
        width: 24px;
        height: 24px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 12px;
        transition: all 0.2s ease;
      }
      
      .extraction-field-btn.reset {
        background: #f3f4f6;
        color: #6b7280;
      }
      
      .extraction-field-btn.reset:hover {
        background: #e5e7eb;
        color: #374151;
      }
      
      .extraction-field-btn.clear {
        background: #fee2e2;
        color: #dc2626;
      }
      
      .extraction-field-btn.clear:hover {
        background: #fecaca;
      }
      
      .extraction-field-empty {
        color: #9ca3af;
        font-style: italic;
        font-size: 14px;
      }
      
      .extraction-review-footer {
        padding: 20px 24px;
        border-top: 1px solid #e5e7eb;
        background: #f8fafc;
        display: flex;
        justify-content: space-between;
        align-items: center;
      }
      
      .extraction-review-stats {
        display: flex;
        gap: 20px;
        font-size: 12px;
        color: #6b7280;
      }
      
      .extraction-review-stat {
        display: flex;
        align-items: center;
        gap: 4px;
      }
      
      .extraction-review-actions {
        display: flex;
        gap: 12px;
      }
      
      .extraction-review-btn {
        padding: 10px 20px;
        border-radius: 6px;
        font-size: 14px;
        font-weight: 500;
        cursor: pointer;
        transition: all 0.2s ease;
        border: 1px solid;
      }
      
      .extraction-review-btn.secondary {
        background: white;
        color: #374151;
        border-color: #d1d5db;
      }
      
      .extraction-review-btn.secondary:hover {
        background: #f9fafb;
        border-color: #9ca3af;
      }
      
      .extraction-review-btn.primary {
        background: #3b82f6;
        color: white;
        border-color: #3b82f6;
      }
      
      .extraction-review-btn.primary:hover {
        background: #2563eb;
        border-color: #2563eb;
      }
      
      .extraction-review-btn.danger {
        background: #dc2626;
        color: white;
        border-color: #dc2626;
      }
      
      .extraction-review-btn.danger:hover {
        background: #b91c1c;
        border-color: #b91c1c;
      }
      
      .extraction-array-items {
        display: flex;
        flex-direction: column;
        gap: 8px;
      }
      
      .extraction-array-item {
        display: grid;
        grid-template-columns: 1fr auto auto;
        gap: 8px;
        align-items: center;
        padding: 8px 12px;
        background: #f9fafb;
        border: 1px solid #e5e7eb;
        border-radius: 6px;
      }
      
      .extraction-array-item-input {
        padding: 4px 8px;
        border: 1px solid #d1d5db;
        border-radius: 4px;
        font-size: 13px;
      }
      
      .extraction-array-actions {
        display: flex;
        gap: 4px;
      }
      
      .extraction-add-item {
        margin-top: 8px;
        padding: 6px 12px;
        background: #f3f4f6;
        border: 1px dashed #d1d5db;
        border-radius: 6px;
        cursor: pointer;
        text-align: center;
        font-size: 12px;
        color: #6b7280;
        transition: all 0.2s ease;
      }
      
      .extraction-add-item:hover {
        background: #e5e7eb;
        border-color: #9ca3af;
        color: #374151;
      }
      
      .extraction-confidence-summary {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 12px;
      }
      
      .extraction-overall-confidence {
        font-size: 14px;
        font-weight: 600;
      }
      
      .extraction-confidence-breakdown {
        display: flex;
        gap: 12px;
        font-size: 12px;
      }
      
      .extraction-confidence-item {
        display: flex;
        align-items: center;
        gap: 4px;
      }
      
      @media (max-width: 768px) {
        .extraction-review-modal {
          width: 95%;
          margin: 20px;
        }
        
        .extraction-field {
          grid-template-columns: 1fr;
          gap: 8px;
        }
        
        .extraction-field-label {
          font-weight: 600;
        }
      }
    `;
    
    document.head.appendChild(styles);
  }

  /**
   * Create modal DOM structure
   */
  createModalStructure() {
    const overlay = document.createElement('div');
    overlay.className = 'extraction-review-overlay';
    overlay.id = 'extraction-review-overlay';
    
    const modal = document.createElement('div');
    modal.className = 'extraction-review-modal';
    modal.id = 'extraction-review-modal';
    
    // Header
    const header = document.createElement('div');
    header.className = 'extraction-review-header';
    header.innerHTML = `
      <h2 class="extraction-review-title">Review Extracted Data</h2>
      <p class="extraction-review-subtitle">Review and edit the extracted data before applying to your model</p>
    `;
    
    // Body
    const body = document.createElement('div');
    body.className = 'extraction-review-body';
    body.id = 'extraction-review-content';
    
    // Footer
    const footer = document.createElement('div');
    footer.className = 'extraction-review-footer';
    footer.innerHTML = `
      <div class="extraction-review-stats" id="extraction-review-stats"></div>
      <div class="extraction-review-actions">
        <button class="extraction-review-btn secondary" id="extraction-review-cancel">Cancel</button>
        <button class="extraction-review-btn danger" id="extraction-review-reject">Reject All</button>
        <button class="extraction-review-btn primary" id="extraction-review-approve">Apply Changes</button>
      </div>
    `;
    
    modal.appendChild(header);
    modal.appendChild(body);
    modal.appendChild(footer);
    overlay.appendChild(modal);
    
    document.body.appendChild(overlay);
  }

  /**
   * Bind event handlers
   */
  bindEvents() {
    // Close modal on overlay click
    document.getElementById('extraction-review-overlay').addEventListener('click', (e) => {
      if (e.target.id === 'extraction-review-overlay') {
        this.close();
      }
    });
    
    // Button handlers
    document.getElementById('extraction-review-cancel').addEventListener('click', () => {
      this.close();
    });
    
    document.getElementById('extraction-review-reject').addEventListener('click', () => {
      this.reject();
    });
    
    document.getElementById('extraction-review-approve').addEventListener('click', () => {
      this.approve();
    });
    
    // Escape key handler
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && this.isOpen) {
        this.close();
      }
    });
  }

  /**
   * Show the review modal with extraction data
   */
  show(extractionData, options = {}) {
    const {
      title = 'Review Extracted Data',
      subtitle = 'Review and edit the extracted data before applying to your model',
      onApprove = null,
      onReject = null,
      confidenceIndicator = null
    } = options;
    
    this.currentData = extractionData;
    this.onApprove = onApprove;
    this.onReject = onReject;
    this.confidenceIndicator = confidenceIndicator;
    this.editedValues.clear();
    
    // Update header
    document.querySelector('.extraction-review-title').textContent = title;
    document.querySelector('.extraction-review-subtitle').textContent = subtitle;
    
    // Render content
    this.renderContent();
    this.updateStats();
    
    // Show modal
    const overlay = document.getElementById('extraction-review-overlay');
    overlay.classList.add('active');
    this.isOpen = true;
    
    // Focus first input
    setTimeout(() => {
      const firstInput = overlay.querySelector('.extraction-field-input');
      if (firstInput) firstInput.focus();
    }, 300);
  }

  /**
   * Render modal content
   */
  renderContent() {
    const content = document.getElementById('extraction-review-content');
    content.innerHTML = '';
    
    // Group data by section
    const sections = this.groupDataBySections(this.currentData);
    
    // Render confidence summary
    const confidenceSummary = this.createConfidenceSummary();
    content.appendChild(confidenceSummary);
    
    // Render each section
    for (const [sectionName, sectionData] of Object.entries(sections)) {
      const sectionElement = this.createSection(sectionName, sectionData);
      content.appendChild(sectionElement);
    }
  }

  /**
   * Group data by logical sections
   */
  groupDataBySections(data) {
    const sections = {
      'High-Level Parameters': {},
      'Deal Assumptions': {},
      'Revenue Items': {},
      'Cost Items': {},
      'Debt Model': {},
      'Exit Assumptions': {},
      'Other': {}
    };
    
    const fieldMapping = {
      // High-Level Parameters
      currency: 'High-Level Parameters',
      projectStartDate: 'High-Level Parameters',
      projectEndDate: 'High-Level Parameters',
      modelPeriods: 'High-Level Parameters',
      
      // Deal Assumptions
      dealName: 'Deal Assumptions',
      dealValue: 'Deal Assumptions',
      transactionFee: 'Deal Assumptions',
      dealLTV: 'Deal Assumptions',
      equityContribution: 'Deal Assumptions',
      debtFinancing: 'Deal Assumptions',
      
      // Revenue Items
      revenueItems: 'Revenue Items',
      totalRevenue: 'Revenue Items',
      revenueGrowthRate: 'Revenue Items',
      revenueCurrency: 'Revenue Items',
      
      // Cost Items
      operatingExpenses: 'Cost Items',
      capitalExpenses: 'Cost Items',
      totalOpEx: 'Cost Items',
      totalCapEx: 'Cost Items',
      costInflationRate: 'Cost Items',
      
      // Debt Model
      loanIssuanceFees: 'Debt Model',
      interestRateType: 'Debt Model',
      interestRate: 'Debt Model',
      baseRate: 'Debt Model',
      creditMargin: 'Debt Model',
      loanTerm: 'Debt Model',
      loanAmount: 'Debt Model',
      debtType: 'Debt Model',
      
      // Exit Assumptions
      disposalCost: 'Exit Assumptions',
      terminalCapRate: 'Exit Assumptions',
      exitMultiple: 'Exit Assumptions',
      exitMultipleType: 'Exit Assumptions',
      targetIRR: 'Exit Assumptions',
      expectedExitDate: 'Exit Assumptions',
      exitRoute: 'Exit Assumptions'
    };
    
    for (const [field, value] of Object.entries(data)) {
      if (field.startsWith('_')) continue; // Skip metadata
      
      const section = fieldMapping[field] || 'Other';
      sections[section][field] = value;
    }
    
    // Remove empty sections
    for (const [sectionName, sectionData] of Object.entries(sections)) {
      if (Object.keys(sectionData).length === 0) {
        delete sections[sectionName];
      }
    }
    
    return sections;
  }

  /**
   * Create confidence summary
   */
  createConfidenceSummary() {
    const summary = document.createElement('div');
    summary.className = 'extraction-confidence-summary';
    
    const stats = this.calculateOverallStats();
    
    summary.innerHTML = `
      <div class="extraction-overall-confidence">
        Overall Confidence: ${stats.averageConfidence}%
      </div>
      <div class="extraction-confidence-breakdown">
        <div class="extraction-confidence-item">
          <span style="color: #22c55e;">●</span> High: ${stats.high}
        </div>
        <div class="extraction-confidence-item">
          <span style="color: #f59e0b;">●</span> Medium: ${stats.medium}
        </div>
        <div class="extraction-confidence-item">
          <span style="color: #ef4444;">●</span> Low: ${stats.low}
        </div>
        <div class="extraction-confidence-item">
          <span style="color: #6b7280;">●</span> Empty: ${stats.empty}
        </div>
      </div>
    `;
    
    return summary;
  }

  /**
   * Create section element
   */
  createSection(sectionName, sectionData) {
    const section = document.createElement('div');
    section.className = 'extraction-review-section';
    
    const extractedCount = Object.values(sectionData).filter(v => v && v.value !== null && v.value !== undefined).length;
    const totalCount = Object.keys(sectionData).length;
    
    const header = document.createElement('div');
    header.className = 'extraction-section-header';
    header.innerHTML = `
      <span>${sectionName}</span>
      <span class="extraction-section-stats">${extractedCount}/${totalCount} fields extracted</span>
    `;
    
    const fields = document.createElement('div');
    fields.className = 'extraction-fields';
    
    for (const [fieldName, fieldData] of Object.entries(sectionData)) {
      const fieldElement = this.createField(fieldName, fieldData);
      fields.appendChild(fieldElement);
    }
    
    section.appendChild(header);
    section.appendChild(fields);
    
    return section;
  }

  /**
   * Create field element
   */
  createField(fieldName, fieldData) {
    const field = document.createElement('div');
    field.className = 'extraction-field';
    
    const label = document.createElement('div');
    label.className = 'extraction-field-label';
    label.textContent = this.formatFieldLabel(fieldName);
    
    const valueContainer = document.createElement('div');
    valueContainer.className = 'extraction-field-value';
    
    if (Array.isArray(fieldData?.value)) {
      // Handle array fields (revenue items, expenses)
      valueContainer.appendChild(this.createArrayInput(fieldName, fieldData));
    } else {
      // Handle single value fields
      valueContainer.appendChild(this.createSingleInput(fieldName, fieldData));
    }
    
    const actions = document.createElement('div');
    actions.className = 'extraction-field-actions';
    actions.appendChild(this.createFieldActions(fieldName, fieldData));
    
    field.appendChild(label);
    field.appendChild(valueContainer);
    field.appendChild(actions);
    
    return field;
  }

  /**
   * Create input for single value field
   */
  createSingleInput(fieldName, fieldData) {
    const container = document.createElement('div');
    container.style.flex = '1';
    
    const input = document.createElement('input');
    input.className = 'extraction-field-input';
    input.type = this.getInputType(fieldName);
    input.value = fieldData?.value || '';
    input.placeholder = fieldData?.value === null ? 'No data extracted' : '';
    input.dataset.field = fieldName;
    input.dataset.original = fieldData?.value || '';
    
    // Add confidence indicator
    if (this.confidenceIndicator && fieldData?.confidence !== undefined) {
      this.confidenceIndicator.addToField(input, fieldData, {
        position: 'after',
        showTooltip: true
      });
    }
    
    // Track changes
    input.addEventListener('input', () => {
      const originalValue = input.dataset.original;
      const currentValue = input.value;
      
      if (currentValue !== originalValue) {
        input.classList.add('edited');
        this.editedValues.set(fieldName, currentValue);
      } else {
        input.classList.remove('edited');
        this.editedValues.delete(fieldName);
      }
      
      this.updateStats();
    });
    
    container.appendChild(input);
    
    // Show original value if different
    if (fieldData?.source && fieldData.source !== 'not_found') {
      const original = document.createElement('div');
      original.className = 'extraction-field-original';
      original.textContent = `Source: ${this.formatSource(fieldData.source)}`;
      container.appendChild(original);
    }
    
    return container;
  }

  /**
   * Create input for array field
   */
  createArrayInput(fieldName, fieldData) {
    const container = document.createElement('div');
    container.style.flex = '1';
    
    const itemsContainer = document.createElement('div');
    itemsContainer.className = 'extraction-array-items';
    itemsContainer.dataset.field = fieldName;
    
    const items = fieldData?.value || [];
    
    items.forEach((item, index) => {
      const itemElement = this.createArrayItem(fieldName, item, index);
      itemsContainer.appendChild(itemElement);
    });
    
    // Add item button
    const addButton = document.createElement('div');
    addButton.className = 'extraction-add-item';
    addButton.textContent = '+ Add Item';
    addButton.addEventListener('click', () => {
      const newItem = this.createDefaultArrayItem(fieldName);
      const itemElement = this.createArrayItem(fieldName, newItem, items.length);
      itemsContainer.appendChild(itemElement);
      this.updateArrayData(fieldName);
    });
    
    container.appendChild(itemsContainer);
    container.appendChild(addButton);
    
    return container;
  }

  /**
   * Create array item element
   */
  createArrayItem(fieldName, item, index) {
    const itemDiv = document.createElement('div');
    itemDiv.className = 'extraction-array-item';
    
    // Name input
    const nameInput = document.createElement('input');
    nameInput.className = 'extraction-array-item-input';
    nameInput.placeholder = 'Item name';
    nameInput.value = item.name || '';
    
    // Value input
    const valueInput = document.createElement('input');
    valueInput.className = 'extraction-array-item-input';
    valueInput.type = 'number';
    valueInput.placeholder = 'Value';
    valueInput.value = item.value || '';
    
    // Actions
    const actions = document.createElement('div');
    actions.className = 'extraction-array-actions';
    
    const removeBtn = document.createElement('button');
    removeBtn.className = 'extraction-field-btn clear';
    removeBtn.innerHTML = '×';
    removeBtn.title = 'Remove item';
    removeBtn.addEventListener('click', () => {
      itemDiv.remove();
      this.updateArrayData(fieldName);
    });
    
    // Track changes
    [nameInput, valueInput].forEach(input => {
      input.addEventListener('input', () => {
        this.updateArrayData(fieldName);
      });
    });
    
    actions.appendChild(removeBtn);
    
    itemDiv.appendChild(nameInput);
    itemDiv.appendChild(valueInput);
    itemDiv.appendChild(actions);
    
    return itemDiv;
  }

  /**
   * Create field action buttons
   */
  createFieldActions(fieldName, fieldData) {
    const container = document.createElement('div');
    container.style.display = 'flex';
    container.style.gap = '4px';
    
    // Reset button
    const resetBtn = document.createElement('button');
    resetBtn.className = 'extraction-field-btn reset';
    resetBtn.innerHTML = '↺';
    resetBtn.title = 'Reset to original';
    resetBtn.addEventListener('click', () => {
      this.resetField(fieldName);
    });
    
    // Clear button
    const clearBtn = document.createElement('button');
    clearBtn.className = 'extraction-field-btn clear';
    clearBtn.innerHTML = '×';
    clearBtn.title = 'Clear value';
    clearBtn.addEventListener('click', () => {
      this.clearField(fieldName);
    });
    
    container.appendChild(resetBtn);
    container.appendChild(clearBtn);
    
    return container;
  }

  /**
   * Get appropriate input type for field
   */
  getInputType(fieldName) {
    const typeMap = {
      projectStartDate: 'date',
      projectEndDate: 'date',
      expectedExitDate: 'date',
      dealValue: 'number',
      transactionFee: 'number',
      dealLTV: 'number',
      equityContribution: 'number',
      debtFinancing: 'number',
      totalRevenue: 'number',
      totalOpEx: 'number',
      totalCapEx: 'number',
      interestRate: 'number',
      baseRate: 'number',
      creditMargin: 'number',
      loanIssuanceFees: 'number',
      disposalCost: 'number',
      terminalCapRate: 'number',
      targetIRR: 'number',
      exitMultiple: 'number',
      loanTerm: 'number',
      loanAmount: 'number'
    };
    
    return typeMap[fieldName] || 'text';
  }

  /**
   * Format field label for display
   */
  formatFieldLabel(fieldName) {
    const labelMap = {
      dealName: 'Deal Name',
      dealValue: 'Deal Value',
      transactionFee: 'Transaction Fee (%)',
      dealLTV: 'LTV (%)',
      equityContribution: 'Equity Contribution',
      debtFinancing: 'Debt Financing',
      currency: 'Currency',
      projectStartDate: 'Project Start Date',
      projectEndDate: 'Project End Date',
      modelPeriods: 'Model Periods',
      revenueItems: 'Revenue Items',
      totalRevenue: 'Total Revenue',
      revenueGrowthRate: 'Revenue Growth Rate (%)',
      operatingExpenses: 'Operating Expenses',
      capitalExpenses: 'Capital Expenses',
      totalOpEx: 'Total OpEx',
      totalCapEx: 'Total CapEx',
      costInflationRate: 'Cost Inflation Rate (%)',
      loanIssuanceFees: 'Loan Issuance Fees (%)',
      interestRateType: 'Interest Rate Type',
      interestRate: 'Interest Rate (%)',
      baseRate: 'Base Rate (%)',
      creditMargin: 'Credit Margin (%)',
      loanTerm: 'Loan Term (years)',
      loanAmount: 'Loan Amount',
      debtType: 'Debt Type',
      disposalCost: 'Disposal Cost (%)',
      terminalCapRate: 'Terminal Cap Rate (%)',
      exitMultiple: 'Exit Multiple',
      exitMultipleType: 'Exit Multiple Type',
      targetIRR: 'Target IRR (%)',
      expectedExitDate: 'Expected Exit Date',
      exitRoute: 'Exit Route'
    };
    
    return labelMap[fieldName] || fieldName.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase());
  }

  /**
   * Format source for display
   */
  formatSource(source) {
    const sourceMap = {
      'ai_extraction': 'AI Analysis',
      'pattern_matching': 'Pattern Recognition',
      'calculated': 'Calculated',
      'inferred': 'Inferred',
      'not_found': 'Not Found'
    };
    
    return sourceMap[source] || source.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
  }

  /**
   * Reset field to original value
   */
  resetField(fieldName) {
    const input = document.querySelector(`[data-field="${fieldName}"]`);
    if (input) {
      input.value = input.dataset.original || '';
      input.classList.remove('edited');
      this.editedValues.delete(fieldName);
      this.updateStats();
    }
  }

  /**
   * Clear field value
   */
  clearField(fieldName) {
    const input = document.querySelector(`[data-field="${fieldName}"]`);
    if (input) {
      input.value = '';
      input.classList.add('edited');
      this.editedValues.set(fieldName, '');
      this.updateStats();
    }
  }

  /**
   * Update array data from inputs
   */
  updateArrayData(fieldName) {
    const container = document.querySelector(`[data-field="${fieldName}"]`);
    const items = Array.from(container.children).map(itemDiv => {
      const inputs = itemDiv.querySelectorAll('.extraction-array-item-input');
      return {
        name: inputs[0]?.value || '',
        value: parseFloat(inputs[1]?.value) || 0
      };
    }).filter(item => item.name.trim() !== '');
    
    this.editedValues.set(fieldName, items);
    this.updateStats();
  }

  /**
   * Create default array item
   */
  createDefaultArrayItem(fieldName) {
    const defaults = {
      revenueItems: { name: 'New Revenue Stream', value: 0, growthType: 'linear', growthRate: 0 },
      operatingExpenses: { name: 'New Operating Expense', value: 0, category: 'other' },
      capitalExpenses: { name: 'New Capital Expense', value: 0, category: 'other' }
    };
    
    return defaults[fieldName] || { name: 'New Item', value: 0 };
  }

  /**
   * Calculate overall statistics
   */
  calculateOverallStats() {
    let totalFields = 0;
    let extractedFields = 0;
    let totalConfidence = 0;
    let high = 0, medium = 0, low = 0, empty = 0;
    
    for (const [key, value] of Object.entries(this.currentData)) {
      if (key.startsWith('_')) continue;
      
      totalFields++;
      
      if (value && value.value !== null && value.value !== undefined) {
        extractedFields++;
        totalConfidence += value.confidence || 0;
        
        const confidence = value.confidence || 0;
        if (confidence >= 0.8) high++;
        else if (confidence >= 0.5) medium++;
        else low++;
      } else {
        empty++;
      }
    }
    
    return {
      totalFields,
      extractedFields,
      averageConfidence: extractedFields > 0 ? Math.round((totalConfidence / extractedFields) * 100) : 0,
      high,
      medium,
      low,
      empty
    };
  }

  /**
   * Update statistics display
   */
  updateStats() {
    const stats = this.calculateOverallStats();
    const editedCount = this.editedValues.size;
    
    const statsElement = document.getElementById('extraction-review-stats');
    statsElement.innerHTML = `
      <div class="extraction-review-stat">
        <span>Fields Extracted:</span>
        <span>${stats.extractedFields}/${stats.totalFields}</span>
      </div>
      <div class="extraction-review-stat">
        <span>Average Confidence:</span>
        <span>${stats.averageConfidence}%</span>
      </div>
      <div class="extraction-review-stat">
        <span>Edited Fields:</span>
        <span>${editedCount}</span>
      </div>
    `;
  }

  /**
   * Apply changes and close modal
   */
  approve() {
    const finalData = this.getFinalData();
    
    if (this.onApprove) {
      this.onApprove(finalData);
    }
    
    this.close();
  }

  /**
   * Reject all changes and close modal
   */
  reject() {
    if (this.onReject) {
      this.onReject();
    }
    
    this.close();
  }

  /**
   * Get final data with user edits
   */
  getFinalData() {
    const finalData = { ...this.currentData };
    
    // Apply edited values
    for (const [fieldName, editedValue] of this.editedValues) {
      if (finalData[fieldName]) {
        finalData[fieldName] = {
          ...finalData[fieldName],
          value: editedValue,
          source: 'user_edited',
          confidence: 1.0
        };
      } else {
        finalData[fieldName] = {
          value: editedValue,
          source: 'user_added',
          confidence: 1.0
        };
      }
    }
    
    return finalData;
  }

  /**
   * Close the modal
   */
  close() {
    const overlay = document.getElementById('extraction-review-overlay');
    overlay.classList.remove('active');
    this.isOpen = false;
    this.currentData = null;
    this.editedValues.clear();
  }

  /**
   * Check if modal is currently open
   */
  isModalOpen() {
    return this.isOpen;
  }
}

// Export for use
window.ExtractionReviewModal = ExtractionReviewModal;