/**
 * ElicitationHandler.ts
 * Handles elicitation requests from MCP servers, allowing them to request user input
 */

export interface ElicitationRequest {
  method: 'elicitation/requestInput';
  params: {
    message: string;
    schema: {
      type: string;
      properties: Record<string, any>;
      required?: string[];
    };
    title?: string;
    priority?: 'low' | 'normal' | 'high';
    timeout?: number;
  };
}

export interface ElicitationResponse {
  data?: Record<string, any>;
  cancelled?: boolean;
  error?: string;
}

export class ElicitationHandler {
  private pendingElicitations: Map<string, ElicitationRequest> = new Map();
  private responseCallbacks: Map<string, (response: ElicitationResponse) => void> = new Map();

  /**
   * Handle an elicitation request from a server
   */
  async handleElicitationRequest(
    serverId: string,
    request: ElicitationRequest
  ): Promise<ElicitationResponse> {
    console.log(`📝 Elicitation request from server ${serverId}:`, request);

    const requestId = `elicit_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    
    try {
      // Create and show elicitation UI
      const response = await this.showElicitationUI(requestId, request);
      
      // Validate response against schema
      const validatedResponse = this.validateResponse(response, request.params.schema);
      
      return validatedResponse;
      
    } catch (error: any) {
      console.error('❌ Elicitation failed:', error);
      return {
        error: `Elicitation failed: ${error.message}`,
        cancelled: true
      };
    }
  }

  /**
   * Show elicitation UI to user
   */
  private async showElicitationUI(
    requestId: string,
    request: ElicitationRequest
  ): Promise<ElicitationResponse> {
    return new Promise((resolve) => {
      // Store callback for this request
      this.responseCallbacks.set(requestId, resolve);
      
      // Create and show UI
      const modal = this.createElicitationModal(requestId, request);
      document.body.appendChild(modal);
      
      // Auto-timeout if specified
      if (request.params.timeout) {
        setTimeout(() => {
          if (this.responseCallbacks.has(requestId)) {
            this.handleElicitationResponse(requestId, { 
              cancelled: true, 
              error: 'Request timed out' 
            });
          }
        }, request.params.timeout * 1000);
      }
    });
  }

  /**
   * Create elicitation modal UI
   */
  private createElicitationModal(
    requestId: string,
    request: ElicitationRequest
  ): HTMLElement {
    const modal = document.createElement('div');
    modal.className = 'elicitation-modal';
    modal.id = `elicitation-${requestId}`;
    
    const priorityIcon = {
      'low': '💡',
      'normal': '📝',
      'high': '⚠️'
    }[request.params.priority || 'normal'];
    
    modal.innerHTML = `
      <div class="elicitation-overlay">
        <div class="elicitation-content">
          <div class="elicitation-header">
            <h3>${priorityIcon} ${request.params.title || 'Information Needed'}</h3>
            <button class="close-btn" onclick="window.cancelElicitation('${requestId}')">&times;</button>
          </div>
          
          <div class="elicitation-body">
            <p class="elicitation-message">${this.escapeHtml(request.params.message)}</p>
            
            <form class="elicitation-form" id="form-${requestId}">
              ${this.generateFormFields(request.params.schema)}
            </form>
          </div>
          
          <div class="elicitation-actions">
            <button type="button" class="cancel-btn" onclick="window.cancelElicitation('${requestId}')">
              Cancel
            </button>
            <button type="button" class="submit-btn" onclick="window.submitElicitation('${requestId}')">
              Submit
            </button>
          </div>
        </div>
      </div>
    `;
    
    // Add global handlers
    (window as any).submitElicitation = this.submitElicitation.bind(this);
    (window as any).cancelElicitation = this.cancelElicitation.bind(this);
    
    return modal;
  }

  /**
   * Generate form fields based on schema
   */
  private generateFormFields(schema: any): string {
    if (!schema.properties) return '';
    
    return Object.entries(schema.properties).map(([key, field]: [string, any]) => {
      const isRequired = schema.required?.includes(key);
      const label = field.description || key;
      
      return `
        <div class="elicitation-field">
          <label for="field-${key}">
            ${label}${isRequired ? ' *' : ''}
          </label>
          ${this.generateInput(key, field, isRequired)}
        </div>
      `;
    }).join('');
  }

  /**
   * Generate input element based on field type
   */
  private generateInput(name: string, field: any, required: boolean): string {
    const commonAttrs = `
      id="field-${name}" 
      name="${name}" 
      ${required ? 'required' : ''}
      ${field.default !== undefined ? `value="${this.escapeHtml(String(field.default))}"` : ''}
    `;
    
    switch (field.type) {
      case 'string':
        if (field.enum) {
          return `
            <select ${commonAttrs}>
              <option value="">Select...</option>
              ${field.enum.map((option: string) => 
                `<option value="${this.escapeHtml(option)}">${this.escapeHtml(option)}</option>`
              ).join('')}
            </select>
          `;
        }
        return `<input type="text" ${commonAttrs} placeholder="${field.description || ''}">`;
        
      case 'number':
        return `<input type="number" ${commonAttrs} ${field.minimum ? `min="${field.minimum}"` : ''} ${field.maximum ? `max="${field.maximum}"` : ''}>`;
        
      case 'boolean':
        return `
          <div class="checkbox-container">
            <input type="checkbox" ${commonAttrs} ${field.default ? 'checked' : ''}>
            <span class="checkbox-label">${field.description || name}</span>
          </div>
        `;
        
      case 'array':
        return `<textarea ${commonAttrs} placeholder="Enter values separated by commas"></textarea>`;
        
      default:
        return `<input type="text" ${commonAttrs}>`;
    }
  }

  /**
   * Handle form submission
   */
  private submitElicitation(requestId: string): void {
    const form = document.getElementById(`form-${requestId}`) as HTMLFormElement;
    if (!form) return;
    
    const formData = new FormData(form);
    const data: Record<string, any> = {};
    
    // Collect form data
    for (const [key, value] of formData.entries()) {
      const field = form.querySelector(`[name="${key}"]`) as HTMLInputElement;
      
      if (field?.type === 'checkbox') {
        data[key] = field.checked;
      } else if (field?.type === 'number') {
        data[key] = value ? Number(value) : undefined;
      } else {
        data[key] = value || undefined;
      }
    }
    
    this.handleElicitationResponse(requestId, { data });
  }

  /**
   * Handle cancellation
   */
  private cancelElicitation(requestId: string): void {
    this.handleElicitationResponse(requestId, { cancelled: true });
  }

  /**
   * Handle elicitation response
   */
  private handleElicitationResponse(requestId: string, response: ElicitationResponse): void {
    const callback = this.responseCallbacks.get(requestId);
    if (callback) {
      callback(response);
      this.responseCallbacks.delete(requestId);
    }
    
    // Remove modal from DOM
    const modal = document.getElementById(`elicitation-${requestId}`);
    if (modal) {
      modal.remove();
    }
  }

  /**
   * Validate response against schema
   */
  private validateResponse(response: ElicitationResponse, schema: any): ElicitationResponse {
    if (response.cancelled || response.error) {
      return response;
    }
    
    if (!response.data) {
      return { error: 'No data provided', cancelled: true };
    }
    
    // Basic validation
    const errors: string[] = [];
    
    if (schema.required) {
      for (const requiredField of schema.required) {
        if (response.data[requiredField] === undefined || response.data[requiredField] === '') {
          errors.push(`${requiredField} is required`);
        }
      }
    }
    
    // Type validation
    for (const [key, value] of Object.entries(response.data)) {
      const fieldSchema = schema.properties?.[key];
      if (fieldSchema && value !== undefined) {
        const isValid = this.validateFieldValue(value, fieldSchema);
        if (!isValid) {
          errors.push(`${key} has invalid value`);
        }
      }
    }
    
    if (errors.length > 0) {
      return { 
        error: `Validation errors: ${errors.join(', ')}`,
        cancelled: true 
      };
    }
    
    return response;
  }

  /**
   * Validate individual field value
   */
  private validateFieldValue(value: any, schema: any): boolean {
    switch (schema.type) {
      case 'string':
        if (typeof value !== 'string') return false;
        if (schema.enum && !schema.enum.includes(value)) return false;
        break;
        
      case 'number':
        if (typeof value !== 'number' || isNaN(value)) return false;
        if (schema.minimum !== undefined && value < schema.minimum) return false;
        if (schema.maximum !== undefined && value > schema.maximum) return false;
        break;
        
      case 'boolean':
        if (typeof value !== 'boolean') return false;
        break;
        
      case 'array':
        if (!Array.isArray(value)) return false;
        break;
    }
    
    return true;
  }

  /**
   * Escape HTML for safe display
   */
  private escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }
}

export default ElicitationHandler;