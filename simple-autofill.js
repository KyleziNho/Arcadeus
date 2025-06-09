// Simple AI Autofill - Uses Netlify function (no hardcoded API keys)
class SimpleAutofill {
  constructor() {
    // No API key needed - uses Netlify function with configured environment variables
  }

  async autofillFromFile(fileContent, fileName) {
    console.log('🚀 Starting simple autofill for:', fileName);
    
    try {
      // Call your Netlify function that has the API key configured
      const response = await fetch('/.netlify/functions/chat', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          message: "Extract financial data from uploaded file",
          fileContents: [`File: ${fileName}\nContent: ${fileContent}`],
          autoFillMode: true,
          batchType: 'basic'
        })
      });

      if (!response.ok) {
        throw new Error(`API error: ${response.status}`);
      }

      const data = await response.json();
      let extracted;
      
      // Handle different response formats from Netlify function
      if (data.extractedData) {
        extracted = data.extractedData;
      } else if (data.response) {
        try {
          extracted = JSON.parse(data.response);
        } catch (e) {
          throw new Error('Could not parse AI response');
        }
      } else {
        throw new Error('No data returned from AI');
      }
      
      console.log('✅ Extracted data:', extracted);
      
      // Apply directly to form fields
      this.applyToForm(extracted);
      
      return extracted;
      
    } catch (error) {
      console.error('❌ Autofill failed:', error);
      this.showNotification(`Autofill failed: ${error.message}`, 'error');
    }
  }

  applyToForm(data) {
    console.log('📝 Applying data to form...');
    
    // Handle nested data structure from Netlify function
    const highLevel = data.highLevelParameters || {};
    const dealData = data.dealAssumptions || {};
    
    // High Level Parameters
    if (data.currency || highLevel.currency) this.setField('currency', data.currency || highLevel.currency);
    if (data.projectStartDate || highLevel.projectStartDate) this.setField('project-start-date', data.projectStartDate || highLevel.projectStartDate);
    if (data.projectEndDate || highLevel.projectEndDate) this.setField('project-end-date', data.projectEndDate || highLevel.projectEndDate);
    
    // Deal Assumptions  
    if (data.dealName || dealData.dealName) this.setField('deal-name', data.dealName || dealData.dealName);
    if (data.dealValue || dealData.dealValue) this.setField('deal-value', data.dealValue || dealData.dealValue);
    if (data.transactionFee || dealData.transactionFee) this.setField('transaction-fee', data.transactionFee || dealData.transactionFee);
    if (data.dealLTV || dealData.dealLTV) this.setField('deal-ltv', data.dealLTV || dealData.dealLTV);
    if (data.equityContribution) this.setField('equity-contribution', data.equityContribution);
    if (data.debtFinancing) this.setField('debt-financing', data.debtFinancing);
    
    this.showNotification('✅ Autofill completed! Check the form fields.', 'success');
  }

  setField(fieldId, value) {
    const field = document.getElementById(fieldId);
    if (field && value !== null && value !== undefined) {
      field.value = value;
      console.log(`✅ Set ${fieldId} = ${value}`);
      
      // Trigger change event so the form knows it was updated
      field.dispatchEvent(new Event('change', { bubbles: true }));
      field.dispatchEvent(new Event('input', { bubbles: true }));
    } else {
      console.warn(`⚠️ Field not found or value is null: ${fieldId} = ${value}`);
    }
  }

  // Show notification without using alert (Excel add-in compatible)
  showNotification(message, type = 'info') {
    console.log(`📢 ${type.toUpperCase()}: ${message}`);
    
    // Create notification element
    const notification = document.createElement('div');
    notification.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      padding: 15px 20px;
      border-radius: 6px;
      color: white;
      font-size: 14px;
      font-weight: 500;
      z-index: 10000;
      max-width: 400px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.2);
      background: ${type === 'success' ? '#4CAF50' : type === 'error' ? '#f44336' : '#2196F3'};
    `;
    notification.textContent = message;
    
    document.body.appendChild(notification);
    
    // Remove after 5 seconds
    setTimeout(() => {
      if (notification.parentNode) {
        notification.parentNode.removeChild(notification);
      }
    }, 5000);
  }
}

// Initialize simple autofill when page loads
window.simpleAutofill = new SimpleAutofill();

// Add simple autofill button to the interface
function addSimpleAutofillButton() {
  // Look for the file upload area
  const uploadArea = document.querySelector('.file-upload-area') || 
                     document.querySelector('#file-upload') ||
                     document.querySelector('.file-drop-zone') ||
                     document.querySelector('[class*="upload"]');
  
  if (uploadArea) {
    const button = document.createElement('button');
    button.innerHTML = '🤖 Simple AI Autofill';
    button.type = 'button';
    button.style.cssText = `
      background: #4CAF50; 
      color: white; 
      padding: 12px 24px; 
      border: none; 
      border-radius: 6px; 
      margin: 10px 5px; 
      cursor: pointer; 
      font-size: 14px; 
      font-weight: 500;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      transition: background-color 0.2s;
    `;
    
    button.onmouseover = () => button.style.backgroundColor = '#45a049';
    button.onmouseout = () => button.style.backgroundColor = '#4CAF50';
    
    button.onclick = async () => {
      // Get uploaded files from the FileUploader system
      let files = [];
      
      // Try multiple ways to find uploaded files
      if (window.fileUploader && window.fileUploader.uploadedFiles) {
        files = window.fileUploader.uploadedFiles;
      } else if (window.formHandler && window.formHandler.fileUploader && window.formHandler.fileUploader.uploadedFiles) {
        files = window.formHandler.fileUploader.uploadedFiles;
      } else {
        // Fallback: check the main file input directly
        const fileInput = document.getElementById('mainFileInput');
        if (fileInput && fileInput.files && fileInput.files.length > 0) {
          files = Array.from(fileInput.files);
        }
      }
      
      if (!files || files.length === 0) {
        window.simpleAutofill.showNotification('Please upload a file first using the upload area above', 'error');
        return;
      }
      
      button.innerHTML = '⏳ Processing...';
      button.disabled = true;
      
      try {
        const file = files[0];
        let fileContent;
        
        // Check if file already has content processed
        if (file.content) {
          fileContent = file.content;
        } else if (file instanceof File) {
          fileContent = await file.text();
        } else {
          throw new Error('Unable to read file content');
        }
        
        console.log('🚀 Processing file:', file.name || 'unknown');
        await window.simpleAutofill.autofillFromFile(fileContent, file.name || 'uploaded-file');
      } catch (error) {
        console.error('Autofill error:', error);
        window.simpleAutofill.showNotification(`Error: ${error.message}`, 'error');
      } finally {
        button.innerHTML = '🤖 Simple AI Autofill';
        button.disabled = false;
      }
    };
    
    uploadArea.appendChild(button);
    console.log('✅ Simple autofill button added to page');
  } else {
    console.warn('⚠️ Could not find upload area to add autofill button');
    
    // Fallback: add to body
    setTimeout(() => {
      const container = document.querySelector('.container') || document.body;
      if (container && !document.querySelector('[data-simple-autofill]')) {
        const button = document.createElement('button');
        button.innerHTML = '🤖 Simple AI Autofill';
        button.setAttribute('data-simple-autofill', 'true');
        button.style.cssText = `
          position: fixed; 
          top: 20px; 
          right: 20px; 
          background: #4CAF50; 
          color: white; 
          padding: 12px 24px; 
          border: none; 
          border-radius: 6px; 
          cursor: pointer; 
          font-size: 14px; 
          z-index: 1000;
          box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        `;
        container.appendChild(button);
        console.log('✅ Simple autofill button added as fallback');
      }
    }, 2000);
  }
}

// Add button when page loads
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', addSimpleAutofillButton);
} else {
  addSimpleAutofillButton();
}