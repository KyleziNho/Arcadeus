// Simple AI Document Reader Test
class SimpleAIReader {
  constructor() {
    console.log('🔍 Simple AI Reader initialized');
  }

  async testReadDocument() {
    console.log('🔍 Starting document read test...');
    
    try {
      // Step 1: Find uploaded files
      console.log('📁 Step 1: Looking for uploaded files...');
      
      let files = [];
      if (window.fileUploader && window.fileUploader.uploadedFiles) {
        files = window.fileUploader.uploadedFiles;
        console.log('✅ Found files in fileUploader:', files.length);
      } else {
        console.log('❌ No fileUploader found');
        return;
      }

      if (files.length === 0) {
        console.log('❌ No files uploaded');
        return;
      }

      // Step 2: Get file content
      console.log('📄 Step 2: Getting file content...');
      const file = files[0];
      console.log('📄 File name:', file.name);
      console.log('📄 File type:', file.type);
      console.log('📄 File size:', file.size);
      
      let content = '';
      if (file.content) {
        content = file.content;
        console.log('✅ File already has processed content');
      } else if (file instanceof File) {
        content = await file.text();
        console.log('✅ Read raw file content');
      }
      
      console.log('📄 Content preview (first 500 chars):', content.substring(0, 500));
      console.log('📄 Content length:', content.length);

      // Step 3: Call AI to analyze
      console.log('🤖 Step 3: Calling AI to analyze document...');
      
      const requestBody = {
        message: "List the key information you can find in this document",
        fileContents: [`File: ${file.name}\n\nContent:\n${content}`],
        autoFillMode: true,
        batchType: 'basic'
      };
      
      console.log('📤 Request body:', {
        message: requestBody.message,
        fileContentsLength: requestBody.fileContents[0].length,
        autoFillMode: requestBody.autoFillMode,
        batchType: requestBody.batchType
      });

      const response = await fetch('/.netlify/functions/chat', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody)
      });

      console.log('📥 Response status:', response.status);
      console.log('📥 Response ok:', response.ok);
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error('❌ API Error:', errorText);
        return;
      }

      const data = await response.json();
      console.log('📥 Response data:', data);

      // Step 4: Show what AI found
      console.log('🎯 Step 4: AI Analysis Results:');
      
      if (data.extractedData) {
        console.log('✅ AI extracted data:', JSON.stringify(data.extractedData, null, 2));
        
        // Log each field found
        if (data.extractedData.highLevelParameters) {
          console.log('📊 High Level Parameters:', data.extractedData.highLevelParameters);
        }
        if (data.extractedData.dealAssumptions) {
          console.log('💰 Deal Assumptions:', data.extractedData.dealAssumptions);
        }
      } else if (data.response) {
        console.log('📝 AI Response:', data.response);
      }

      console.log('✅ Test complete!');
      
    } catch (error) {
      console.error('❌ Test failed:', error);
      console.error('❌ Error stack:', error.stack);
    }
  }
}

// Initialize and add test button
window.simpleAIReader = new SimpleAIReader();

// Add test button to page
function addTestButton() {
  const button = document.createElement('button');
  button.innerHTML = '🧪 Test AI Read';
  button.style.cssText = `
    position: fixed;
    bottom: 20px;
    right: 20px;
    background: #2196F3;
    color: white;
    padding: 10px 20px;
    border: none;
    border-radius: 5px;
    cursor: pointer;
    z-index: 1000;
  `;
  
  button.onclick = () => {
    console.clear();
    console.log('========== AI READ TEST ==========');
    window.simpleAIReader.testReadDocument();
  };
  
  document.body.appendChild(button);
  console.log('🧪 Test button added - look for blue button in bottom right');
}

// Add button when ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', addTestButton);
} else {
  addTestButton();
}