/**
 * Test AutoFill Flow - Debug Script
 * Run this in the browser console to test the autofill pipeline
 */

async function testAutoFillFlow() {
    console.log('🧪 Starting AutoFill Flow Test...');
    
    // Test 1: Check if components are loaded
    console.log('\n🔧 Test 1: Component Availability');
    const components = {
        'AutoFillIntegrator': window.autoFillIntegrator,
        'FileUploader': window.fileUploader,
        'FormHandler': window.formHandler,
        'AIExtractionService': window.AIExtractionService,
        'FieldMappingEngine': window.FieldMappingEngine
    };
    
    for (const [name, component] of Object.entries(components)) {
        console.log(`${component ? '✅' : '❌'} ${name}: ${component ? 'Available' : 'Missing'}`);
    }
    
    // Test 2: Test API connection
    console.log('\n🌐 Test 2: API Connection');
    try {
        const response = await fetch('/.netlify/functions/chat', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                message: 'Test connection',
                autoFillMode: false
            })
        });
        
        console.log(`${response.ok ? '✅' : '❌'} API Status: ${response.status}`);
        
        if (response.ok) {
            const data = await response.json();
            console.log('📄 API Response sample:', Object.keys(data));
        }
    } catch (error) {
        console.log(`❌ API Error: ${error.message}`);
    }
    
    // Test 3: Test with sample data
    console.log('\n📊 Test 3: Sample Data Extraction');
    try {
        const sampleFileContent = `
Company: TechCorp Inc.
Deal Value: $50,000,000
Transaction Fee: 2.5%
LTV: 75%

Revenue Streams:
- Software Licensing: $15,000,000 (3% annual growth)
- Support Services: $8,000,000 (2% annual growth)
- Professional Services: $5,000,000 (5% annual growth)

Operating Expenses:
- Staff Costs: $12,000,000 (4% annual growth)
- Marketing: $3,000,000 (2% annual growth)
- Office Rent: $1,200,000 (1% annual growth)

Capital Expenses:
- IT Equipment: $2,000,000 (no growth)
- Office Furniture: $500,000 (no growth)

Exit Assumptions:
- Disposal Cost: 2.0%
- Terminal Cap Rate: 8.5%
        `;
        
        const response = await fetch('/.netlify/functions/chat', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                message: 'Extract all financial data from this document',
                fileContents: [`File: sample-data.txt\nContent: ${sampleFileContent}`],
                autoFillMode: true,
                batchType: 'master_analysis'
            })
        });
        
        console.log(`${response.ok ? '✅' : '❌'} Sample Extraction Status: ${response.status}`);
        
        if (response.ok) {
            const data = await response.json();
            console.log('📈 Extraction Result:', data);
            
            // Test 4: Apply the data using FieldMappingEngine
            if (window.FieldMappingEngine && data.extractedData) {
                console.log('\n🗺️ Test 4: Field Mapping');
                try {
                    const fieldMappingEngine = new window.FieldMappingEngine();
                    const result = await fieldMappingEngine.applyDataToForm(data.extractedData);
                    console.log('✅ Field mapping completed:', result);
                } catch (mappingError) {
                    console.log('❌ Field mapping failed:', mappingError.message);
                }
            }
        } else {
            const errorText = await response.text();
            console.log('❌ API Error Response:', errorText);
        }
    } catch (error) {
        console.log(`❌ Sample Data Test Error: ${error.message}`);
    }
    
    // Test 5: Manual AutoFill Trigger (if available)
    console.log('\n🤖 Test 5: Manual AutoFill Trigger');
    if (window.autoFillIntegrator) {
        try {
            // Create a mock file
            const mockFile = {
                name: 'test-file.txt',
                content: 'Deal Value: $25,000,000\nRevenue: $10,000,000\nGrowth Rate: 5%',
                type: 'text/plain',
                size: 100
            };
            
            // Set mock files
            window.autoFillIntegrator.uploadedFiles = [mockFile];
            
            console.log('🎯 Triggering AutoFill with mock data...');
            await window.autoFillIntegrator.handleAutoFill();
            
        } catch (autoFillError) {
            console.log('❌ AutoFill test failed:', autoFillError.message);
        }
    } else {
        console.log('❌ AutoFillIntegrator not available');
    }
    
    console.log('\n🎉 Test Complete! Check the form to see if data was applied.');
}

// Auto-run if in browser console
if (typeof window !== 'undefined') {
    console.log('🧪 AutoFill Flow Test Script Loaded');
    console.log('💡 Run testAutoFillFlow() to start the test');
    window.testAutoFillFlow = testAutoFillFlow;
}