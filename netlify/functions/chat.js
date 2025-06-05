exports.handler = async (event, context) => {
  // Handle CORS preflight requests
  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization',
        'Access-Control-Allow-Methods': 'POST, OPTIONS'
      },
      body: ''
    };
  }

  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization',
        'Access-Control-Allow-Methods': 'POST, OPTIONS'
      },
      body: JSON.stringify({ error: 'Method not allowed' })
    };
  }

  try {
    // Validate request body
    if (!event.body) {
      throw new Error('Request body is required');
    }

    let requestData;
    try {
      requestData = JSON.parse(event.body);
    } catch (e) {
      throw new Error('Invalid JSON in request body');
    }

    const { message, excelContext, fileContents, autoFillMode, batchType } = requestData;
    
    if (!message || typeof message !== 'string') {
      throw new Error('Message is required and must be a string');
    }

    // Get OpenAI API key from Netlify environment variables
    const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
    const hasOpenAIKey = OPENAI_API_KEY && OPENAI_API_KEY.length > 10;
    
    console.log('API Key check:', {
      hasKey: !!OPENAI_API_KEY,
      keyLength: OPENAI_API_KEY ? OPENAI_API_KEY.length : 0,
      hasOpenAIKey: hasOpenAIKey,
      batchType: batchType
    });
    
    let documentContext = '';
    if (fileContents && Array.isArray(fileContents) && fileContents.length > 0) {
      documentContext = `\n\nUploaded Documents:\n${fileContents.join('\n\n')}`;
    }

    // If we have OpenAI API key, use it for intelligent processing
    if (hasOpenAIKey) {
      console.log(`Using OpenAI API for processing - Batch Type: ${batchType || 'standard'}`);
      
      let systemPrompt;
      let maxTokens = 1500;
      
      if (autoFillMode) {
        console.log('Auto-fill mode activated with batch processing');
        
        // Adjust token limits and system prompts based on batch type
        if (batchType === 'basic') {
          maxTokens = 1500; // Smaller for basic info
          systemPrompt = `You are an expert financial analyst AI. Extract ONLY high-level parameters and deal assumptions.

CRITICAL: Return ONLY the JSON structure below with the extracted data. Focus on basic deal information.

REQUIRED JSON STRUCTURE:
{
  "extractedData": {
    "highLevelParameters": {
      "currency": "USD",
      "projectStartDate": "2025-03-31",
      "projectEndDate": "2030-03-31",
      "modelPeriods": "monthly"
    },
    "dealAssumptions": {
      "dealName": "Company Name",
      "dealValue": 100000000,
      "transactionFee": 2.5,
      "dealLTV": 75
    }
  }
}

Document Content: ${documentContext}`;
          
        } else if (batchType === 'revenue') {
          maxTokens = 2000; // Medium for revenue items
          systemPrompt = `You are an expert financial analyst AI. Extract ONLY revenue items and growth rates.

CRITICAL: Return ONLY the JSON structure below with revenue data extracted from the documents.

REQUIRED JSON STRUCTURE:
{
  "extractedData": {
    "revenueItems": [
      {
        "name": "Revenue Stream 1",
        "initialValue": 500000,
        "growthType": "linear",
        "growthRate": 2
      }
    ]
  }
}

EXTRACTION RULES:
- Look for "Revenue Item 1", "Revenue Item 2", etc.
- Extract exact values and growth rates
- Match growth patterns to items by number
- If no revenue items found, return empty array []

Document Content: ${documentContext}`;
          
        } else if (batchType === 'cost') {
          maxTokens = 2000; // Medium for cost items
          systemPrompt = `You are an expert financial analyst AI. Extract ONLY cost items and inflation rates.

CRITICAL: Return ONLY the JSON structure below with cost data extracted from the documents.

REQUIRED JSON STRUCTURE:
{
  "extractedData": {
    "costItems": [
      {
        "name": "Staff expenses",
        "initialValue": 60000,
        "growthType": "linear",
        "growthRate": 0.5
      }
    ]
  }
}

EXTRACTION RULES:
- Look for "Cost Item 1", "Cost Item 2", "Staff expenses", operating expenses
- Extract exact values and inflation rates
- Match inflation to items by category
- If no cost items found, return empty array []

Document Content: ${documentContext}`;
          
        } else {
          // Fallback to original large prompt (legacy support)
          maxTokens = 4000;
          systemPrompt = `You are an expert financial analyst AI specialized in extracting data from M&A/PE documents and financial reports.

CRITICAL INSTRUCTIONS:
1. Analyze ALL content from uploaded files
2. Extract data matching the required fields below
3. Return ONLY valid JSON with extracted data
4. If data not found, use null for that field

REVENUE EXTRACTION:
- Look for "Revenue Item 1", "Revenue Item 2", etc.
- Extract exact values and growth rates
- Match growth patterns: "Rent Growth 1: 2%" → Revenue Item 1 has 2% linear growth
- Use exact names and values from files

COST EXTRACTION - MANDATORY:
- Look for "Cost Item 1", "Cost Item 2", "Staff expenses", etc.
- Extract exact values: "Cost Item 1,200000" → name: "Cost Item 1", initialValue: 200000
- Match inflation: "OpEx Cost Inflation,2" → apply 2% to OpEx items
- ALWAYS extract cost items if present - DO NOT SKIP
- Use same extraction logic as revenue items

REQUIRED DATA STRUCTURE:
{
  "extractedData": {
    "highLevelParameters": {
      "currency": "USD",
      "projectStartDate": "2025-03-31",
      "projectEndDate": "2027-03-31", 
      "modelPeriods": "monthly"
    },
    "dealAssumptions": {
      "dealName": "Company name",
      "dealValue": 100000000,
      "transactionFee": 2.5,
      "dealLTV": 75
    },
    "revenueItems": [
      {
        "name": "Revenue Item 1",
        "initialValue": 500000,
        "growthType": "linear",
        "growthRate": 2
      }
    ],
    "costItems": [
      {
        "name": "Staff expenses", 
        "initialValue": 60000,
        "growthType": "linear",
        "growthRate": 0.5
      },
      {
        "name": "Cost Item 1",
        "initialValue": 200000,
        "growthType": "linear",
        "growthRate": 2
      }
    ]
  }
}

Document Content to Analyze:
${documentContext}`;
        }
      } else {
        // Regular chat mode
        systemPrompt = `You are an AI assistant for Excel M&A modeling. Help with Excel commands and data analysis.`;
      }

      try {
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${OPENAI_API_KEY}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            model: 'gpt-4-turbo-preview',
            messages: [
              { role: 'system', content: systemPrompt },
              { role: 'user', content: message }
            ],
            response_format: { type: "json_object" },
            temperature: autoFillMode ? 0.3 : 0.7,
            max_tokens: maxTokens
          })
        });

        if (!response.ok) {
          const errorText = await response.text();
          console.error('OpenAI API error:', response.status, errorText);
          throw new Error(`OpenAI API error: ${response.status}`);
        }

        const data = await response.json();
        const aiResponse = data.choices[0].message.content;
        
        let parsedResponse;
        try {
          parsedResponse = JSON.parse(aiResponse);
        } catch (e) {
          console.error('Failed to parse AI response:', aiResponse);
          return {
            statusCode: 200,
            headers: {
              'Access-Control-Allow-Origin': '*',
              'Access-Control-Allow-Headers': 'Content-Type, Authorization',
              'Access-Control-Allow-Methods': 'POST, OPTIONS'
            },
            body: JSON.stringify({
              response: aiResponse,
              commands: [],
              error: 'Failed to parse response'
            })
          };
        }

        // Handle auto-fill mode response
        if (autoFillMode && parsedResponse.extractedData) {
          console.log('Returning extracted data for auto-fill');
          return {
            statusCode: 200,
            headers: {
              'Access-Control-Allow-Origin': '*',
              'Access-Control-Allow-Headers': 'Content-Type, Authorization',
              'Access-Control-Allow-Methods': 'POST, OPTIONS'
            },
            body: JSON.stringify({
              extractedData: parsedResponse.extractedData,
              response: "Data extracted successfully"
            })
          };
        }

        // Regular chat mode response
        return {
          statusCode: 200,
          headers: {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization',
            'Access-Control-Allow-Methods': 'POST, OPTIONS'
          },
          body: JSON.stringify({
            response: parsedResponse.response || "I've processed your request.",
            commands: parsedResponse.commands || []
          })
        };
        
      } catch (apiError) {
        console.error('OpenAI API call failed:', apiError);
        // Fall through to fallback logic below
      }
    }
    
    // Fallback logic when no API key or API call fails
    return {
      statusCode: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization',
        'Access-Control-Allow-Methods': 'POST, OPTIONS'
      },
      body: JSON.stringify({
        response: "I'm processing your request. Please try again.",
        commands: []
      })
    };

  } catch (error) {
    console.error('Function error:', error);
    
    return {
      statusCode: 500,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization',
        'Access-Control-Allow-Methods': 'POST, OPTIONS'
      },
      body: JSON.stringify({ 
        response: `I encountered an error: ${error.message}. Please try again.`,
        error: true
      })
    };
  }
};