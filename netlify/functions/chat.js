// Import fetch for Node.js environment
const fetch = require('node-fetch');

exports.handler = async (event, context) => {
  console.log('🔥 FUNCTION CALLED! Method:', event.httpMethod);
  console.log('🔥 Event body:', event.body);
  console.log('🔥 Environment check - API Key exists:', !!process.env.OPENAI_API_KEY);
  console.log('🔥 Environment check - API Key length:', process.env.OPENAI_API_KEY ? process.env.OPENAI_API_KEY.length : 0);
  
  // Handle CORS preflight requests
  if (event.httpMethod === 'OPTIONS') {
    console.log('🔥 Handling OPTIONS request');
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
    console.log('🔥 Method not allowed:', event.httpMethod);
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
    console.log('🔥 Starting main function logic...');
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
    let { systemPrompt, temperature, maxTokens } = requestData;
    
    // Log the request structure for debugging
    console.log('📋 Request structure:', {
      hasMessage: !!message,
      messageType: typeof message,
      hasFileContents: !!fileContents,
      fileContentsType: typeof fileContents,
      fileContentsLength: Array.isArray(fileContents) ? fileContents.length : 'not array',
      autoFillMode: autoFillMode,
      batchType: batchType,
      hasSystemPrompt: !!systemPrompt,
      hasTemperature: !!temperature,
      hasMaxTokens: !!maxTokens
    });
    
    if (!message || typeof message !== 'string') {
      throw new Error('Message is required and must be a string');
    }

    // Get OpenAI API key from Netlify environment variables
    const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
    const hasOpenAIKey = OPENAI_API_KEY && OPENAI_API_KEY.length > 10;
    
    console.log('API Key check:', {
      hasKey: !!OPENAI_API_KEY,
      keyLength: OPENAI_API_KEY ? OPENAI_API_KEY.length : 0,
      keyPreview: OPENAI_API_KEY ? OPENAI_API_KEY.substring(0, 10) + '...' : 'NO KEY',
      hasOpenAIKey: hasOpenAIKey,
      batchType: batchType,
      autoFillMode: autoFillMode
    });
    
    let documentContext = '';
    if (fileContents && Array.isArray(fileContents) && fileContents.length > 0) {
      documentContext = `\n\nUploaded Documents:\n${fileContents.join('\n\n')}`;
    }

    // If we have OpenAI API key, use it for intelligent processing
    if (hasOpenAIKey) {
      console.log(`Using OpenAI API for processing - Batch Type: ${batchType || 'standard'}`);
      
      let finalSystemPrompt;
      let finalMaxTokens = 1500;
      let finalTemperature = autoFillMode ? 0.3 : 0.7;
      
      if (autoFillMode) {
        console.log('Auto-fill mode activated with batch processing');
        
        // Use custom parameters if provided by AIExtractionService, otherwise use batch-based logic
        if (systemPrompt && temperature !== undefined && maxTokens) {
          console.log('Using custom parameters from AIExtractionService');
          finalSystemPrompt = systemPrompt;
          finalMaxTokens = maxTokens;
          finalTemperature = temperature;
        } else if (batchType === 'basic') {
          finalMaxTokens = 2000; // Increased for comprehensive extraction
          finalSystemPrompt = `You are an expert financial analyst. Extract ALL financial data from the provided CSV/document.

EXTRACT AND RETURN:
{
  "extractedData": {
    "highLevelParameters": {
      "currency": "USD_from_Currency_field",
      "projectStartDate": "YYYY-MM-DD_from_Acquisition_date", 
      "projectEndDate": "calculate_from_start+holding_period",
      "modelPeriods": "monthly"
    },
    "dealAssumptions": {
      "dealName": "company_name_from_document",
      "dealValue": number_from_Equity+Debt_or_similar,
      "transactionFee": percentage_from_Transaction_Fees,
      "dealLTV": percentage_from_Acquisition_LTV
    },
    "revenueItems": [
      {"name": "Revenue Item 1", "initialValue": 500000, "growthType": "linear", "growthRate": 2}
    ],
    "operatingExpenses": [
      {"name": "Staff expenses", "initialValue": 60000, "growthType": "linear", "growthRate": 0.5}
    ],
    "capitalExpenses": [
      {"name": "Cost Item 3", "initialValue": 20000, "growthType": "linear", "growthRate": 1.5}
    ],
    "exitAssumptions": {
      "disposalCost": percentage_from_Disposal_Costs,
      "terminalCapRate": percentage_from_Terminal_Cap_Rate
    }
  }
}

ANALYZE: ${documentContext}`;
          
        } else if (batchType === 'revenue') {
          finalMaxTokens = 2000; // Medium for revenue items
          finalSystemPrompt = `You are an expert financial analyst AI. Extract revenue items and growth rates from the provided documents.

CRITICAL INSTRUCTIONS:
1. Analyze the actual document content provided below
2. Extract ONLY real revenue data found in the documents
3. Do NOT use placeholder or example values
4. If no revenue items found, return empty array []

REQUIRED JSON STRUCTURE:
{
  "extractedData": {
    "revenueItems": [
      {
        "name": "actual_revenue_stream_name",
        "initialValue": actual_number_value,
        "growthType": "linear|annual|custom",
        "growthRate": actual_rate_value
      }
    ]
  }
}

EXTRACTION RULES:
- Look for revenue streams, sales, income items
- Extract exact values and growth rates from the documents
- Use actual names and values found in the content
- Return empty array if no revenue data is found

ANALYZE THESE DOCUMENTS:
${documentContext}`;
          
        } else if (batchType === 'cost') {
          finalMaxTokens = 2000; // Medium for cost items
          finalSystemPrompt = `You are an expert financial analyst AI. Extract cost items and inflation rates from the provided documents.

CRITICAL INSTRUCTIONS:
1. Analyze the actual document content provided below
2. Extract ONLY real cost data found in the documents
3. Do NOT use placeholder or example values
4. If no cost items found, return empty array []

REQUIRED JSON STRUCTURE:
{
  "extractedData": {
    "costItems": [
      {
        "name": "actual_cost_item_name",
        "initialValue": actual_number_value,
        "growthType": "linear|annual|custom",
        "growthRate": actual_rate_value
      }
    ]
  }
}

EXTRACTION RULES:
- Look for operating expenses, staff costs, rent, marketing expenses
- Extract exact values and inflation rates from the documents
- Use actual names and values found in the content
- Return empty array if no cost data is found

ANALYZE THESE DOCUMENTS:
${documentContext}`;
          
        } else if (batchType === 'exit') {
          finalMaxTokens = 800; // Smaller for exit assumptions
          finalSystemPrompt = `You are an expert financial analyst AI. Extract disposal cost and terminal cap rate from the provided documents.

CRITICAL INSTRUCTIONS:
1. Analyze the actual document content provided below
2. Extract ONLY real data found in the documents
3. Do NOT use placeholder or example values
4. If values cannot be found, use null

REQUIRED JSON:
{
  "extractedData": {
    "exitAssumptions": {
      "disposalCost": actual_percentage_or_null,
      "terminalCapRate": actual_percentage_or_null
    }
  }
}

EXTRACTION RULES:
- Find: disposal cost, exit fees, terminal cap rate, exit yield
- Convert % to numbers: 2.5% → 2.5
- Use null if values not found in documents

ANALYZE THESE DOCUMENTS:
${documentContext}`;
          
        } else if (batchType === 'master_analysis') {
          finalMaxTokens = 4000; // Large for comprehensive analysis
          finalSystemPrompt = `You are an expert M&A analyst. Create a comprehensive, standardized data table from the provided documents for financial modeling.

ACT AS: Senior M&A analyst reviewing deal documents
TASK: Extract and organize ALL financial information systematically
GOAL: Create standardized data for P&L, FCF, and IRR calculations

ANALYZE FOR:
- Company details and business metrics
- Transaction structure and financing
- Revenue streams and growth patterns
- Operating and capital expenses
- Exit strategy and valuation assumptions
- Key financial ratios and projections

REQUIRED JSON STRUCTURE:
{
  "extractedData": {
    "standardizedData": {
      "companyOverview": {
        "companyName": "string",
        "industry": "string",
        "businessDescription": "string"
      },
      "transactionDetails": {
        "dealName": "string",
        "dealValue": number,
        "currency": "string",
        "transactionFees": number,
        "closingDate": "YYYY-MM-DD"
      },
      "financingStructure": {
        "debtLTV": number,
        "equityContribution": number,
        "debtFinancing": number
      },
      "historicalFinancials": {
        "revenueStreams": [{
          "name": "string",
          "currentValue": number,
          "growthRate": number
        }],
        "operatingExpenses": [{
          "name": "string",
          "currentValue": number,
          "inflationRate": number
        }],
        "capitalExpenses": [{
          "name": "string",
          "currentValue": number
        }]
      },
      "projectionAssumptions": {
        "reportingFrequency": "monthly|quarterly|annually",
        "projectionPeriod": "string"
      },
      "exitAssumptions": {
        "exitStrategy": "string",
        "disposalCosts": number,
        "terminalValue": number
      }
    }
  }
}

INSTRUCTIONS:
1. Extract numerical values, percentages, dates
2. Identify revenue streams and cost categories
3. Note financing terms and deal structure
4. Fill gaps with reasonable assumptions
5. Return ONLY the JSON structure

Document Content: ${documentContext}`;
          
        } else if (batchType === 'financial_analysis') {
          finalMaxTokens = 1500; // Adequate for detailed IRR/MOIC calculations
          finalSystemPrompt = `You are an expert M&A financial analyst specializing in IRR and MOIC calculations.

TASK: Create Excel formulas for investment return analysis based on actual cash flow data.

CRITICAL REQUIREMENTS:
1. IRR formulas must include initial investment as first cash flow (negative value)
2. Use proper Excel array syntax: =IRR({-investment;range_of_cash_flows})
3. Wrap in IFERROR to handle calculation errors
4. MOIC = Total Returns / Initial Investment
5. Use exact Excel sheet references provided in the prompt

EXCEL SYNTAX EXAMPLES:
- IRR with investment: =IFERROR(IRR({-12000000;'Free Cash Flow'!B21:BJ21}), "N/A")
- MOIC calculation: =SUM('Free Cash Flow'!B21:BJ21)/12000000

REQUIRED JSON OUTPUT:
{
  "calculations": {
    "leveredIRR": {
      "formula": "=IFERROR(IRR({-investment;cash_flow_range}), \"N/A\")",
      "description": "Levered IRR with initial investment"
    },
    "unleveredIRR": {
      "formula": "=IFERROR(IRR({-investment;cash_flow_range}), \"N/A\")",
      "description": "Unlevered IRR with initial investment"
    },
    "leveredMOIC": {
      "formula": "=SUM(cash_flow_range)/investment",
      "description": "Levered MOIC calculation"
    },
    "unleveredMOIC": {
      "formula": "=SUM(cash_flow_range)/investment", 
      "description": "Unlevered MOIC calculation"
    }
  }
}

Analyze the provided data and generate working Excel formulas.`;
          
        } else {
          // Fallback to original large prompt (legacy support)
          finalMaxTokens = 4000;
          finalSystemPrompt = `You are an expert financial analyst AI specialized in extracting data from M&A/PE documents and financial reports.

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
        finalSystemPrompt = `You are an AI assistant for Excel M&A modeling. Help with Excel commands and data analysis.`;
      }

      try {
        console.log('🤖 About to call OpenAI API...');
        console.log('🤖 Request parameters:', {
          model: 'gpt-4-turbo-preview',
          temperature: finalTemperature,
          max_tokens: finalMaxTokens,
          messageLength: message.length,
          systemPromptLength: finalSystemPrompt.length
        });
        
        const requestBody = {
          model: 'gpt-4-turbo-preview',
          messages: [
            { role: 'system', content: finalSystemPrompt },
            { role: 'user', content: message }
          ],
          response_format: { type: "json_object" },
          temperature: finalTemperature,
          max_tokens: finalMaxTokens
        };
        
        console.log('🤖 OpenAI request body size:', JSON.stringify(requestBody).length);
        
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${OPENAI_API_KEY}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(requestBody)
          // Note: fetch timeout is controlled by Netlify function timeout (up to 10 seconds for free tier)
        });
        
        console.log('🤖 OpenAI API response status:', response.status);

        if (!response.ok) {
          const errorText = await response.text();
          console.error('🤖 OpenAI API error:', response.status, errorText);
          
          // Return structured error response instead of throwing
          return {
            statusCode: 500,
            headers: {
              'Access-Control-Allow-Origin': '*',
              'Access-Control-Allow-Headers': 'Content-Type, Authorization',
              'Access-Control-Allow-Methods': 'POST, OPTIONS'
            },
            body: JSON.stringify({
              error: `OpenAI API error: ${response.status}`,
              response: "OpenAI service temporarily unavailable",
              details: errorText
            })
          };
        }

        console.log('🤖 Parsing OpenAI response...');
        const data = await response.json();
        console.log('🤖 OpenAI response received, choices length:', data.choices?.length);
        
        if (!data.choices || !data.choices[0] || !data.choices[0].message) {
          console.error('🤖 Invalid OpenAI response structure:', data);
          return {
            statusCode: 500,
            headers: {
              'Access-Control-Allow-Origin': '*',
              'Access-Control-Allow-Headers': 'Content-Type, Authorization',
              'Access-Control-Allow-Methods': 'POST, OPTIONS'
            },
            body: JSON.stringify({
              error: "Invalid response from OpenAI",
              response: "AI service returned invalid response",
              details: JSON.stringify(data)
            })
          };
        }
        
        const aiResponse = data.choices[0].message.content;
        console.log('🤖 AI response content length:', aiResponse?.length);
        
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
        console.error('🤖 OpenAI API call failed with error:', apiError.message);
        console.error('🤖 Full error details:', apiError);
        
        // Return structured error response instead of falling through
        return {
          statusCode: 500,
          headers: {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization',
            'Access-Control-Allow-Methods': 'POST, OPTIONS'
          },
          body: JSON.stringify({
            error: `API call failed: ${apiError.message}`,
            response: "AI service connection failed",
            type: "network_error"
          })
        };
      }
    }
    
    // Fallback logic when no API key or API call fails
    console.error('CRITICAL: AI extraction cannot proceed!');
    console.error('Reason: No OpenAI API key configured or API call failed');
    console.error('Has API key:', hasOpenAIKey);
    console.error('Auto-fill mode:', autoFillMode);
    
    if (autoFillMode) {
      const errorMessage = !hasOpenAIKey ? 
        'OpenAI API key not configured. Please set OPENAI_API_KEY environment variable in Netlify.' :
        'OpenAI API call failed. Please check service status.';
        
      return {
        statusCode: 500,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization',
          'Access-Control-Allow-Methods': 'POST, OPTIONS'
        },
        body: JSON.stringify({
          error: errorMessage,
          response: "AI extraction service is unavailable.",
          commands: []
        })
      };
    }
    
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
    console.error('🔥 CRITICAL ERROR in function:', error);
    console.error('🔥 Error message:', error.message);
    console.error('🔥 Error stack:', error.stack);
    console.error('🔥 Error type:', typeof error);
    console.error('🔥 Full error object:', JSON.stringify(error, null, 2));
    
    return {
      statusCode: 500,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization',
        'Access-Control-Allow-Methods': 'POST, OPTIONS'
      },
      body: JSON.stringify({ 
        response: `I encountered an error: ${error.message}. Please try again.`,
        error: true,
        debug: {
          message: error.message,
          stack: error.stack,
          timestamp: new Date().toISOString()
        }
      })
    };
  }
};// Force function rebuild Mon  9 Jun 2025 12:02:46 BST
