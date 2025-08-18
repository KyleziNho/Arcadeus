// Using modern fetch API (with polyfill for older Node.js)
if (typeof fetch === 'undefined') {
  try {
    global.fetch = require('node-fetch');
  } catch (error) {
    console.warn('node-fetch not available, using native implementation');
  }
}

exports.handler = async (event, context) => {
  // Only allow POST requests
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Access-Control-Allow-Methods': 'POST, OPTIONS'
      },
      body: JSON.stringify({ error: 'Method not allowed' })
    };
  }

  // Handle CORS preflight
  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Access-Control-Allow-Methods': 'POST, OPTIONS'
      },
      body: ''
    };
  }

  try {
    // Parse request body
    let requestData;
    try {
      requestData = JSON.parse(event.body);
    } catch (parseError) {
      console.error('Failed to parse request body:', event.body);
      throw new Error(`Invalid request body: ${parseError.message}`);
    }
    
    // Handle different request formats
    let message, autoFillMode, batchType, systemPrompt, temperature, maxTokens, context;
    
    // Check if this is a direct message (simple format)
    if (typeof requestData.message === 'string') {
      message = requestData.message;
      autoFillMode = requestData.autoFillMode;
      batchType = requestData.batchType;
      systemPrompt = requestData.systemPrompt;
      temperature = requestData.temperature;
      maxTokens = requestData.maxTokens;
      context = requestData.context;
    } else {
      // Handle ChatHandler format where everything might be nested
      console.log('Request structure:', Object.keys(requestData));
      message = requestData.message || 'No message found';
      autoFillMode = requestData.autoFillMode;
      batchType = requestData.batchType || 'chat';
      systemPrompt = requestData.systemPrompt;
      temperature = requestData.temperature;
      maxTokens = requestData.maxTokens;
      context = requestData;
    }
    
    console.log('Extracted message:', message);
    console.log('Message type:', typeof message);
    
    console.log('📝 Chat function called with:', {
      messageLength: message?.length,
      autoFillMode,
      batchType,
      temperature,
      maxTokens
    });

    // Get OpenAI API key from environment
    const apiKey = process.env.OPENAI_API_KEY;
    if (!apiKey) {
      throw new Error('OpenAI API key not configured');
    }
    
    // Validate API key format
    if (!apiKey.startsWith('sk-')) {
      throw new Error('Invalid OpenAI API key format (should start with sk-)');
    }
    
    console.log('API key found, length:', apiKey.length, 'starts with:', apiKey.substring(0, 7) + '...');

    // Determine system prompt based on batch type and query analysis (Hebbia-style routing)
    let finalSystemPrompt = systemPrompt;
    
    if (batchType === 'mcp_function_calling') {
      // Handle MCP function calling - message contains the full OpenAI request
      try {
        const mcpRequest = JSON.parse(message);
        
        console.log('🔧 MCP Function Calling Request:', mcpRequest);
        
        const openaiData = {
          model: mcpRequest.model || "gpt-4o-mini",
          messages: mcpRequest.messages,
          functions: mcpRequest.functions,
          function_call: mcpRequest.function_call || 'auto',
          temperature: mcpRequest.temperature || 0.1,
          max_tokens: maxTokens || 3000
        };

        console.log('🤖 Calling OpenAI API with MCP function calling data:', JSON.stringify(openaiData, null, 2));

        // Make request to OpenAI using modern fetch API
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`
          },
          body: JSON.stringify(openaiData)
        });

        console.log('📡 OpenAI API response status:', response.status);
        
        if (!response.ok) {
          const errorText = await response.text();
          console.error('OpenAI API error response:', errorText);
          throw new Error(`OpenAI API error: ${response.status} - ${errorText}`);
        }

        const responseData = await response.json();
        console.log('✅ OpenAI MCP API response received');

        // Return the raw OpenAI response for MCP processing
        return {
          statusCode: 200,
          headers: {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Headers': 'Content-Type',
            'Access-Control-Allow-Methods': 'POST, OPTIONS',
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            success: true,
            response: JSON.stringify(responseData), // Return full OpenAI response as string
            usage: responseData.usage
          })
        };
        
      } catch (parseError) {
        console.error('❌ Failed to parse MCP request:', parseError);
        throw new Error(`Invalid MCP request format: ${parseError.message}`);
      }
      
    } else if (batchType === 'financial_analysis') {
      finalSystemPrompt = `You are analyzing an Excel financial model. The user will provide Excel data in this format:
"Sheet!Address: Label = Value"

RESPONSE RULES:
1. Extract the EXACT cell reference from the Excel data provided
2. Format: "The [metric] is [value] (exact cell reference from data)"
3. If you see "Sheet1!B23: MOIC = 17.94", respond with "MOIC is 17.94x (B23)"
4. Always use the cell address shown in the Excel data
5. If no cell reference is provided, respond without parentheses

Examples:
Input: "Sheet1!B47: IRR = 0.235"
Output: "IRR is 23.5% (B47)."

Input: "Assumptions!B15: Exit Multiple = 8.2"  
Output: "Exit multiple is 8.2x (Assumptions!B15)."

Be extremely concise. Maximum 2 sentences.`;

    } else if (requestData.queryType === 'excel_structure') {
      finalSystemPrompt = `You are a specialized Excel formula and structure analysis agent. Your expertise:

• Analyze Excel formulas and calculation logic
• Explain complex formula relationships
• Identify calculation dependencies and potential errors
• Suggest formula optimizations and best practices
• Provide specific cell references and range explanations

Focus on technical accuracy and clear explanations of Excel mechanics.`;

    } else if (requestData.queryType === 'data_validation') {
      finalSystemPrompt = `You are a specialized data validation and error detection agent. Your expertise:

• Identify data inconsistencies and calculation errors
• Validate financial model logic and assumptions
• Check for missing critical inputs
• Highlight potential red flags in the model
• Provide specific, actionable fixes

Be thorough, precise, and focus on what needs to be corrected immediately.`;

    } else {
      finalSystemPrompt = systemPrompt || `You are analyzing an Excel model. Response rules:

1. Answer in ONE sentence with the value and cell reference
2. Format: "[Answer] [value] (cell)"
3. Include cell references for ALL numbers mentioned
4. Maximum 2 sentences total

Example:
Q: "What's the equity contribution?"
A: "Equity contribution is $57.5M (B12)."

Q: "Explain the returns"
A: "IRR is 23.5% (B47) with MOIC of 6.93x (B48). Driven by entry at 4.2x EBITDA (B8) and exit at 8.5x (B32)."

Be extremely brief.`;
    }

    // Ensure message is a string
    if (!message || typeof message !== 'string') {
      throw new Error('Message is required and must be a string');
    }

    // Prepare OpenAI request with modern model
    const openaiData = {
      model: "gpt-4o-mini", // Use current model instead of deprecated gpt-3.5-turbo
      messages: [
        {
          role: "system",
          content: finalSystemPrompt || "You are a helpful assistant that provides financial analysis and Excel formulas."
        },
        {
          role: "user",
          content: String(message) // Ensure it's a string
        }
      ],
      temperature: temperature || 0.7,
      max_tokens: maxTokens || 3000
    };

    console.log('🤖 Calling OpenAI API with data:', JSON.stringify(openaiData, null, 2));

    // Make request to OpenAI using modern fetch API
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify(openaiData)
    });

    console.log('📡 OpenAI API response status:', response.status);
    
    if (!response.ok) {
      const errorText = await response.text();
      console.error('OpenAI API error response:', errorText);
      throw new Error(`OpenAI API error: ${response.status} - ${errorText}`);
    }

    const responseData = await response.json();

    console.log('✅ OpenAI API response received');

    // Debug the OpenAI response
    console.log('OpenAI response structure:', JSON.stringify(responseData, null, 2));
    
    // Extract the response content
    const content = responseData.choices?.[0]?.message?.content;
    if (!content) {
      // Provide detailed error info
      const errorDetails = {
        hasChoices: !!responseData.choices,
        choicesLength: responseData.choices?.length,
        firstChoice: responseData.choices?.[0],
        errorFromAPI: responseData.error
      };
      console.error('OpenAI response analysis:', errorDetails);
      throw new Error(`No content in OpenAI response. Details: ${JSON.stringify(errorDetails)}`);
    }

    // Return the response
    return {
      statusCode: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        success: true,
        response: content, // ChatHandler expects 'response' field
        content: content,  // Keep both for compatibility
        usage: responseData.usage
      })
    };

  } catch (error) {
    console.error('❌ Chat function error:', error);
    
    return {
      statusCode: 500,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        success: false,
        error: error.message || 'Internal server error'
      })
    };
  }
};