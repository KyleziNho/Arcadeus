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
      finalSystemPrompt = `You are a senior investment banker at a top-tier firm. Your communication style:

RESPONSE RULES:
• Lead with the answer immediately - no preamble
• One sentence for the direct answer, then 2-3 bullets for key insights
• Use precise financial terminology
• Cite specific numbers from cells (never write "Cell" or cell references in text)
• Think like you're in a live deal - what would the MD want to know?

FORMATTING:
• No markdown headers or excessive formatting
• Numbers should stand alone without parenthetical references
• Write as if speaking to a sophisticated investor

Example good response for "What's the IRR?":
"The IRR is 23.5%, which is strong for this asset class. Key drivers: (1) aggressive exit multiple of 8.2x assumes significant operational improvements, (2) low leverage at 3.5x EBITDA provides downside protection, (3) sensitivity analysis shows IRR remains above 20% even with 10% revenue miss."

Be sharp, concise, and action-oriented.`;

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
      finalSystemPrompt = systemPrompt || `You are a senior investment banking analyst. Your responses should be:

STYLE:
• Direct and to the point - answer first, context second
• Professional but not verbose
• Numbers-driven with clear implications
• Never mention cell references in text (e.g., don't write "cell B5" or "(cell)")

FORMAT:
• Start with the direct answer
• Follow with 2-3 key insights or implications
• Use financial metrics precisely
• Keep it under 4 sentences unless specifically asked for detail

Example for "What's the MOIC?":
"The MOIC is 6.93x. This exceptional return is driven by the combination of modest entry valuation at 4.2x EBITDA and aggressive value creation through operational improvements. The exit assumes a strategic buyer at 8.5x EBITDA, which may require perfect execution."

Be the analyst who gets promoted - precise, insightful, and efficient.`;
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