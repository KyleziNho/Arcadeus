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

    // Determine system prompt based on batch type
    let finalSystemPrompt = systemPrompt;
    if (batchType === 'financial_analysis') {
      finalSystemPrompt = `You are a financial modeling expert specializing in M&A analysis. 
You create Excel formulas for IRR and MOIC calculations.
Return responses in JSON format with Excel formulas.
Focus on accuracy and proper financial modeling practices.`;
    } else if (batchType === 'chat') {
      finalSystemPrompt = systemPrompt || `You are an expert Excel and M&A financial modeling assistant. 
Provide clear, conversational responses about financial data and Excel analysis.
Give specific, data-driven insights in natural language.
Be helpful and analytical while maintaining a conversational tone.`;
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