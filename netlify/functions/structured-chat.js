// Modern OpenAI Structured Outputs API with Streaming Support
// Using native fetch API for compatibility

exports.handler = async (event, context) => {
  // CORS headers
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Content-Type': 'application/json'
  };

  // Handle preflight requests
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  // Only allow POST requests
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers,
      body: JSON.stringify({ error: 'Method not allowed' })
    };
  }

  try {
    // Parse request body
    const requestData = JSON.parse(event.body);
    console.log('🎯 Structured chat request:', {
      hasMessage: !!requestData.message,
      hasSchema: !!requestData.schema,
      streaming: requestData.streaming,
      structuredOutput: requestData.structuredOutput
    });

    const { 
      message, 
      schema, 
      excelContext, 
      streaming = false, 
      structuredOutput = false,
      systemPrompt 
    } = requestData;

    // Validate inputs
    if (!message || typeof message !== 'string') {
      throw new Error('Message is required and must be a string');
    }

    // Get OpenAI API key
    const apiKey = process.env.OPENAI_API_KEY;
    if (!apiKey) {
      throw new Error('OpenAI API key not configured');
    }

    if (!apiKey.startsWith('sk-')) {
      throw new Error('Invalid OpenAI API key format');
    }

    // Prepare the messages array
    const messages = [
      {
        role: "system",
        content: systemPrompt || `You are an expert M&A financial analyst. Provide structured, actionable insights.

CRITICAL: If a JSON schema is provided, your response must perfectly match that schema structure. Use specific Excel cell references and provide concrete, data-driven insights.

Example structured response format:
- Always include summary, key metrics with locations, insights, and recommendations
- Reference specific Excel cells when mentioning calculations
- Provide actionable next steps for the user`
      },
      {
        role: "user",
        content: message + (excelContext ? `\n\nExcel Context:\n${JSON.stringify(excelContext, null, 2)}` : '')
      }
    ];

    // Prepare OpenAI request
    const openaiRequest = {
      model: "gpt-4o-2024-08-06", // Use latest model that supports structured outputs
      messages: messages,
      temperature: 0.3, // Lower temperature for more consistent structured output
      max_tokens: 4000
    };

    // Add structured output format if schema is provided
    if (structuredOutput && schema) {
      console.log('🏗️ Using structured output with schema');
      openaiRequest.response_format = {
        type: "json_schema",
        json_schema: {
          name: "ma_analysis_response",
          strict: true,
          schema: schema
        }
      };
    } else if (structuredOutput) {
      // Fallback to JSON mode if no specific schema
      openaiRequest.response_format = {
        type: "json_object"
      };
    }

    console.log('📡 Calling OpenAI API with structured output configuration...');

    // Handle streaming vs non-streaming
    if (streaming) {
      return await handleStreamingResponse(openaiRequest, apiKey, headers);
    } else {
      return await handleRegularResponse(openaiRequest, apiKey, headers);
    }

  } catch (error) {
    console.error('❌ Structured chat error:', error);
    
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({
        success: false,
        error: error.message || 'Internal server error',
        details: 'Check server logs for more information'
      })
    };
  }
};

/**
 * Handle regular (non-streaming) structured response
 */
async function handleRegularResponse(openaiRequest, apiKey, headers) {
  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify(openaiRequest)
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error('OpenAI API error:', errorText);
    throw new Error(`OpenAI API error: ${response.status} - ${errorText}`);
  }

  const responseData = await response.json();
  console.log('✅ OpenAI structured response received');

  // Check for refusals (OpenAI safety feature)
  const message = responseData.choices?.[0]?.message;
  if (message?.refusal) {
    console.warn('🛡️ OpenAI refused the request:', message.refusal);
    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({
        success: false,
        refusal: message.refusal,
        reason: 'The AI declined to process this request for safety reasons'
      })
    };
  }

  // Extract content
  const content = message?.content;
  if (!content) {
    throw new Error('No content in OpenAI response');
  }

  // Validate JSON if structured output was requested
  let parsedContent = content;
  if (openaiRequest.response_format?.type === "json_schema" || 
      openaiRequest.response_format?.type === "json_object") {
    try {
      parsedContent = JSON.parse(content);
      console.log('✅ Structured JSON response validated');
    } catch (parseError) {
      console.error('❌ Failed to parse structured JSON response:', parseError);
      console.error('Raw content:', content);
      throw new Error('Invalid JSON in structured response');
    }
  }

  return {
    statusCode: 200,
    headers,
    body: JSON.stringify({
      success: true,
      response: content,
      parsed: parsedContent,
      structured: !!openaiRequest.response_format,
      usage: responseData.usage,
      model: responseData.model
    })
  };
}

/**
 * Handle streaming structured response
 * Note: This is a placeholder for when OpenAI fully supports structured streaming
 */
async function handleStreamingResponse(openaiRequest, apiKey, headers) {
  console.log('🌊 Streaming structured response requested');
  
  // For now, use regular response and simulate streaming client-side
  // In the future, this would use OpenAI's streaming structured outputs
  const regularResponse = await handleRegularResponse(openaiRequest, apiKey, headers);
  
  // Add streaming indicator to response
  const responseBody = JSON.parse(regularResponse.body);
  responseBody.streaming = true;
  responseBody.note = 'Streaming simulation - real streaming will be implemented with OpenAI streaming structured outputs';
  
  return {
    ...regularResponse,
    body: JSON.stringify(responseBody)
  };
}