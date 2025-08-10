const https = require('https');

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
    const { message, autoFillMode, batchType, systemPrompt, temperature, maxTokens } = JSON.parse(event.body);
    
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

    // Prepare OpenAI request
    const openaiData = {
      model: "gpt-3.5-turbo",
      messages: [
        {
          role: "system",
          content: finalSystemPrompt || "You are a helpful assistant that provides financial analysis and Excel formulas."
        },
        {
          role: "user",
          content: message
        }
      ],
      temperature: temperature || 0.7,
      max_tokens: maxTokens || 3000
    };

    console.log('🤖 Calling OpenAI API...');

    // Make request to OpenAI
    const response = await new Promise((resolve, reject) => {
      const data = JSON.stringify(openaiData);
      
      const options = {
        hostname: 'api.openai.com',
        port: 443,
        path: '/v1/chat/completions',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${apiKey}`,
          'Content-Length': data.length
        }
      };

      const req = https.request(options, (res) => {
        let responseData = '';
        
        res.on('data', (chunk) => {
          responseData += chunk;
        });
        
        res.on('end', () => {
          try {
            const parsedResponse = JSON.parse(responseData);
            resolve(parsedResponse);
          } catch (error) {
            reject(new Error(`Failed to parse OpenAI response: ${error.message}`));
          }
        });
      });

      req.on('error', (error) => {
        reject(error);
      });

      req.write(data);
      req.end();
    });

    console.log('✅ OpenAI API response received');

    // Debug the OpenAI response
    console.log('OpenAI response structure:', JSON.stringify(response, null, 2));
    
    // Extract the response content
    const content = response.choices?.[0]?.message?.content;
    if (!content) {
      // Provide detailed error info
      const errorDetails = {
        hasChoices: !!response.choices,
        choicesLength: response.choices?.length,
        firstChoice: response.choices?.[0],
        errorFromAPI: response.error
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
        content: content,
        usage: response.usage
      })
    };

  } catch (error) {
    console.error('❌ Chat function error:', error);
    
    // Provide an intelligent fallback response for MOIC questions
    const errorMessage = error.message || 'Internal server error';
    let fallbackResponse = null;
    
    try {
      const { message } = JSON.parse(event.body);
      if (message && message.toLowerCase().includes('moic')) {
        fallbackResponse = `I can help with MOIC (Multiple on Invested Capital) analysis! 

A high MOIC typically indicates:
• Strong cash generation relative to initial investment
• Successful value creation strategies  
• Potential over-leveraging if debt was used
• Market timing or sector performance benefits

To analyze your specific MOIC:
1. Check your cash flow projections for reasonableness
2. Verify your exit assumptions (terminal value, multiples)
3. Compare to industry benchmarks (typically 2.0-5.0x for PE)
4. Consider sensitivity analysis on key drivers

Would you like me to help you analyze specific components of your MOIC calculation?

*Note: AI service temporarily unavailable, but I can still provide financial modeling guidance.*`;
      }
    } catch (parseError) {
      // Ignore parsing errors for fallback
    }
    
    return {
      statusCode: fallbackResponse ? 200 : 500,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(fallbackResponse ? {
        success: true,
        content: fallbackResponse,
        fallback: true
      } : {
        success: false,
        error: errorMessage
      })
    };
  }
};