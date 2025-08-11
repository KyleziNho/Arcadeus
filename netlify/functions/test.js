// Simple test function to verify OpenAI connectivity
exports.handler = async (event, context) => {
  console.log('Test function called');
  
  const apiKey = process.env.OPENAI_API_KEY;
  
  if (!apiKey) {
    return {
      statusCode: 500,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        success: false,
        error: 'OPENAI_API_KEY not found in environment',
        nodeVersion: process.version,
        fetchAvailable: typeof fetch !== 'undefined'
      })
    };
  }
  
  // Test with minimal OpenAI request
  try {
    const testData = {
      model: "gpt-4o-mini",
      messages: [
        { role: "user", content: "Hello" }
      ],
      max_tokens: 5
    };
    
    console.log('Testing OpenAI with minimal request...');
    
    // Using modern fetch API (with polyfill for older Node.js)
    if (typeof fetch === 'undefined') {
      try {
        global.fetch = require('node-fetch');
      } catch (error) {
        console.warn('node-fetch not available, using native implementation');
      }
    }
    
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify(testData)
    });
    
    console.log('OpenAI response status:', response.status);
    
    if (!response.ok) {
      const errorText = await response.text();
      console.error('OpenAI error:', errorText);
      return {
        statusCode: 200,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          success: false,
          error: `OpenAI API error: ${response.status}`,
          details: errorText,
          nodeVersion: process.version,
          fetchAvailable: typeof fetch !== 'undefined'
        })
      };
    }
    
    const data = await response.json();
    
    return {
      statusCode: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        success: true,
        message: 'OpenAI connection successful',
        response: data.choices?.[0]?.message?.content || 'No content',
        nodeVersion: process.version,
        fetchAvailable: typeof fetch !== 'undefined'
      })
    };
    
  } catch (error) {
    console.error('Test function error:', error);
    return {
      statusCode: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        success: false,
        error: error.message,
        stack: error.stack,
        nodeVersion: process.version,
        fetchAvailable: typeof fetch !== 'undefined'
      })
    };
  }
};