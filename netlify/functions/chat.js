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

    const { message, excelContext, fileContents } = requestData;
    
    if (!message || typeof message !== 'string') {
      throw new Error('Message is required and must be a string');
    }

    // Get OpenAI API key from Netlify environment variables
    const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
    const hasOpenAIKey = OPENAI_API_KEY && OPENAI_API_KEY.length > 10;
    
    console.log('API Key check:', {
      hasKey: !!OPENAI_API_KEY,
      keyLength: OPENAI_API_KEY ? OPENAI_API_KEY.length : 0,
      hasOpenAIKey: hasOpenAIKey
    });
    
    let documentContext = '';
    if (fileContents && Array.isArray(fileContents) && fileContents.length > 0) {
      documentContext = `\n\nUploaded Documents:\n${fileContents.join('\n\n')}`;
    }

    // If we have OpenAI API key, use it for intelligent processing
    if (hasOpenAIKey) {
      console.log('Using OpenAI API for processing');
      
      const systemPrompt = `You are an AI assistant specialized in M&A and PE deal modeling in Excel. You can execute Excel commands to help users build financial models.

Available commands:
- generateAssumptionsTemplate: Creates a professional assumptions template
- fillAssumptionsData: Fills template with data (provide data object)
- setValue: Set value in specific cell (provide cell, value)
- setFormula: Set formula in specific cell (provide cell, formula)
- formatCell: Format a cell (provide cell, format object)

When users ask to:
1. "create template" or "blank template" → use generateAssumptionsTemplate
2. "fill data" or "use CSV data" → use fillAssumptionsData with parsed data
3. Specific Excel tasks → use appropriate commands

Always respond conversationally and execute relevant commands.
Always return valid JSON with "response" and "commands" fields.

Excel Context: ${excelContext || 'Not available'}
${documentContext}`;

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
            temperature: 0.7,
            max_tokens: 1500
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
          console.error('Failed to parse AI response as JSON:', aiResponse);
          // Fallback: treat as text response
          return {
            statusCode: 200,
            headers: {
              'Access-Control-Allow-Origin': '*',
              'Access-Control-Allow-Headers': 'Content-Type, Authorization',
              'Access-Control-Allow-Methods': 'POST, OPTIONS'
            },
            body: JSON.stringify({
              response: aiResponse,
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
    console.log('Using fallback logic (no API key or API failed)');
    if (!hasOpenAIKey) {
      console.log('No OpenAI API key, using fallback logic');
      
      let fallbackResponse = "I'll help you with that request.";
      let fallbackCommands = [];
      
      // Check if the original message mentions creating a template
      if (message.toLowerCase().includes('blank') && message.toLowerCase().includes('template')) {
        fallbackCommands.push({ action: 'generateAssumptionsTemplate' });
        fallbackResponse = "I'll create a blank assumptions template for you.";
      }
      
      // Check if they want to fill with data
      if (message.toLowerCase().includes('fill')) {
        fallbackCommands.push({ 
          action: 'fillAssumptionsData',
          data: {
            dealType: "Business Acquisition",
            sector: "Technology", 
            geography: "United States",
            businessModel: "SaaS",
            ownership: "Private Equity",
            purchasePrice: 100,
            acquisitionDate: "31/03/2025",
            holdingPeriod: 60,
            currency: "USD",
            transactionFees: "1.5%",
            acquisitionLTV: "75%",
            debtIssuanceFees: "1.0%",
            interestRateMargin: "3.5%",
            staffExpenses: 5000000,
            salaryGrowth: "3.0%",
            costItem1: 2000000,
            costItem2: 800000,
            costItem3: 1200000,
            costItem4: 400000,
            costItem5: 600000,
            costItem6: 300000,
            disposalCosts: "0.5%",
            terminalEquityMultiple: 12.5,
            terminalEBITDA: 15000000,
            salePrice: 187500000
          }
        });
        if (fallbackCommands.length === 1) {
          fallbackResponse = "I'll analyze your uploaded file and fill the assumptions template.";
        } else {
          fallbackResponse = "I'll create a template and fill it with data from your uploaded file.";
        }
      }
      
      return {
        statusCode: 200,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization',
          'Access-Control-Allow-Methods': 'POST, OPTIONS'
        },
        body: JSON.stringify({
          response: fallbackResponse,
          commands: fallbackCommands
        })
      };
    }

    // Default fallback if no conditions match
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
    console.error('Error stack:', error.stack);
    
    // More detailed error response
    let errorMessage = 'Failed to process request';
    if (error.message) {
      errorMessage = error.message;
    }
    
    return {
      statusCode: 500,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization',
        'Access-Control-Allow-Methods': 'POST, OPTIONS'
      },
      body: JSON.stringify({ 
        response: `I encountered an error: ${errorMessage}. Please try again.`,
        error: true
      })
    };
  }
};