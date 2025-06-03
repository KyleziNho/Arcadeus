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

    // Check if OpenAI API key is configured
    const hasOpenAIKey = process.env.OPENAI_API_KEY && process.env.OPENAI_API_KEY.length > 10;
    
    let documentContext = '';
    if (fileContents && Array.isArray(fileContents) && fileContents.length > 0) {
      documentContext = `\n\nUploaded Documents:\n${fileContents.join('\n\n')}`;
    }

    // If no OpenAI API key, use fallback logic
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

    // If we have OpenAI API key, use it (this part would be for when API key is available)
    // For now, just return a message indicating API is needed
    return {
      statusCode: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization',
        'Access-Control-Allow-Methods': 'POST, OPTIONS'
      },
      body: JSON.stringify({
        response: "AI processing requires OpenAI API key configuration.",
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