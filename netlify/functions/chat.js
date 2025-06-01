exports.handler = async (event, context) => {
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      body: JSON.stringify({ error: 'Method not allowed' })
    };
  }

  try {
    const { message, excelContext } = JSON.parse(event.body);
    
    const prompt = `You are an Excel M&A modeling expert that can directly execute Excel operations. The user says: "${message}"
    
Current Excel context: ${excelContext}

IMPORTANT: Always respond with valid JSON format.

If the user wants to generate a BLANK ASSUMPTIONS PAGE, use the "generateAssumptionsTemplate" action:
{
  "response": "I'll create a blank M&A assumptions template for you.",
  "commands": [
    {
      "action": "generateAssumptionsTemplate"
    }
  ]
}

If the user wants to FILL THE ASSUMPTIONS with sample data, use the "fillAssumptionsData" action with ALL data at once:
{
  "response": "I'll fill the assumptions template with sample M&A data.",
  "commands": [
    {
      "action": "fillAssumptionsData",
      "data": {
        "dealType": "Business Acquisition",
        "sector": "Technology",
        "geography": "United States",
        "businessModel": "SaaS",
        "ownership": "Private Equity",
        "acquisitionDate": "31/03/2025",
        "holdingPeriod": 60,
        "currency": "USD",
        "transactionFees": "1.5%",
        "acquisitionLTV": "75%",
        "equityContribution": 25000000,
        "debtFinancing": 75000000,
        "debtIssuanceFees": "1.0%",
        "interestRateMargin": "3.5%",
        "staffExpenses": 5000000,
        "salaryGrowth": "3.0%",
        "costItem1": 2000000,
        "costItem2": 800000,
        "costItem3": 1200000,
        "costItem4": 400000,
        "costItem5": 600000,
        "costItem6": 300000,
        "disposalCosts": "0.5%",
        "terminalEquityMultiple": 12.5,
        "terminalEBITDA": 15000000,
        "salePrice": 187500000
      }
    }
  ]
}

If the user wants to modify Excel data, respond with JSON containing both a text response AND executable commands:
{
  "response": "I'll add 2 to cell E2 for you.",
  "commands": [
    {
      "action": "addToCell",
      "cell": "E2", 
      "value": 2
    }
  ]
}

Available commands:
- setValue: Set cell to specific value {"action": "setValue", "cell": "A1", "value": 100}
- addToCell: Add value to existing cell {"action": "addToCell", "cell": "A1", "value": 5}
- setFormula: Set Excel formula {"action": "setFormula", "cell": "A1", "formula": "=SUM(B1:B10)"}
- formatCell: Format cell {"action": "formatCell", "cell": "A1", "format": {"bold": true, "color": "red"}}
- generateAssumptionsTemplate: Create blank M&A assumptions template {"action": "generateAssumptionsTemplate"}
- fillAssumptionsData: Fill entire assumptions template with sample data {"action": "fillAssumptionsData", "data": {"dealType": "Business Acquisition", "sector": "Technology", ...}}

For general advice without Excel modifications, respond with JSON containing just the response:
{
  "response": "IRR stands for Internal Rate of Return..."
}

Examples:
User: "add 2 to E2 cell" -> Return JSON with addToCell command
User: "set A1 to 100" -> Return JSON with setValue command  
User: "generate blank assumptions page" -> Return JSON with generateAssumptionsTemplate command
User: "fill assumptions with sample data" -> Return JSON with fillAssumptionsData command
User: "what is IRR?" -> Return JSON with just response text`;

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${process.env.OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: 'gpt-4-turbo',
        messages: [
          { role: 'system', content: 'You are an expert Excel M&A modeling assistant.' },
          { role: 'user', content: prompt }
        ],
        temperature: 0.3,
        max_tokens: 4000,
        response_format: { type: "json_object" }
      })
    });

    if (!response.ok) {
      console.error('OpenAI API error:', response.status, response.statusText);
      const errorText = await response.text();
      console.error('Error details:', errorText);
      throw new Error(`OpenAI API error: ${response.status}`);
    }

    const data = await response.json();
    const aiResponse = data.choices[0].message.content;
    
    // Parse JSON response (should always be JSON now)
    let parsedResponse;
    try {
      parsedResponse = JSON.parse(aiResponse);
      
      return {
        statusCode: 200,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Headers': 'Content-Type',
          'Access-Control-Allow-Methods': 'POST, OPTIONS'
        },
        body: JSON.stringify({
          response: parsedResponse.response,
          commands: parsedResponse.commands || []
        })
      };
    } catch (e) {
      console.error('JSON parse error:', e);
      // Fallback for non-JSON responses
      return {
        statusCode: 200,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Headers': 'Content-Type',
          'Access-Control-Allow-Methods': 'POST, OPTIONS'
        },
        body: JSON.stringify({ 
          response: aiResponse,
          commands: []
        })
      };
    }
  } catch (error) {
    console.error('Function error:', error);
    return {
      statusCode: 500,
      headers: {
        'Access-Control-Allow-Origin': '*'
      },
      body: JSON.stringify({ 
        error: 'Failed to process request',
        details: error.message 
      })
    };
  }
};