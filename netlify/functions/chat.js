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

For general advice without Excel modifications, respond with just text (no JSON).

Examples:
User: "add 2 to E2 cell" -> Return JSON with addToCell command
User: "set A1 to 100" -> Return JSON with setValue command  
User: "what is IRR?" -> Return plain text explanation`;

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${process.env.OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: 'gpt-4',
        messages: [
          { role: 'system', content: 'You are an expert Excel M&A modeling assistant.' },
          { role: 'user', content: prompt }
        ],
        temperature: 0.3,
        max_tokens: 1500
      })
    });

    const data = await response.json();
    const aiResponse = data.choices[0].message.content;
    
    // Try to parse as JSON first (for commands)
    let parsedResponse;
    try {
      parsedResponse = JSON.parse(aiResponse);
      // If it's valid JSON with commands, return it structured
      if (parsedResponse.response && parsedResponse.commands) {
        return {
          statusCode: 200,
          headers: {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Headers': 'Content-Type',
            'Access-Control-Allow-Methods': 'POST, OPTIONS'
          },
          body: JSON.stringify({
            response: parsedResponse.response,
            commands: parsedResponse.commands
          })
        };
      }
    } catch (e) {
      // Not JSON, treat as plain text response
    }
    
    // Return plain text response
    return {
      statusCode: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Access-Control-Allow-Methods': 'POST, OPTIONS'
      },
      body: JSON.stringify({ 
        response: aiResponse 
      })
    };
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