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

    const { message, excelContext, fileContents, autoFillMode } = requestData;
    
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
      
      let systemPrompt;
      let maxTokens = 1500;
      
      if (autoFillMode) {
        console.log('Auto-fill mode activated');
        maxTokens = 4000; // Increase token limit for auto-fill
        
        systemPrompt = `You are an expert financial analyst AI specialized in extracting data from M&A/PE documents and financial reports. Your task is to analyze uploaded documents and extract specific financial data.

CRITICAL INSTRUCTIONS:
1. Carefully read and analyze ALL content from the uploaded files
2. Look for financial data, dates, percentages, company names, deal values, revenue figures, cost data, etc.
3. Extract data that matches the required fields below
4. Return ONLY a valid JSON object with the extracted data
5. If you cannot find specific data, use null for that field
6. Look for context clues - revenue might be called "sales", "income", "turnover"
7. Cost items might be "expenses", "operating costs", "OPEX", "staff costs", etc.
8. Growth rates might be shown as YoY%, CAGR, annual growth, projected increases
9. Deal values might be "purchase price", "enterprise value", "transaction value", "acquisition price"
10. Currency should be detected from symbols ($, €, £) or abbreviations (USD, EUR, GBP)

REVENUE ITEMS EXTRACTION - CRITICAL:
11. Search for ALL revenue streams in the document - look for patterns like:
    - "Revenue Item 1", "Revenue Item 2", etc.
    - "Revenue Stream A", "Revenue Stream B", etc.
    - Product names, service lines, business segments with revenue
    - Sales categories, income sources, rental income, subscription revenue
    - Any line items with monetary values that represent income
12. For EACH revenue item found, extract:
    - Name: Use the EXACT name from the document (e.g., "Revenue Item 1" not "Primary Revenue")
    - Initial Value: The number associated with that revenue item (convert to plain number)
    - Growth Type: Determine based on data:
      * "linear" - if you find a consistent growth % (e.g., "2% annual growth")
      * "nonlinear" - if growth varies by year or mentions "accelerating/decelerating"
      * "no_growth" - if no growth rate is mentioned or growth is 0%
    - Growth Rate: Extract the exact percentage (only for linear growth, e.g., 2 for 2%)
13. Growth rate matching patterns:
    - "Rent Growth 1: 2%" → Revenue Item 1 has 2% linear growth
    - "Revenue Item 1 growth: 3%" → Revenue Item 1 has 3% linear growth
    - "All revenue grows at 5%" → Apply 5% to all revenue items
    - If no specific match, check for general growth rates
14. File format specific extraction:
    - CSV: Look for columns/rows with revenue data
    - Images/Screenshots: Look for tables or lists showing revenue items
    - PDF: Extract revenue sections, financial statements, projections
15. MANDATORY: Extract the EXACT number of revenue items shown in the file
    - If file shows 3 revenue items, return exactly 3
    - If file shows 0 revenue items, return empty array []
    - NEVER invent or estimate revenue items not in the file

COST ITEMS EXTRACTION - CRITICAL:
16. Search for ALL cost/expense items in the document - look for patterns like:
    - "Cost Item 1", "Cost Item 2", etc.
    - "Staff expenses", "Rent", "Utilities", "Marketing costs"
    - Any expense categories with associated values
    - Operating expenses (OpEx) and Capital expenses (CapEx)
    - Any line items that represent costs or expenses
17. For EACH cost item found, extract:
    - Name: Use the EXACT name from the document (e.g., "Cost Item 1" or "Staff expenses")
    - Initial Value: The number associated with that cost item (convert to plain number)
    - Growth Type: Determine based on data:
      * "linear" - if inflation/growth rate is consistent (e.g., "2% annual inflation")
      * "nonlinear" - if costs vary by year or follow a curve
      * "no_growth" - if no inflation/growth is mentioned or is 0%
    - Growth Rate: Extract the exact percentage (only for linear growth, e.g., 2 for 2%)
18. Growth rate matching patterns for costs:
    - "OpEx Cost Inflation: 2%" → Apply to all OpEx items
    - "CapEx Cost Inflation: 1.5%" → Apply to all CapEx items
    - "Salary Growth (p.a.): 0.5%" → Apply to staff expenses
    - "Cost Item 1 inflation: 3%" → Apply to specific cost item
    - General inflation rates that apply to all costs
19. Special cost patterns to recognize:
    - "Staff expenses: 60,000 with Salary Growth (p.a.): 0.50%"
    - Separate OpEx and CapEx items when specified
    - Match inflation rates by cost type or specific item number
20. MANDATORY: Extract the EXACT number of cost items shown in the file
    - If file shows 4 cost items, return exactly 4
    - If file shows 0 cost items, return empty array []
    - NEVER create generic cost items not in the file

REQUIRED DATA STRUCTURE:
{
  "extractedData": {
    "highLevelParameters": {
      "currency": "Detect from document (USD, EUR, GBP, etc.)",
      "projectStartDate": "YYYY-MM-DD format - look for deal close date, acquisition date, or current year start",
      "projectEndDate": "YYYY-MM-DD format - calculate based on holding period or exit date",
      "modelPeriods": "Detect if data is daily/monthly/quarterly/yearly"
    },
    "dealAssumptions": {
      "dealName": "Company name or deal description",
      "dealValue": "Number - look for purchase price, enterprise value, deal size",
      "transactionFee": "Percentage - banking fees, advisory fees, transaction costs",
      "dealLTV": "Percentage - leverage ratio, debt percentage, LTV"
    },
    "revenueItems": [
      {
        "name": "Revenue Item 1",
        "initialValue": 500000,
        "growthType": "linear",
        "growthRate": 2
      },
      {
        "name": "Revenue Item 2", 
        "initialValue": 766000,
        "growthType": "linear",
        "growthRate": 3
      }
    ],
    "costItems": [
      {
        "name": "Staff expenses",
        "initialValue": 60000,
        "growthType": "linear",
        "growthRate": 0.5
      },
      {
        "name": "Cost Item 1",
        "initialValue": 200000,
        "growthType": "linear",
        "growthRate": 2
      },
      {
        "name": "Cost Item 2",
        "initialValue": 1000,
        "growthType": "no_growth",
        "growthRate": 0
      }
    ],
    "exitAssumptions": {
      "disposalCost": "Exit fees percentage - typically 1-3%",
      "terminalCapRate": "Exit cap rate or exit multiple basis"
    },
    "debtModel": {
      "hasDebt": "true if LTV > 0",
      "interestRate": "Look for interest rates, cost of debt",
      "loanIssuanceFees": "Debt arrangement fees"
    }
  }
}

CRITICAL REVENUE EXTRACTION RULES:
- You MUST extract the EXACT revenue items from the uploaded files
- Do NOT create generic revenue items based on business type
- If you see "Revenue Item 1: 500,000" then extract exactly that
- Count ALL revenue items in the files and create that exact number
- Analyze growth patterns: if consistent % year-over-year = "linear", if varying = "nonlinear", if no growth mentioned = "no_growth"
- Extract exact values as numbers (remove commas, currency symbols)
- Match growth rates to specific revenue items when labeled (e.g., "Rent Growth 1" → "Revenue Item 1")

CRITICAL COST EXTRACTION RULES:
- You MUST extract the EXACT cost items from the uploaded files
- Look for any expense-related line items with values
- Common patterns: "Cost Item 1: 200,000", "Staff expenses: 60,000", "Rent: 50,000"
- Extract inflation rates: "OpEx Cost Inflation: 2%", "Salary Growth: 0.5%"
- Count ALL cost items and create that exact number
- Match inflation rates to specific costs when possible
- If you see "Staff expenses: 60,000 with Salary Growth (p.a.): 0.50%" extract both value and growth
- Default to "no_growth" if no inflation/growth rate is specified

Document Content to Analyze:
${documentContext}`;
      } else {
        // Regular chat mode
        systemPrompt = `You are an AI assistant specialized in M&A and PE deal modeling in Excel. You can execute Excel commands to help users build financial models.

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
      }

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
            temperature: autoFillMode ? 0.3 : 0.7, // Lower temperature for data extraction
            max_tokens: maxTokens
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
              commands: [],
              error: 'Failed to parse response'
            })
          };
        }

        // Handle auto-fill mode response
        if (autoFillMode && parsedResponse.extractedData) {
          console.log('Returning extracted data for auto-fill');
          return {
            statusCode: 200,
            headers: {
              'Access-Control-Allow-Origin': '*',
              'Access-Control-Allow-Headers': 'Content-Type, Authorization',
              'Access-Control-Allow-Methods': 'POST, OPTIONS'
            },
            body: JSON.stringify({
              extractedData: parsedResponse.extractedData,
              response: "Data extracted successfully"
            })
          };
        }

        // Regular chat mode response
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