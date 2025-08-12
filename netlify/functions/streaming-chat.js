/**
 * Streaming Chat Function with Chain-of-Thought
 * Uses OpenAI SDK with structured outputs and streaming
 */

import OpenAI from 'openai';
import { z } from 'zod';
import { zodResponseFormat } from 'openai/helpers/zod';

// Initialize OpenAI client
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

// Define Zod schemas for chain-of-thought analysis
const AnalysisStep = z.object({
  step_number: z.number().describe("Step number in the analysis sequence"),
  action: z.string().describe("What I'm analyzing (e.g., 'Locating IRR calculation')"),
  excel_reference: z.string().describe("Excel cells being examined (e.g., 'FCF!B22')"),
  observation: z.string().describe("What I found in these cells"),
  calculation: z.string().optional().describe("Any calculations performed"),
  reasoning: z.string().describe("Why this step matters for the analysis")
});

const KeyMetric = z.object({
  name: z.string(),
  value: z.string(),
  location: z.string(),
  formula: z.string().optional(),
  interpretation: z.enum(["Excellent", "Strong", "Good", "Fair", "Poor", "Critical"])
});

const Recommendation = z.object({
  action: z.string(),
  expected_impact: z.string(),
  cells_to_modify: z.array(z.string()),
  priority: z.enum(["high", "medium", "low"])
});

const FinancialAnalysisSchema = z.object({
  query_interpretation: z.string().describe("Clear understanding of what the user is asking"),
  analysis_steps: z.array(AnalysisStep).describe("Step-by-step walkthrough of the analysis"),
  key_metrics: z.object({
    primary: KeyMetric,
    supporting: z.array(KeyMetric)
  }),
  insights: z.array(z.object({
    title: z.string(),
    content: z.string(),
    type: z.enum(["positive", "negative", "neutral", "warning"]),
    impact: z.enum(["high", "medium", "low"])
  })),
  final_answer: z.string().describe("Direct, comprehensive answer to the user's question"),
  recommendations: z.array(Recommendation),
  next_steps: z.array(z.string()).describe("Immediate actionable steps")
});

// Excel Structure Analysis Schema
const ExcelStructureSchema = z.object({
  query_interpretation: z.string(),
  structure_analysis: z.array(z.object({
    step_number: z.number(),
    action: z.string(),
    location: z.string(),
    formula: z.string().optional(),
    dependencies: z.array(z.string()).optional(),
    explanation: z.string()
  })),
  formula_breakdown: z.object({
    main_formula: z.string(),
    components: z.array(z.object({
      part: z.string(),
      purpose: z.string(),
      cells_referenced: z.array(z.string())
    }))
  }),
  validation_checks: z.array(z.object({
    check: z.string(),
    result: z.enum(["passed", "failed", "warning"]),
    details: z.string()
  })),
  final_answer: z.string()
});

// General Query Schema (fallback)
const GeneralQuerySchema = z.object({
  answer: z.string(),
  explanation: z.string().optional(),
  references: z.array(z.string()).optional(),
  follow_up_questions: z.array(z.string()).optional()
});

/**
 * Determine which schema to use based on query type
 */
function selectSchema(message) {
  const lowerMessage = message.toLowerCase();
  
  if (lowerMessage.includes('irr') || lowerMessage.includes('moic') || 
      lowerMessage.includes('return') || lowerMessage.includes('multiple') ||
      lowerMessage.includes('cash flow') || lowerMessage.includes('npv')) {
    return { schema: FinancialAnalysisSchema, name: 'financial_analysis' };
  }
  
  if (lowerMessage.includes('formula') || lowerMessage.includes('calculation') ||
      lowerMessage.includes('cell') || lowerMessage.includes('reference')) {
    return { schema: ExcelStructureSchema, name: 'excel_structure' };
  }
  
  return { schema: GeneralQuerySchema, name: 'general_query' };
}

/**
 * Build system prompt based on query type and Excel context
 */
function buildSystemPrompt(queryType, excelContext) {
  const basePrompt = `You are an expert M&A financial analyst working directly in Excel as an add-in assistant.`;
  
  if (queryType === 'financial_analysis') {
    return `${basePrompt}
    
You are analyzing a live M&A financial model. Walk through your analysis step-by-step, showing exactly how you arrive at conclusions.
For each step:
1. State what you're looking for
2. Reference the specific Excel cells you're examining
3. Show what you found (values, formulas)
4. Explain any calculations you perform
5. Explain why this matters for answering the question

Current Excel Context:
${JSON.stringify(excelContext?.financialMetrics || {}, null, 2)}

Key areas in the model:
- IRR calculations: ${excelContext?.analysis?.financialAreas?.irr?.[0]?.location || 'Not found'}
- MOIC calculations: ${excelContext?.analysis?.financialAreas?.moic?.[0]?.location || 'Not found'}
- Cash flows: ${excelContext?.analysis?.financialAreas?.cashFlow?.[0]?.location || 'Not found'}

Be specific with cell references and show your work.`;
  }
  
  if (queryType === 'excel_structure') {
    return `${basePrompt}
    
You are examining Excel formulas and cell relationships. Break down complex formulas step-by-step.
Show formula dependencies and explain how calculations flow through the model.
Reference specific cells and ranges.

Current worksheet structure:
${JSON.stringify(excelContext?.structure || {}, null, 2)}`;
  }
  
  return `${basePrompt}
  
Provide clear, actionable guidance about M&A financial modeling in Excel.
Reference specific features and best practices.`;
}

/**
 * Main handler function with streaming support
 */
export const handler = async (event, context) => {
  try {
    const { message, excelContext, streaming = true } = JSON.parse(event.body);
    
    // Select appropriate schema
    const { schema, name: queryType } = selectSchema(message);
    
    // Build system prompt
    const systemPrompt = buildSystemPrompt(queryType, excelContext);
    
    if (streaming) {
      // Create streaming response with structured output
      const stream = await openai.beta.chat.completions.stream({
        model: 'gpt-4o-2024-08-06',
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: message }
        ],
        response_format: zodResponseFormat(schema, queryType),
        stream: true,
        temperature: queryType === 'financial_analysis' ? 0.3 : 0.5,
        max_tokens: 2000
      });
      
      // Set up Server-Sent Events headers
      return {
        statusCode: 200,
        headers: {
          'Content-Type': 'text/event-stream',
          'Cache-Control': 'no-cache',
          'Connection': 'keep-alive',
          'Access-Control-Allow-Origin': '*'
        },
        // This won't work with Netlify Functions - need to use regular response
        // We'll handle streaming differently
        body: JSON.stringify({
          streaming: false,
          message: "Streaming not supported in Netlify Functions. Using parse method instead."
        })
      };
    } else {
      // Non-streaming parse method with structured output
      const completion = await openai.beta.chat.completions.parse({
        model: 'gpt-4o-2024-08-06',
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: message }
        ],
        response_format: zodResponseFormat(schema, queryType),
        temperature: queryType === 'financial_analysis' ? 0.3 : 0.5,
        max_tokens: 2000
      });
      
      const parsed = completion.choices[0].message.parsed;
      
      return {
        statusCode: 200,
        headers: {
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*'
        },
        body: JSON.stringify({
          success: true,
          queryType,
          parsed,
          usage: completion.usage
        })
      };
    }
    
  } catch (error) {
    console.error('Streaming chat error:', error);
    
    // Check for specific OpenAI errors
    if (error.message?.includes('refusal')) {
      return {
        statusCode: 200,
        headers: {
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*'
        },
        body: JSON.stringify({
          success: false,
          error: 'The model refused to answer this query',
          refusal: error.message
        })
      };
    }
    
    return {
      statusCode: 500,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*'
      },
      body: JSON.stringify({
        success: false,
        error: error.message
      })
    };
  }
};