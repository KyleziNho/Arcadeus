class ModelProviderManager {
  constructor() {
    this.providers = {
      'gpt-4': {
        name: 'ChatGPT 4',
        endpoint: 'https://api.openai.com/v1/chat/completions',
        model: 'gpt-4-turbo-preview',
        headers: (apiKey) => ({
          'Authorization': `Bearer ${apiKey}`,
          'Content-Type': 'application/json'
        }),
        supportsStreaming: true
      },
      'claude-opus': {
        name: 'Claude Opus 3',
        endpoint: 'https://api.anthropic.com/v1/messages',
        model: 'claude-3-opus-20240229',
        headers: (apiKey) => ({
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'Content-Type': 'application/json'
        }),
        supportsStreaming: true
      },
      'claude-sonnet': {
        name: 'Claude Sonnet 3.5',
        endpoint: 'https://api.anthropic.com/v1/messages',
        model: 'claude-3-5-sonnet-20241022',
        headers: (apiKey) => ({
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'Content-Type': 'application/json'
        }),
        supportsStreaming: true
      },
      'gemini-pro': {
        name: 'Gemini 1.5 Pro',
        endpoint: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:streamGenerateContent',
        headers: (apiKey) => ({
          'Content-Type': 'application/json'
        }),
        supportsStreaming: true
      }
    };
    
    this.currentProvider = 'gpt-4';
    this.apiKeys = this.loadApiKeys();
  }

  loadApiKeys() {
    // Load from secure storage or environment variables
    return {
      'gpt-4': localStorage.getItem('openai_api_key'),
      'claude-opus': localStorage.getItem('anthropic_api_key'),
      'claude-sonnet': localStorage.getItem('anthropic_api_key'),
      'gemini-pro': localStorage.getItem('google_api_key')
    };
  }

  setProvider(providerId) {
    if (this.providers[providerId]) {
      this.currentProvider = providerId;
      console.log(`Switched to ${this.providers[providerId].name}`);
      return true;
    }
    return false;
  }

  async sendMessage(messages, context, onStream) {
    const provider = this.providers[this.currentProvider];
    const apiKey = this.apiKeys[this.currentProvider];
    
    if (!apiKey) {
      throw new Error(`API key not configured for ${provider.name}`);
    }

    // Format message based on provider
    const formattedRequest = this.formatRequest(messages, context, provider);
    
    if (provider.supportsStreaming && onStream) {
      return this.streamResponse(provider, apiKey, formattedRequest, onStream);
    } else {
      return this.normalResponse(provider, apiKey, formattedRequest);
    }
  }

  formatRequest(messages, context, provider) {
    const systemPrompt = this.buildSystemPrompt(context);
    
    switch (this.currentProvider) {
      case 'gpt-4':
        return {
          model: provider.model,
          messages: [
            { role: 'system', content: systemPrompt },
            ...messages
          ],
          stream: true,
          temperature: 0.7
        };
        
      case 'claude-opus':
      case 'claude-sonnet':
        return {
          model: provider.model,
          system: systemPrompt,
          messages: messages.map(m => ({
            role: m.role === 'assistant' ? 'assistant' : 'user',
            content: m.content
          })),
          stream: true,
          max_tokens: 4096
        };
        
      case 'gemini-pro':
        return {
          contents: [
            {
              role: 'user',
              parts: [{ text: systemPrompt + '\n\n' + messages[messages.length - 1].content }]
            }
          ],
          generationConfig: {
            temperature: 0.7,
            maxOutputTokens: 4096
          }
        };
        
      default:
        throw new Error(`Unknown provider: ${this.currentProvider}`);
    }
  }

  buildSystemPrompt(context) {
    let prompt = `You are an expert M&A financial analyst assistant operating within Excel. 
You have access to the user's Excel workbook and can see the currently selected cells and all worksheet data.

CURRENT EXCEL CONTEXT:
${JSON.stringify(context.excel, null, 2)}

FORM DATA:
${JSON.stringify(context.formData, null, 2)}

You can:
1. Read and analyze any data in the Excel workbook
2. Suggest formulas and financial calculations
3. Generate new sheets or modify existing ones
4. Validate financial models
5. Provide M&A insights based on the data

When suggesting changes, be specific about cell references and formulas.`;
    
    return prompt;
  }

  async streamResponse(provider, apiKey, request, onStream) {
    const endpoint = provider.endpoint;
    const headers = provider.headers(apiKey);
    
    try {
      const response = await fetch(endpoint, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify(request)
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';
      let fullResponse = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
          if (line.trim().startsWith('data:')) {
            const data = line.slice(5).trim();
            if (data === '[DONE]') continue;
            
            try {
              const parsed = JSON.parse(data);
              const chunk = this.extractChunk(parsed);
              if (chunk) {
                fullResponse += chunk;
                onStream(chunk);
              }
            } catch (e) {
              console.error('Error parsing stream chunk:', e);
            }
          }
        }
      }

      return fullResponse;
    } catch (error) {
      console.error('Stream error:', error);
      throw error;
    }
  }

  extractChunk(parsed) {
    switch (this.currentProvider) {
      case 'gpt-4':
        return parsed.choices?.[0]?.delta?.content || '';
      case 'claude-opus':
      case 'claude-sonnet':
        return parsed.delta?.text || '';
      case 'gemini-pro':
        return parsed.candidates?.[0]?.content?.parts?.[0]?.text || '';
      default:
        return '';
    }
  }

  async normalResponse(provider, apiKey, request) {
    const endpoint = provider.endpoint;
    const headers = provider.headers(apiKey);
    
    const response = await fetch(endpoint, {
      method: 'POST',
      headers: headers,
      body: JSON.stringify(request)
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    
    switch (this.currentProvider) {
      case 'gpt-4':
        return data.choices?.[0]?.message?.content || '';
      case 'claude-opus':
      case 'claude-sonnet':
        return data.content?.[0]?.text || '';
      case 'gemini-pro':
        return data.candidates?.[0]?.content?.parts?.[0]?.text || '';
      default:
        return '';
    }
  }
}

window.ModelProviderManager = ModelProviderManager;