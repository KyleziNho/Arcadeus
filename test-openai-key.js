#!/usr/bin/env node

const fetch = require('node-fetch');

// Colors for console output
const colors = {
    green: '\x1b[32m',
    red: '\x1b[31m',
    yellow: '\x1b[33m',
    blue: '\x1b[34m',
    reset: '\x1b[0m',
    bold: '\x1b[1m'
};

console.log(`${colors.blue}${colors.bold}🔑 OpenAI API Key Test${colors.reset}\n`);

async function testOpenAIKey() {
    // Get API key from environment variable
    const apiKey = process.env.OPENAI_API_KEY;
    
    if (!apiKey) {
        console.log(`${colors.red}❌ No OpenAI API key found!${colors.reset}`);
        console.log(`Please set your API key:`);
        console.log(`export OPENAI_API_KEY="your-key-here"`);
        process.exit(1);
    }
    
    console.log(`${colors.blue}🔍 Testing API key...${colors.reset}`);
    console.log(`Key prefix: ${apiKey.substring(0, 8)}...`);
    console.log(`Key length: ${apiKey.length} characters\n`);
    
    try {
        // Make a minimal test request
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${apiKey}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                model: 'gpt-4-turbo-preview',
                messages: [
                    { role: 'user', content: 'Say "API test successful"' }
                ],
                max_tokens: 10,
                temperature: 0.1
            })
        });
        
        console.log(`${colors.blue}📡 API Response Status: ${response.status}${colors.reset}`);
        
        if (!response.ok) {
            const errorText = await response.text();
            let errorData;
            try {
                errorData = JSON.parse(errorText);
            } catch (e) {
                errorData = { message: errorText };
            }
            
            console.log(`${colors.red}❌ API Error:${colors.reset}`);
            console.log(`Status: ${response.status}`);
            console.log(`Error: ${errorData.error?.message || errorData.message || errorText}\n`);
            
            // Specific error guidance
            switch (response.status) {
                case 401:
                    console.log(`${colors.yellow}💡 Troubleshooting Tips:${colors.reset}`);
                    console.log(`- Check if your API key is correct`);
                    console.log(`- Make sure the key hasn't expired`);
                    console.log(`- Verify you're using the right OpenAI account`);
                    break;
                case 429:
                    console.log(`${colors.yellow}💡 Troubleshooting Tips:${colors.reset}`);
                    console.log(`- You may have exceeded rate limits`);
                    console.log(`- Check if you have sufficient credits`);
                    console.log(`- Try again in a few minutes`);
                    break;
                case 400:
                    console.log(`${colors.yellow}💡 Troubleshooting Tips:${colors.reset}`);
                    console.log(`- Request format issue (this shouldn't happen with our test)`);
                    console.log(`- Check if the model 'gpt-4-turbo-preview' is available to your account`);
                    break;
                default:
                    console.log(`${colors.yellow}💡 Unexpected error - check OpenAI status${colors.reset}`);
            }
            return false;
        }
        
        const data = await response.json();
        
        console.log(`${colors.green}✅ API Test Successful!${colors.reset}\n`);
        
        // Show response details
        console.log(`${colors.bold}📊 Response Details:${colors.reset}`);
        console.log(`Model: ${data.model}`);
        console.log(`Response: "${data.choices[0].message.content}"`);
        
        if (data.usage) {
            console.log(`\n${colors.bold}📈 Token Usage:${colors.reset}`);
            console.log(`Prompt tokens: ${data.usage.prompt_tokens}`);
            console.log(`Completion tokens: ${data.usage.completion_tokens}`);
            console.log(`Total tokens: ${data.usage.total_tokens}`);
        }
        
        console.log(`\n${colors.green}🎉 Your OpenAI API key is working correctly!${colors.reset}`);
        console.log(`You can now use this key in your Netlify environment variables.`);
        
        return true;
        
    } catch (error) {
        console.log(`${colors.red}❌ Network/Connection Error:${colors.reset}`);
        console.log(`Error: ${error.message}\n`);
        
        console.log(`${colors.yellow}💡 Troubleshooting Tips:${colors.reset}`);
        console.log(`- Check your internet connection`);
        console.log(`- Verify you can access https://api.openai.com`);
        console.log(`- Try again in a few minutes`);
        
        return false;
    }
}

// Run the test
testOpenAIKey().then(success => {
    process.exit(success ? 0 : 1);
}).catch(error => {
    console.error(`${colors.red}Unexpected error:${colors.reset}`, error);
    process.exit(1);
});