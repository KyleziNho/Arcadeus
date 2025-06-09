# API Configuration Setup

## Required Configuration

The M&A Intelligence Suite requires an OpenAI API key to function. Without this configuration, the system will not work.

## Setup Steps

### 1. Get OpenAI API Key
1. Go to https://platform.openai.com/api-keys
2. Sign in to your OpenAI account
3. Create a new API key
4. Copy the key (starts with `sk-`)

### 2. Configure on Netlify
1. Go to your Netlify dashboard
2. Select your deployed site
3. Go to **Site Settings** → **Environment Variables**
4. Add new environment variable:
   - **Key**: `OPENAI_API_KEY`
   - **Value**: Your OpenAI API key (sk-...)
5. Save the configuration
6. Trigger a new deployment

### 3. Verify Setup
1. Upload a file and click "Auto Fill with AI"
2. If configured correctly, extraction will work
3. If not configured, you'll see: "🚨 AI service is currently down"

## Current Status
- ❌ OpenAI API key not configured
- ❌ AI extraction will fail with 500 errors
- ❌ System is non-functional until API key is added

## Cost Information
- OpenAI charges per API call
- Typical extraction costs ~$0.01-0.05 per document
- Monitor usage in OpenAI dashboard

## Support
If you continue seeing errors after configuration:
1. Check OpenAI API key validity
2. Verify environment variable is set correctly
3. Check OpenAI service status
4. Contact support