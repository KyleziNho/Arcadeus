#!/bin/bash

echo "🔍 Checking deployment status..."
echo ""
echo "📋 Recent commits:"
git log --oneline -5

echo ""
echo "📁 Files in last commit:"
git show --name-only HEAD

echo ""
echo "🌐 Testing deployed version..."
echo "Checking if changes are live at https://guavaexcel.netlify.app/"

# Check if the deployed chat.js has our revenue extraction changes
echo ""
echo "🔍 Checking for revenue extraction patterns in deployed API..."
curl -s "https://guavaexcel.netlify.app/.netlify/functions/chat" | head -20 || echo "API endpoint check (expected to fail with 405)"

echo ""
echo "💡 Next steps:"
echo "1. Check Netlify dashboard for deployment status"
echo "2. Wait 2-3 minutes for deployment to complete"
echo "3. Clear browser cache (Ctrl+Shift+R or Cmd+Shift+R)"
echo "4. Test again with your CSV file"
echo ""
echo "📝 If changes aren't showing:"
echo "- Verify Netlify is connected to your GitHub repo"
echo "- Check if auto-deploy is enabled in Netlify settings"
echo "- Manually trigger a deploy from Netlify dashboard"