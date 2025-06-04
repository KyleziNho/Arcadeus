#!/bin/bash

echo "🔧 Deploying Function Fix - Remove undefined calculateHoldingPeriods call..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Stage updated files
git add taskpane.js

# Commit the function fix
git commit -m "FIX: Remove undefined calculateHoldingPeriods function call

🔧 FUNCTION ERROR FIX: Remove missing function call

❌ Error Fixed:
- Removed call to this.calculateHoldingPeriods() which was undefined
- Simplified applyExtractedData function to avoid missing dependencies
- Cleaned up function calls to only use existing methods

✅ High-Level Parameters Extraction:
- Currency detection and setting
- Project start date extraction and formatting
- Project end date calculation and setting
- Model periods determination and application
- Detailed console logging for each step

🧪 Testing:
- Upload CSV file with financial data
- Click 'Auto Fill with AI'
- Check console for detailed extraction logs
- Verify High-Level Parameters section populates

The simplified extraction should now work without function errors.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Function Fix Deployed!"
echo ""
echo "🔧 Fixed Issues:"
echo "• Removed undefined calculateHoldingPeriods() call"
echo "• Simplified parameter application function"
echo "• Cleaned up function dependencies"
echo ""
echo "🧪 Test Now:"
echo "1. Upload your CSV file"
echo "2. Click 'Auto Fill with AI'"
echo "3. Check console for debug logs"
echo "4. Verify parameters get filled"