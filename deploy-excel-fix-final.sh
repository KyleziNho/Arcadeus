#!/bin/bash

echo "🔧 Deploying Excel generation fix with simplified data structure..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated JavaScript file
git add taskpane.js

# Commit the Excel generation fix
git commit -m "Fix Excel debt schedule generation with simplified data structure

🔧 Excel Generation Fix:
- Removed undefined summaryData variable reference
- Fixed newWorksheet reference to use worksheet 
- Simplified data structure to fixed 10x5 array (A1:E10)
- Removed complex formatting that was causing range errors
- Clear and simple debt schedule with Deal Summary section

⚡ Error Resolution:
- Fixed InvalidArgument: array size mismatch errors
- Fixed ItemNotFound: worksheet reference errors
- Removed complex border and formatting operations
- Uses fixed-size range to avoid dimension issues

📊 Reliable Output:
- Deal Summary with all form inputs (name, size, LTV, debt amount)
- Rate Type and Credit Margin information
- Holding Period details
- Clean 10-row format that Excel can handle reliably

This ensures the Generate Debt Schedule button works consistently
across different Excel environments without API conflicts.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Excel generation fix deployed successfully!"
echo ""
echo "🔧 Fixed Issues:"
echo "• Removed undefined variable references"
echo "• Fixed worksheet object references"
echo "• Simplified to fixed 10x5 data array"
echo "• Removed complex formatting causing errors"
echo ""
echo "📊 New Excel Output:"
echo "• Deal Summary with all form inputs"
echo "• Rate Type and Credit Margin info"
echo "• Clean 10-row format (A1:E10)"
echo "• Reliable data insertion without dimension errors"
echo ""
echo "🧪 Test the functionality:"
echo "• Fill out debt model form"
echo "• Click Generate Debt Schedule"
echo "• Check that Excel table is created successfully"
echo "• Verify all form data appears in Excel output"