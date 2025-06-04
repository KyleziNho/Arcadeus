#!/bin/bash

echo "📊 Deploying final debt model with comprehensive deal summary..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated JavaScript file
git add taskpane.js

# Commit the comprehensive debt model
git commit -m "Finalize debt model with comprehensive deal summary and correct calculations

📊 Comprehensive Excel Output:
- Added Deal Summary section with all form inputs
- Deal Name, Deal Size, LTV, Debt Amount, Rate Type, Credit Margin, Holding Period
- Professional blue header for Deal Summary section
- Clean Rate Schedule section with Base Rate and All-in Rate only
- Removed Outstanding Debt and Interest Payment rows for cleaner output

🎯 Accurate Calculations:
- Uses actual Deal Size from form input (not hardcoded values)
- Calculates Debt Amount = Deal Size × LTV percentage
- Base Rate from user input (fixed or floating)
- All-in Rate = Base Rate + User-specified Credit Margin
- Holding Period converted to years for proper period calculation

💼 Professional Formatting:
- Blue header for Deal Summary section with white text
- Bold formatting for all labels and important data
- Proper table borders and auto-fitted columns
- Two-section layout: Summary at top, Rate Schedule below
- Merged header cells for clean appearance

🔧 Form Integration:
- Pulls Deal Name from Deal Assumptions section
- Uses Deal Size, LTV, Holding Period from existing inputs
- Integrates Rate Type (Fixed/Floating) selection
- Includes user-specified Credit Margin for floating rates
- All calculations use actual form values

✨ Preview Table Updates:
- Simplified preview table with Base Rate and All-in Rate only
- Removed debt and interest calculations from preview
- Real-time updates when any input changes
- Consistent with Excel output format

This creates a professional debt model worksheet that analysts
can use for client presentations and deal documentation.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com)"

# Push to main
git push origin main

echo "✅ Final debt model deployed successfully!"
echo ""
echo "📊 Excel Output Features:"
echo "• Comprehensive Deal Summary with all form inputs"
echo "• Professional blue header with deal information"
echo "• Rate Schedule with Base Rate and All-in Rate"
echo "• Removed Outstanding Debt and Interest Payment"
echo "• Uses actual form values for all calculations"
echo ""
echo "🎯 Accurate Data:"
echo "• Deal Name from Deal Assumptions"
echo "• Deal Size, LTV, Holding Period from form inputs"
echo "• Debt Amount = Deal Size × LTV"
echo "• All-in Rate = Base Rate + Credit Margin"
echo "• Period calculation from Holding Period"
echo ""
echo "🧪 Test the functionality:"
echo "• Fill out all deal assumption fields"
echo "• Configure debt settings (rate type, margin)"
echo "• Generate debt schedule to see comprehensive Excel output"
echo "• Verify all form values appear correctly in worksheet"