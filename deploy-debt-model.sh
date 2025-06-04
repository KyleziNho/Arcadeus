#!/bin/bash

echo "💰 Deploying comprehensive debt model section..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage all updated files
git add taskpane.html
git add taskpane.css
git add taskpane.js

# Commit the debt model implementation
git commit -m "Add comprehensive debt model section with intelligent rate calculations

💰 Debt Model Features:
- Use Debt Financing toggle (Yes/No) with conditional display
- Interest Rate Type selection (Fixed vs Floating + 2% margin)
- Intelligent rate calculations based on US Fed base rates
- Debt issuance timing options (Acquisition, CapEx, Working Capital, Dividends)
- Real-time debt schedule preview with dynamic calculations
- Collapsible section with minimize functionality

🎯 Rate Intelligence:
- Fixed rate: User-defined percentage (default 5.5%)
- Floating rate: Base US Fed rate (default 3.9%) + automatic 2% margin
- All-in rate automatically calculated based on selection
- Real-time updates when rates change

📊 Schedule Preview:
- Dynamic table showing Period, Base Rate, All-in Rate, Outstanding Debt, Interest Payment
- Updates automatically based on holding period from Deal Assumptions
- Sample debt schedule with realistic amortization
- Excel generation button for full model creation

✨ UI/UX Design:
- Apple-inspired radio buttons and checkboxes with custom styling
- Smooth transitions and hover effects
- Professional table styling with alternating row colors
- Responsive design that works across screen sizes
- Consistent with existing design system

🔧 Technical Implementation:
- JavaScript event handlers for all form interactions
- Real-time calculation engine for debt schedules
- Integration with Deal Assumptions holding period
- Collapsible functionality with minimize/expand states
- Error handling and console logging for debugging

This creates a professional debt modeling interface that analysts
can use to intelligently calculate debt costs and generate Excel
schedules based on market rates and deal parameters.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Debt model section deployed successfully!"
echo ""
echo "💰 New Debt Model Features:"
echo "• Use Debt Financing toggle with conditional settings"
echo "• Fixed vs Floating rate selection with automatic margin calculation"
echo "• Debt issuance timing configuration options"
echo "• Real-time debt schedule preview with dynamic calculations"
echo "• Professional table layout matching your image requirements"
echo "• Collapsible section with minimize/expand functionality"
echo ""
echo "🧪 Test the functionality:"
echo "• Toggle 'Use Debt Financing' to Yes to reveal settings"
echo "• Switch between Fixed and Floating rate types"
echo "• Observe automatic 2% margin addition for floating rates"
echo "• Watch real-time schedule updates when changing rates"
echo "• Use minimize button to collapse debt model section"
echo "• Try 'Generate Debt Schedule in Excel' button"