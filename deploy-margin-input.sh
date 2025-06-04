#!/bin/bash

echo "💰 Deploying user-specified credit margin input..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated files
git add taskpane.html
git add taskpane.js

# Commit the margin input feature
git commit -m "Add user-specified credit margin input for floating rates

💰 Credit Margin Input:
- Added Credit Margin (%) input field for floating rate calculations
- Appears only when Floating Rate is selected
- Default value of 2.0% with industry standard guidance
- Placeholder suggests 2.0% as typical margin

🎯 Industry Standard Guidance:
- Helper text: 'Industry standard is typically 2% - 3% for most deals'
- Default 2% margin maintains existing behavior
- Users can adjust based on deal specifics and credit profile

🔧 Technical Implementation:
- Shows/hides margin input based on rate type selection
- Updates real-time preview calculations with user margin
- Integrates with Excel generation using specified margin
- Added to input change listeners for live updates

📊 Rate Calculation Updates:
- Fixed Rate: Uses user-specified fixed rate directly
- Floating Rate: Base Rate + User-specified Margin
- All-in Rate = Base Rate + Credit Margin (user input)
- Calculations update automatically when margin changes

✨ UI/UX Improvements:
- Clean conditional display logic
- Updated radio button label to 'Floating Rate (Base + Margin)'
- Professional helper text with industry context
- Seamless integration with existing debt model interface

This gives users full control over credit margin assumptions
while providing industry-standard guidance for reference.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Credit margin input deployed successfully!"
echo ""
echo "💰 New Features:"
echo "• Credit Margin (%) input field for floating rates"
echo "• Industry standard guidance (2% - 3% typical)"
echo "• Default 2% margin with user customization"
echo "• Real-time preview and Excel generation updates"
echo ""
echo "🎯 User Experience:"
echo "• Appears only when Floating Rate is selected"
echo "• Helper text provides industry context"
echo "• All-in Rate = Base Rate + User Margin"
echo "• Live updates in preview table and Excel output"
echo ""
echo "🧪 Test the functionality:"
echo "• Select Floating Rate to see margin input"
echo "• Try different margin values (1%, 2.5%, 3%)"
echo "• Watch preview table update in real-time"
echo "• Generate Excel schedule with custom margin"