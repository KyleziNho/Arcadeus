#!/bin/bash

echo "🎨 Deploying Apple-style UI improvements and debt schedule functionality..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated files
git add taskpane.css
git add taskpane.js

# Commit the UI improvements and functionality fixes
git commit -m "Apple-style UI improvements and functional debt schedule generation

🎨 Apple-style Radio Button Improvements:
- Increased size to 20px for better touch targets
- Added subtle shadow and hover animations with scale effect
- Enhanced spacing and padding for cleaner appearance
- Blue highlight with shadow when selected
- Smooth transitions and transform effects

✨ Apple-style Checkbox Enhancements:
- Card-style design with borders and hover effects
- Lift animation on hover with translateY and shadow
- Blue background tint when selected
- Better spacing and typography
- Scale animations for interactive feedback

⚡ Functional Debt Schedule Generation:
- Real Excel integration using Office.js API
- Calculates debt amount from Deal Size × LTV
- Generates professional formatted table in Excel
- Includes headers, borders, and auto-fitted columns
- Uses actual deal parameters from form inputs
- Smart amortization schedule with interest calculations
- Fallback messaging when Excel API unavailable

🔧 Technical Improvements:
- Async/await for proper Excel API handling
- Error handling with user-friendly messages
- Integration with Deal Assumptions data
- Real-time parameter validation
- Professional table formatting with Excel styling

💰 Enhanced Debt Calculations:
- Debt Amount = Deal Size × LTV percentage
- Interest Payment = Outstanding Debt × All-in Rate
- Simple amortization over holding period
- Proper number formatting for financial data

The debt model now has a polished Apple-style interface
and generates functional Excel debt schedules.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Apple-style UI and debt schedule functionality deployed!"
echo ""
echo "🎨 UI Improvements:"
echo "• Apple-style radio buttons with hover animations and shadows"
echo "• Card-style checkboxes with lift effects and blue highlights"
echo "• Better spacing, typography, and interactive feedback"
echo "• Smooth transitions throughout all form elements"
echo ""
echo "⚡ Functional Improvements:"
echo "• Generate Debt Schedule button now creates real Excel tables"
echo "• Uses actual deal parameters from your form inputs"
echo "• Professional formatting with borders and auto-fit columns"
echo "• Smart error handling and user feedback messages"
echo ""
echo "🧪 Test the improvements:"
echo "• Notice smoother radio button and checkbox interactions"
echo "• Fill out debt parameters and click Generate Debt Schedule"
echo "• Check Excel for professional formatted debt schedule table"
echo "• Try both fixed and floating rate calculations"