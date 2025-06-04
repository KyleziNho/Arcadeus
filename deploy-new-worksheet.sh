#!/bin/bash

echo "📊 Deploying debt schedule with new worksheet creation..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated JavaScript file
git add taskpane.js

# Commit the new worksheet feature
git commit -m "Create debt schedule in new worksheet with proper formatting

📊 New Worksheet Creation:
- Creates new 'Debt Schedule' worksheet instead of modifying current sheet
- Deletes existing 'Debt Schedule' worksheet if it exists before creating new one
- Activates the new worksheet automatically for immediate viewing
- Protects user's current worksheet from being modified

🎯 Transposed Layout (like provided image):
- Period headers across columns (monthly periods: 1-Jan-25, 2-Feb-25, etc.)
- Base interest rate and All-in interest rate as rows
- Proper date formatting with month abbreviations and 2-digit years
- Row labels in first column for easy identification

🎨 Professional Formatting:
- Gray header row with 'Debt Model' title and merged cells
- Green background for period header row
- Bold formatting for all headers and row labels
- Complete table borders for clean appearance
- Auto-fitted columns for optimal display

📅 Dynamic Period Generation:
- Creates monthly periods based on holding period input
- Caps at 60 months maximum for reasonable display
- Generates proper date sequence starting from current date
- Prepares foundation for future daily/monthly frequency options

This ensures users can generate debt schedules without affecting
their current worksheet while maintaining professional formatting.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ New worksheet debt schedule deployed successfully!"
echo ""
echo "📊 New Features:"
echo "• Creates new 'Debt Schedule' worksheet"
echo "• Protects current worksheet from modification"
echo "• Transposed layout with periods as columns"
echo "• Professional formatting similar to provided image"
echo ""
echo "🎯 Formatting Applied:"
echo "• Gray header with merged 'Debt Model' title"
echo "• Green period header row with date formatting"
echo "• Bold row labels and proper table borders"
echo "• Auto-fitted columns for optimal viewing"
echo ""
echo "📅 Period Structure:"
echo "• Monthly periods (1-Jan-25, 2-Feb-25, etc.)"
echo "• Based on holding period input"
echo "• Capped at 60 months for display"
echo "• Ready for future daily/monthly frequency options"
echo ""
echo "🧪 Test the functionality:"
echo "• Fill out debt model form with holding period"
echo "• Click Generate Debt Schedule"
echo "• Verify new worksheet is created and activated"
echo "• Check formatting matches the provided image style"