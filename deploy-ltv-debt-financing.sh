#!/bin/bash

echo "💰 Deploying LTV-based debt financing with automatic eligibility checking..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated files
git add taskpane.html
git add taskpane.css
git add taskpane.js

# Commit the LTV-based debt financing updates
git commit -m "Implement LTV-based debt financing with automatic eligibility and loan issuance fees

💰 Smart Debt Financing Logic:
- Automatically enables/disables debt financing based on Deal LTV input
- LTV = 0%: Shows 'Please input a higher LTV to access debt financing options'
- LTV > 0%: Shows 'Debt financing available (X% LTV)' with green text
- Removes manual Yes/No radio buttons for streamlined UX

🏦 Loan Issuance Fees Addition:
- New 'Loan Issuance Fees (%)' field as first input in debt settings
- Default value: 1.5% (industry standard for loan arrangement fees)
- Help text: 'Fees for arranging and issuing the debt financing'
- Integrated with existing rate type and margin calculations

🎨 Professional Status Display:
- Clean status message box with bordered container
- Green text for enabled state with bold font weight
- Gray italic text for disabled state
- Real-time updates when Deal LTV changes in Deal Assumptions

🔧 Technical Integration:
- Automatic debt eligibility checking when Deal LTV changes
- Updated debt schedule generation to check LTV instead of manual toggle
- Cross-section communication between Deal Assumptions and Debt Model
- Maintains all existing debt calculation and Excel generation functionality

⚡ Enhanced User Experience:
- Eliminates manual debt financing toggle confusion
- Immediate visual feedback when LTV is adjusted
- Cleaner interface with automatic logic
- Professional status messaging for guidance

🧮 Smart Validation:
- Debt schedule generation requires LTV > 0%
- Clear error messaging for invalid states
- Automatic hiding/showing of debt settings
- Real-time preview updates based on LTV availability

This creates an intelligent debt financing system that automatically
adapts based on deal parameters with professional presentation.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ LTV-based debt financing deployed successfully!"
echo ""
echo "💰 New Debt Financing Logic:"
echo "• Automatic enable/disable based on Deal LTV input"
echo "• LTV = 0%: Disabled with guidance message"
echo "• LTV > 0%: Enabled with confirmation message"
echo "• Removes manual Yes/No toggle for streamlined UX"
echo ""
echo "🏦 Loan Issuance Fees:"
echo "• New field for loan arrangement fees"
echo "• Default 1.5% (industry standard)"
echo "• Positioned as first input in debt settings"
echo "• Integrated with existing calculations"
echo ""
echo "🎨 Professional Status Display:"
echo "• Clean bordered status container"
echo "• Green text for enabled state"
echo "• Gray italic text for disabled state"
echo "• Real-time updates with LTV changes"
echo ""
echo "🧪 Test the functionality:"
echo "• Set Deal LTV to 0% - verify debt section is disabled"
echo "• Set Deal LTV to 70% - verify debt section enables"
echo "• Check status message updates in real-time"
echo "• Try generating debt schedule with different LTV values"