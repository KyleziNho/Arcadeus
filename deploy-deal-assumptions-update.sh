#!/bin/bash

echo "📊 Deploying updated Deal Assumptions with financial calculations..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated files
git add taskpane.html
git add taskpane.js

# Commit the Deal Assumptions updates
git commit -m "Update Deal Assumptions with financial parameters and smart calculations

📊 New Deal Assumptions Structure:
- Deal Name: Text input for transaction identification
- Deal Value: Integer input for total transaction value
- Transaction Fee (%): Percentage input with 2.5% default
- Deal LTV (%): Loan-to-Value ratio with 70% default
- Equity Contribution (Calculated): Auto-calculated from Deal Value × (100% - LTV%)
- Debt Financing (Calculated): Auto-calculated from Deal Value × LTV%

💰 Smart Financial Calculations:
- Real-time calculation updates when Deal Value or LTV changes
- Professional currency formatting based on selected currency
- Proper formatting with thousands separators and currency symbols
- Integration with High-Level Parameters currency selection
- Readonly fields for calculated values to prevent manual override

🎨 Professional UI Features:
- Help text for each field explaining purpose and calculation
- Default values for typical deal parameters
- Currency-aware formatting (USD, EUR, GBP, etc.)
- Proper input validation and error handling
- Consistent styling with other sections

🔧 Technical Integration:
- Updated debt model to use new Deal Value and Deal LTV fields
- Maintains compatibility with existing debt schedule generation
- Proper event listeners for real-time calculations
- Integration with currency selection from High-Level Parameters
- Updated Excel generation to use new field names

📈 Financial Industry Standards:
- Transaction fees typically 2-3% for M&A deals
- LTV ratios commonly 60-80% for leveraged transactions
- Professional currency formatting for international deals
- Clear separation of inputs vs calculated values

This creates a comprehensive deal structuring interface with
intelligent financial calculations and professional presentation.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Deal Assumptions updates deployed successfully!"
echo ""
echo "📊 New Deal Structure:"
echo "• Deal Name: Transaction identifier"
echo "• Deal Value: Total transaction value (integer)"
echo "• Transaction Fee (%): Banking fees (default 2.5%)"
echo "• Deal LTV (%): Loan-to-Value ratio (default 70%)"
echo "• Equity Contribution: Auto-calculated from value and LTV"
echo "• Debt Financing: Auto-calculated from value and LTV"
echo ""
echo "💰 Smart Calculations:"
echo "• Equity = Deal Value × (100% - LTV%)"
echo "• Debt = Deal Value × LTV%"
echo "• Real-time updates with currency formatting"
echo "• Integration with High-Level Parameters currency"
echo ""
echo "🎨 Professional Features:"
echo "• Help text for each field"
echo "• Default values for typical deals"
echo "• Currency-aware formatting"
echo "• Readonly calculated fields"
echo ""
echo "🧪 Test the functionality:"
echo "• Enter Deal Value: 100,000,000"
echo "• Set Deal LTV: 70%"
echo "• Verify Equity Contribution: $30,000,000"
echo "• Verify Debt Financing: $70,000,000"
echo "• Try different currencies and values"