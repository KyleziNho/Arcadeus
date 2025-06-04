#!/bin/bash

echo "📈 Deploying Revenue Items Auto-Fill Extension..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Stage updated files
git add taskpane.js

# Commit the revenue items functionality
git commit -m "EXPAND: Add Revenue Items auto-fill to complete financial data extraction

📈 REVENUE ITEMS EXTRACTION: Add dynamic revenue stream identification and processing

🎯 New Revenue Processing Capabilities:
- Identify all revenue streams from uploaded documents
- Extract revenue names, initial values, and growth patterns
- Dynamically create matching number of revenue items in UI
- Determine growth types: linear, no_growth, or nonlinear
- Apply growth rates for linear growth scenarios

💰 Revenue Data Intelligence:
- Search for 'Revenue', 'Sales', 'Income', 'Turnover' in documents
- Extract specific revenue categories and product lines
- Process current/base year revenue amounts
- Analyze growth patterns and rates from financial projections
- Handle multiple revenue streams with different growth characteristics

🔧 Dynamic UI Management:
- Clear existing revenue items before applying new data
- Create additional revenue items based on extracted count
- Use existing applyRevenueItems function for proper field mapping
- Set revenue names, initial values, growth types, and rates
- Maintain proper revenue item IDs and counters

📋 Growth Type Classification:
- LINEAR: Consistent growth rate (e.g., 5% annually)
- NO_GROWTH: Flat/stable revenue (0% growth)  
- NONLINEAR: Varying growth rates over time periods

🧠 Enhanced AI Extraction Logic:
- Comprehensive revenue stream identification
- Business type-based revenue categorization
- Growth pattern analysis from historical/projected data
- Smart naming conventions for revenue streams
- Conversion of percentages and monetary values

✨ Complete Auto-Fill Flow:
HIGH-LEVEL PARAMETERS:
- Currency, dates, model periods

DEAL ASSUMPTIONS:  
- Deal name, value, transaction fees, LTV

REVENUE ITEMS (NEW):
- Dynamic revenue stream creation
- Growth type and rate application
- Multiple revenue categories support

🧪 Expected Results from Financial Documents:
- All revenue streams identified and created as separate items
- Proper growth classification based on document analysis
- Revenue values extracted and formatted correctly
- Growth rates applied to linear growth scenarios
- Complete revenue section population from single file upload

This completes the core financial data extraction covering all
primary deal parameters and revenue modeling requirements.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Revenue Items Auto-Fill Deployed!"
echo ""
echo "📈 Now Extracts & Fills Three Sections:"
echo "HIGH-LEVEL PARAMETERS:"
echo "• Currency, dates, model periods"
echo ""
echo "DEAL ASSUMPTIONS:"
echo "• Deal name, value, fees, LTV"
echo ""
echo "REVENUE ITEMS (NEW):"
echo "• Dynamic revenue stream creation"
echo "• Growth type classification"
echo "• Initial values and growth rates"
echo ""
echo "🎯 Revenue Processing Features:"
echo "• Identifies all revenue streams in document"
echo "• Creates matching number of revenue items"
echo "• Determines linear/no_growth/nonlinear patterns"
echo "• Applies growth rates for projections"
echo ""
echo "🧪 Test Complete Auto-Fill:"
echo "1. Upload your CSV file"
echo "2. Click 'Auto Fill with AI'"
echo "3. Check all three sections populate"
echo "4. Verify revenue items match document data"
echo ""
echo "📊 Expected Success:"
echo "All financial data sections should now auto-populate from your CSV!"