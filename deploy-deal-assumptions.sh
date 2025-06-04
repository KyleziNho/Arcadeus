#!/bin/bash

echo "📊 Deploying Deal Assumptions Auto-Fill Extension..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Stage updated files
git add taskpane.js

# Commit the deal assumptions functionality
git commit -m "EXPAND: Add Deal Assumptions auto-fill to existing high-level parameters

📊 DEAL ASSUMPTIONS EXTRACTION: Expand AI auto-fill functionality

🎯 New Extraction Capabilities:
- Deal Name: Extract target company name from document headers
- Deal Value: Calculate from Equity + Debt or find total transaction value  
- Transaction Fee: Extract banking/advisory fees or default to 2.5%
- Deal LTV: Extract leverage ratio or calculate from debt/equity breakdown

💰 Financial Data Processing:
- Calculate deal value: Equity Contribution + Debt Financing
- Extract LTV from 'Acquisition LTV,75%' or calculate from ratios
- Process transaction fees from percentage format
- Extract company names from document headers and deal descriptions

🔧 Enhanced AI Prompt:
- Clear extraction rules for both high-level parameters and deal assumptions
- Specific calculation logic for deal value and LTV
- Fallback defaults for missing transaction fees (2.5%)
- Company name extraction from multiple sources

📋 Expected Results from Your CSV:
HIGH-LEVEL PARAMETERS:
- Currency: USD
- Start Date: 2025-03-31  
- End Date: 2030-03-31
- Model Periods: monthly

DEAL ASSUMPTIONS:
- Deal Name: Sample Company Ltd.
- Deal Value: 100,000,000 (25M equity + 75M debt)
- Transaction Fee: 1.5% (from 'Transaction Fees,1.50%')
- Deal LTV: 75% (from 'Acquisition LTV,75%')

✨ Enhanced User Experience:
- Detailed console logging for each section application
- Comprehensive success summary showing extracted data
- Clear section-by-section processing feedback
- Both sections populate simultaneously from single file upload

🧪 Complete Testing Flow:
1. Upload CSV file with financial data
2. Click 'Auto Fill with AI'
3. Watch console for detailed extraction logs
4. Verify both High-Level Parameters AND Deal Assumptions sections populate
5. Check success summary showing all extracted values

This expansion provides intelligent extraction of key deal parameters
while maintaining the reliable high-level parameters functionality.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Deal Assumptions Auto-Fill Deployed!"
echo ""
echo "📊 Now Extracts & Fills:"
echo "HIGH-LEVEL PARAMETERS:"
echo "• Currency detection"
echo "• Project start/end dates"
echo "• Model periods"
echo ""
echo "DEAL ASSUMPTIONS:"
echo "• Target company name"
echo "• Total deal value calculation"
echo "• Transaction fees"
echo "• Leverage ratio (LTV)"
echo ""
echo "💰 Your CSV Should Extract:"
echo "• Deal Name: Sample Company Ltd."
echo "• Deal Value: $100,000,000"
echo "• Transaction Fee: 1.5%"
echo "• Deal LTV: 75%"
echo ""
echo "🧪 Test Complete Flow:"
echo "1. Upload your CSV file"
echo "2. Click 'Auto Fill with AI'"
echo "3. Check both sections get filled"
echo "4. Verify console shows extraction details"
echo ""
echo "🎯 Both sections should now populate from your single CSV upload!"