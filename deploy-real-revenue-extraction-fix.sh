#!/bin/bash

echo "📸 Deploying REAL Revenue Data Extraction with PNG Support..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Stage updated files
git add taskpane.html
git add taskpane.js

# Commit the real revenue extraction with PNG support
git commit -m "REAL DATA: Extract actual revenue items from uploaded files with PNG support

📸 PNG/IMAGE SUPPORT: Added support for financial data screenshots

🎯 New File Format Support:
- Added PNG, JPG, JPEG file upload support
- Updated file acceptance to include image formats
- Enhanced UI to show PNG file format support
- Image processing instructions for AI analysis

📊 REAL Revenue Data Extraction (Based on User Example):
From actual uploaded data showing:
- Revenue Item 1: 500,000 with Rent Growth 1: 2.00%
- Revenue Item 2: 766,000 with Rent Growth 2: 3.00%

Should extract EXACTLY:
- Revenue Stream 1: 500,000 (linear growth 2%)
- Revenue Stream 2: 766,000 (linear growth 3%)

🧠 Enhanced AI Revenue Logic:
1. PRIORITY: Look for explicit 'Revenue Item 1', 'Revenue Item 2' patterns
2. EXTRACT: Actual values (500,000, 766,000) from document
3. MATCH: Growth rates by number (Rent Growth 1 → Revenue Item 1)
4. CONDITIONAL: If no revenue items found → return empty array []
5. NO GUESSING: Don't create fake revenue items without explicit data

📋 Smart Revenue Processing:
- Match 'Rent Growth 1: 2.00%' to 'Revenue Item 1: 500,000'
- Match 'Rent Growth 2: 3.00%' to 'Revenue Item 2: 766,000'
- Convert percentages: '2.00%' → growthRate: 2
- Use business context for meaningful names (Real Estate → rental income)

🖼️ Image/PNG Processing:
- Detect image files and provide AI analysis instructions
- Guide AI to look for specific revenue item patterns in screenshots
- Handle visual financial data extraction
- Support for financial dashboard screenshots

✅ Conditional Logic:
- IF revenue items found → Extract exact data and create items
- IF no revenue items found → Leave Revenue Items section empty
- Enhanced console logging to show extraction decision process

🔧 File Format Support:
- CSV: Direct content analysis
- PDF: Contextual analysis (placeholder for future OCR)
- PNG/JPG: Visual content analysis instructions
- Up to 4 files, 10MB total limit

This ensures the AI extracts REAL revenue data when present
and leaves sections empty when no data exists, matching
the user's specific requirements and data format.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Real Revenue Data Extraction with PNG Support Deployed!"
echo ""
echo "📸 New Features:"
echo "• PNG/JPG/JPEG file upload support"
echo "• Image-based financial data extraction"
echo "• Real revenue item pattern recognition"
echo "• Conditional extraction (empty if no data)"
echo ""
echo "📊 Based on Your Example Data:"
echo "Revenue Items Section should show:"
echo "• Revenue Item 1: 500,000 (2% growth)"
echo "• Revenue Item 2: 766,000 (3% growth)"
echo "• Real Estate rental income context"
echo ""
echo "🧪 Test with Your Data:"
echo "1. Upload PNG screenshot or CSV file"
echo "2. Click 'Auto Fill with AI'"
echo "3. Check console for extraction decisions"
echo "4. Verify exact revenue data appears or section stays empty"
echo ""
echo "💡 AI Now Looks For:"
echo "• 'Revenue Item 1', 'Revenue Item 2' patterns"
echo "• 'Rent Growth 1', 'Rent Growth 2' percentages"
echo "• Exact value matching and growth rate application"
echo "• Business context for meaningful revenue names"