#!/bin/bash

echo "📈 Deploying Revenue Items section with flexible growth modeling..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated files
git add taskpane.html
git add taskpane.css
git add taskpane.js

# Commit the Revenue Items feature
git commit -m "Add comprehensive Revenue Items section with flexible growth modeling

📈 Revenue Items Management:
- New collapsible section with professional card-based layout
- One required revenue item (cannot be removed)
- Add/remove additional revenue items with dynamic functionality
- Clean card design with headers and remove buttons for optional items

💰 Revenue Item Configuration:
- Revenue Source Name: Text input for stream identification
- Initial Value: Numeric input for starting revenue amount
- Currency integration with High-Level Parameters selection
- Professional help text for user guidance

🎯 Flexible Growth Types:
1. No Growth: Revenue remains constant over time
2. Linear Growth: Single annual growth rate (positive or negative)
3. Non-Linear Growth: Year-by-year specific growth rates

📊 Smart Non-Linear Growth:
- Automatically uses project start/end dates from High-Level Parameters
- Generates year-specific input fields based on project timeline
- Fallback to 3-year simple inputs if dates not configured
- Up to 5 years of detailed year-by-year growth planning

🎨 Professional UI Design:
- Card-based layout with bordered containers
- Grid layout for optimal space utilization
- Red remove buttons with hover effects
- Organized input groupings with clear labeling
- Responsive design for different screen sizes

🔧 Technical Features:
- Dynamic DOM manipulation for add/remove functionality
- Event delegation for scalable event handling
- Unique ID system for multiple revenue items
- Real-time growth input updates based on type selection
- Integration with project timeline from High-Level Parameters

⚡ User Experience:
- Intuitive add/remove workflow
- Clear visual distinction between required and optional items
- Contextual help text for each input type
- Smooth transitions and professional animations
- Consistent styling with other sections

🧮 Growth Calculation Ready:
- Structured data collection for Excel generation
- Support for complex growth scenarios
- Year-by-year granular control
- Both linear and non-linear modeling capabilities

This creates a comprehensive revenue modeling system suitable
for sophisticated financial planning and analysis.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Revenue Items section deployed successfully!"
echo ""
echo "📈 Revenue Item Features:"
echo "• One required revenue item (cannot be removed)"
echo "• Add unlimited additional revenue items"
echo "• Revenue Source Name and Initial Value inputs"
echo "• Currency integration with High-Level Parameters"
echo ""
echo "🎯 Growth Type Options:"
echo "• No Growth: Constant revenue over time"
echo "• Linear Growth: Single annual growth rate (+ or -)"
echo "• Non-Linear Growth: Year-by-year specific rates"
echo ""
echo "📊 Smart Non-Linear Features:"
echo "• Uses project dates for year-specific inputs"
echo "• Supports up to 5 years of detailed planning"
echo "• Fallback to 3-year simple inputs"
echo "• Positive/negative growth rate support"
echo ""
echo "🎨 Professional Design:"
echo "• Card-based layout with clean styling"
echo "• Red remove buttons for optional items"
echo "• Grid layout for optimal organization"
echo "• Consistent with existing sections"
echo ""
echo "🧪 Test the functionality:"
echo "• Add/remove revenue items"
echo "• Try different growth types"
echo "• Set project dates and see non-linear adapt"
echo "• Verify required item cannot be removed"