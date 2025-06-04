#!/bin/bash

echo "📈 Deploying enhanced Revenue Items with period-linked growth and click-to-expand..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated files
git add taskpane.html
git add taskpane.css
git add taskpane.js

# Commit the enhanced revenue features
git commit -m "Enhance Revenue Items with period-linked non-linear growth and click-to-expand

📈 Smart Period-Linked Non-Linear Growth:
- Links non-linear growth to High-Level Parameters period selection
- Uses actual calculated holding periods from project dates
- Adapts to selected period type (daily, monthly, quarterly, yearly)
- Dynamic period labeling based on user selection

🎯 Intelligent Growth Input System:
- ≤12 periods: Individual period-by-period growth input fields
- >12 periods: Grouped period ranges with add/remove functionality
- Smart suggestions for period ranges (e.g., Months 1-12, then 13-24)
- Automatic period calculations from High-Level Parameters

📊 Advanced Period Grouping (>12 periods):
- Group 1: Default periods 1-12 with customizable growth rate
- Add Group: Creates new range starting from last period + 1
- Remove Groups: Dynamic group management with smart re-indexing
- Example: Months 1-12 at 1%, Months 13-36 at 0.5%

🖱️ Click-to-Expand Minimized Sections:
- Click anywhere on collapsed section to expand it
- Prevents expansion when clicking minimize button
- Smooth cursor change indicators (pointer for collapsed, default for expanded)
- Works for all sections: High-Level Parameters, Deal Assumptions, Revenue Items, Debt Model

🎨 Professional Period Group UI:
- Clean bordered containers for each period group
- Grid layout for From/To/Growth Rate inputs
- Green '+ Add Group' and red '× Remove' buttons
- Smart button management (add button moves to last group)
- Consistent styling with existing section design

🔧 Technical Integration:
- Real-time period calculation extraction from High-Level Parameters
- Dynamic DOM manipulation for period group management
- Event delegation for scalable functionality
- Automatic period labeling (Day/Month/Quarter/Year)
- Smart default value suggestions for new groups

⚡ Enhanced User Experience:
- Contextual help text with period-specific examples
- Intuitive grouped input for complex growth scenarios
- Visual feedback with cursor changes for collapsed sections
- Professional color coding (green for add, red for remove)

🧮 Financial Modeling Power:
- Supports complex multi-phase growth scenarios
- Period-accurate modeling based on actual project timeline
- Flexible grouping for sophisticated financial projections
- Ready for Excel generation with detailed growth data

This creates a comprehensive revenue modeling system that adapts
to user parameters and provides professional-grade flexibility.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Enhanced Revenue Items features deployed successfully!"
echo ""
echo "📈 Period-Linked Non-Linear Growth:"
echo "• Links to High-Level Parameters period selection"
echo "• Uses calculated holding periods from project dates"
echo "• Adapts labels (Day/Month/Quarter/Year) to selection"
echo "• ≤12 periods: Individual period inputs"
echo "• >12 periods: Grouped period ranges"
echo ""
echo "🎯 Smart Period Grouping Features:"
echo "• Default Group 1: Periods 1-12"
echo "• Add Group: Smart period range suggestions"
echo "• Remove Groups: Dynamic management"
echo "• Example: Months 1-12 at 1%, Months 13-36 at 0.5%"
echo ""
echo "🖱️ Click-to-Expand Functionality:"
echo "• Click collapsed sections to expand them"
echo "• Visual cursor feedback (pointer/default)"
echo "• Works for all collapsible sections"
echo "• Prevents expansion when clicking minimize button"
echo ""
echo "🎨 Professional UI Enhancements:"
echo "• Clean period group containers"
echo "• Grid layout for From/To/Growth inputs"
echo "• Color-coded action buttons (green add, red remove)"
echo "• Smart button placement and management"
echo ""
echo "🧪 Test the functionality:"
echo "• Set different period types in High-Level Parameters"
echo "• Create revenue items with >12 periods for grouping"
echo "• Try click-to-expand on minimized sections"
echo "• Add/remove period groups and verify smart suggestions"