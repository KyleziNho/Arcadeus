#!/bin/bash

echo "💸 Deploying Cost Items section with identical features to Revenue Items..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated files
git add taskpane.html
git add taskpane.css
git add taskpane.js

# Commit the Cost Items feature
git commit -m "Add comprehensive Cost Items section mirroring Revenue Items functionality

💸 Cost Items Management (identical to Revenue Items):
- New collapsible section with professional card-based layout
- One required cost item (cannot be removed)
- Add/remove additional cost items with dynamic functionality
- Clean card design with headers and remove buttons for optional items

💰 Cost Item Configuration:
- Cost Item Name: Text input for cost category identification
- Initial Value: Numeric input for starting cost amount
- Currency integration with High-Level Parameters selection
- Professional help text for user guidance

🎯 Flexible Growth Types (identical to Revenue):
1. No Growth: Costs remain constant over time
2. Linear Growth: Single annual growth rate (positive or negative)
3. Non-Linear Growth: Period-by-period specific growth rates

📊 Smart Non-Linear Growth (period-linked):
- Automatically uses project start/end dates from High-Level Parameters
- Generates period-specific input fields based on project timeline
- ≤12 periods: Individual period-by-period growth input fields
- >12 periods: Grouped period ranges with add/remove functionality
- Smart suggestions for period ranges (e.g., Months 1-12, then 13-24)
- Automatic period calculations from High-Level Parameters

🎨 Professional UI Design (consistent with Revenue Items):
- Card-based layout with bordered containers
- Grid layout for optimal space utilization
- Red remove buttons with hover effects
- Organized input groupings with clear labeling
- Responsive design for different screen sizes

🔧 Technical Features:
- Complete JavaScript implementation with initializeCostItems()
- setupCostItemListeners() and updateCostGrowthInputs() functions
- addCostPeriodGroup() for advanced period grouping functionality
- Dynamic DOM manipulation for add/remove functionality
- Event delegation for scalable event handling
- Unique ID system for multiple cost items
- Real-time growth input updates based on type selection
- Integration with project timeline from High-Level Parameters

⚡ User Experience:
- Intuitive add/remove workflow
- Clear visual distinction between required and optional items
- Contextual help text for each input type (cost-specific)
- Smooth transitions and professional animations
- Consistent styling with other sections
- Click-to-expand functionality for collapsed sections

🖱️ Enhanced Collapsible Functionality:
- Cost Items section fully integrated into collapsible system
- Click anywhere on collapsed section to expand it
- Prevents expansion when clicking minimize button
- Smooth cursor change indicators (pointer for collapsed, default for expanded)
- Professional minimize/expand icon transitions

🧮 Cost Modeling Features:
- Structured data collection for Excel generation
- Support for complex cost growth scenarios
- Period-by-period granular control for detailed planning
- Both linear and non-linear modeling capabilities
- Professional cost escalation modeling

📈 Period-Linked Non-Linear Cost Growth:
- Links non-linear growth to High-Level Parameters period selection
- Uses actual calculated holding periods from project dates
- Adapts to selected period type (daily, monthly, quarterly, yearly)
- Dynamic period labeling based on user selection
- Group 1: Default periods 1-12 with customizable growth rate
- Add Group: Creates new range starting from last period + 1
- Remove Groups: Dynamic group management with smart re-indexing
- Example: Months 1-12 at 3% cost increase, Months 13-36 at 2%

This creates a comprehensive cost modeling system that perfectly
mirrors the Revenue Items functionality and provides professional-grade
flexibility for sophisticated financial planning and analysis.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Cost Items section deployed successfully!"
echo ""
echo "💸 Cost Item Features (identical to Revenue Items):"
echo "• One required cost item (cannot be removed)"
echo "• Add unlimited additional cost items"
echo "• Cost Item Name and Initial Value inputs"
echo "• Currency integration with High-Level Parameters"
echo ""
echo "🎯 Growth Type Options:"
echo "• No Growth: Constant costs over time"
echo "• Linear Growth: Single annual growth rate (+ or -)"
echo "• Non-Linear Growth: Period-by-period specific rates"
echo ""
echo "📊 Smart Non-Linear Features:"
echo "• Uses project dates for period-specific inputs"
echo "• ≤12 periods: Individual period inputs"
echo "• >12 periods: Grouped period ranges"
echo "• Supports sophisticated cost escalation modeling"
echo ""
echo "🎨 Professional Design:"
echo "• Card-based layout with clean styling"
echo "• Red remove buttons for optional items"
echo "• Grid layout for optimal organization"
echo "• Consistent with existing sections"
echo ""
echo "🖱️ Click-to-Expand Functionality:"
echo "• Click collapsed sections to expand them"
echo "• Visual cursor feedback (pointer/default)"
echo "• Works for all collapsible sections including Cost Items"
echo "• Prevents expansion when clicking minimize button"
echo ""
echo "🧪 Test the functionality:"
echo "• Add/remove cost items"
echo "• Try different growth types"
echo "• Set project dates and see non-linear adapt"
echo "• Verify required item cannot be removed"
echo "• Test period grouping for >12 periods"
echo "• Try click-to-expand on minimized Cost Items section"