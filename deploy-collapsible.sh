#!/bin/bash

echo "📋 Deploying collapsible Deal Assumptions feature..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage the updated JavaScript file with collapsible functionality
git add taskpane.js

# Commit the collapsible feature
git commit -m "Add collapsible Deal Assumptions section with smooth animations

✨ New Collapsible Feature:
- Added minimize/expand button to Deal Assumptions section
- Smooth CSS transitions with max-height and opacity animations
- Chevron icon rotates 180° when collapsed/expanded
- Maintains visibility of AI chat when assumptions are minimized
- Accessible with proper ARIA labels for screen readers

🎨 Animation Details:
- 400ms max-height transition with cubic-bezier easing
- 300ms opacity fade with synchronized timing
- Icon rotation animation on state change
- Reduced padding when collapsed for cleaner appearance

🛠 Technical Implementation:
- JavaScript event handling for minimize button clicks
- CSS class toggle for 'collapsed' state
- Accessibility improvements with dynamic aria-label updates
- Console logging for debugging and verification

The Deal Assumptions section can now be hidden to provide more
space for the AI chat interface while maintaining easy access
to expand when needed.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Collapsible functionality deployed successfully!"
echo ""
echo "📋 New Collapsible Features:"
echo "• Minimize/expand button in Deal Assumptions header"
echo "• Smooth animation transitions (400ms max-height, 300ms opacity)"
echo "• Chevron icon rotates when collapsed/expanded"
echo "• AI chat remains fully visible when assumptions minimized"
echo "• Accessible with proper ARIA labels"
echo "• Console logging for debugging"
echo ""
echo "🧪 Test the new functionality:"
echo "• Click the minimize button (chevron icon) in Deal Assumptions header"
echo "• Watch the smooth collapse animation"
echo "• Notice the icon rotation and reduced section height"
echo "• Click again to expand with smooth animation"
echo "• Verify AI chat area remains fully accessible"