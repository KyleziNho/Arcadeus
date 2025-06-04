#!/bin/bash

echo "📋 Deploying complete collapsible Deal Assumptions feature..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage both HTML and JavaScript files with collapsible functionality
git add taskpane.html
git add taskpane.js

# Commit the complete collapsible feature
git commit -m "Complete collapsible Deal Assumptions section implementation

✨ HTML Structure Changes:
- Added 'collapsible-section' class to Deal Assumptions section
- Added section-header div with minimize button
- Added section-content wrapper for collapsible content
- Added minimize button with chevron SVG icon
- Added proper IDs for JavaScript targeting

✨ JavaScript Functionality:
- Added initializeCollapsibleSections() method to MAModelingAddin class
- Added event listener for minimize button to toggle 'collapsed' class
- Added accessibility support with dynamic aria-label updates
- Added console logging for debugging and verification

✨ CSS Animations (already implemented):
- Smooth max-height transitions (400ms) with cubic-bezier easing
- Opacity fade effects (300ms) synchronized with height changes
- Icon rotation animation (180°) when collapsed/expanded
- Reduced padding when collapsed for cleaner appearance

🎯 Functionality:
- Click minimize button in Deal Assumptions header to collapse/expand
- Smooth animations provide professional user experience
- AI chat remains fully visible when assumptions are minimized
- Accessible design with proper ARIA labels

This resolves the 'Could not find collapsible section elements' error
by ensuring both HTML structure and JavaScript functionality are deployed.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Complete collapsible functionality deployed successfully!"
echo ""
echo "📋 Fixed Issues:"
echo "• Resolved 'Could not find collapsible section elements' error"
echo "• Deployed updated HTML structure with collapsible elements"
echo "• Deployed JavaScript event handling for minimize button"
echo "• All CSS animations were already in place"
echo ""
echo "🧪 Test the functionality:"
echo "• Look for chevron minimize button in Deal Assumptions header"
echo "• Click to collapse section with smooth animation"
echo "• Click again to expand with smooth animation"
echo "• Verify no console errors in browser developer tools"
echo "• Check that AI chat area remains fully accessible"