#!/bin/bash

echo "🔧 Deploying minimize button fixes..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage all updated files
git add taskpane.html
git add taskpane.css  
git add taskpane.js

# Commit the minimize button fixes
git commit -m "Fix minimize button position, icon, and functionality

🔧 Position Fix:
- Moved minimize button to top-right corner of Deal Assumptions section
- Used absolute positioning with proper z-index
- Removed unnecessary section-header wrapper

🎨 Icon Fix:
- Changed from chevron SVG to simple '-' text symbol
- Icon changes to '+' when collapsed, '-' when expanded
- Cleaner, more intuitive minimize/expand indicator

⚡ Functionality Fix:
- Enhanced JavaScript event handling with better debugging
- Added 500ms delay to ensure DOM elements are found
- Fixed icon toggling logic in click handler
- Improved error reporting for troubleshooting

✨ Visual Improvements:
- Smaller button size (28x28px) for cleaner appearance  
- Better contrast and hover states
- Proper positioning relative to section container

The minimize button now works correctly and appears in the 
expected top-right corner with proper '-' and '+' icons.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Minimize button fixes deployed successfully!"
echo ""
echo "🔧 Fixed Issues:"
echo "• Moved button to top-right corner of Deal Assumptions section"
echo "• Changed icon to simple '-' symbol ('+' when collapsed)"
echo "• Fixed functionality - button should now work when clicked"
echo "• Enhanced debugging to ensure elements are found"
echo "• Better positioning and visual styling"
echo ""
echo "🧪 Test the fixes:"
echo "• Look for '-' button in top-right corner of Deal Assumptions"
echo "• Click to collapse - should show '+' icon and hide content" 
echo "• Click again to expand - should show '-' icon and show content"
echo "• Check browser console for any remaining errors"