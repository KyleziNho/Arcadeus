#!/bin/bash

echo "🗑️ Deploying sidebar removal - hiding radio/checkbox indicators completely..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated CSS file
git add taskpane.css

# Commit the sidebar removal fix
git commit -m "Remove sidebar indicators completely from selected buttons

🗑️ Complete Sidebar Removal:
- Hidden radio-custom elements when radio buttons are selected
- Hidden checkbox-custom elements when checkboxes are selected
- Clean button selection with no visible indicators
- Full button background highlighting only

✨ Visual Result:
- No more blue circles/squares visible in selected buttons
- Pure text-only appearance when selected
- Clean blue background with white text
- Modern, minimalist selection style

🔧 Technical Fix:
- Used display: none on .radio-custom and .checkbox-custom when selected
- Maintains full button background highlighting
- Removes all visual indicators from button interiors
- Clean, text-only selected state

The buttons now have completely clean selected states
with only background color and text color changes.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Sidebar indicators removed completely!"
echo ""
echo "🗑️ Changes Made:"
echo "• Hidden all radio button circles when selected"
echo "• Hidden all checkbox squares when selected" 
echo "• Clean text-only appearance in selected state"
echo "• Pure blue background with white text"
echo ""
echo "✨ Final Result:"
echo "• Selected buttons show only text on blue background"
echo "• No visible circles, squares, or sidebar indicators"
echo "• Modern, minimalist selection style"
echo "• Clean, professional appearance"