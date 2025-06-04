#!/bin/bash

echo "🎨 Deploying modern selection styling - removing sidebars for full button highlights..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated CSS file
git add taskpane.css

# Commit the modern selection styling
git commit -m "Modern selection styling - full button highlights instead of sidebars

🎨 Radio Button Modern Selection:
- Removed blue sidebar indicators inside buttons
- Full button background turns blue when selected
- White text on blue background for better contrast
- White circular indicator instead of blue
- Clean, modern card-style selection

✨ Checkbox Modern Selection:
- Removed blue sidebar indicators inside buttons
- Full button background turns blue when selected
- White text on blue background for selected state
- White square indicator instead of blue
- Consistent with radio button styling

🔧 Technical Implementation:
- Uses :has() pseudo-selector for parent container styling
- Blue background with white text for selected state
- White indicators with subtle shadows for contrast
- Maintains all hover animations and effects
- Enhanced box-shadow for selected state depth

🎯 User Experience:
- Clear visual feedback with full button highlighting
- Better contrast and readability when selected
- More modern and cohesive design language
- Consistent selection pattern across all form elements

The interface now uses modern full-button highlighting
instead of small sidebar indicators for better UX.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Modern selection styling deployed successfully!"
echo ""
echo "🎨 Selection Style Changes:"
echo "• Removed blue sidebar indicators inside buttons"
echo "• Full button background turns blue when selected"
echo "• White text and indicators on blue background"
echo "• Enhanced shadows for depth and modern feel"
echo "• Consistent styling across radio buttons and checkboxes"
echo ""
echo "✨ Visual Result:"
echo "• Radio buttons: Entire button turns blue with white circle"
echo "• Checkboxes: Entire button turns blue with white square"
echo "• Better contrast and readability"
echo "• Modern, card-based selection pattern"
echo "• Clean, professional appearance"