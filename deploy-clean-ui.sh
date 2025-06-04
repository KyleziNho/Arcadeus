#!/bin/bash

echo "🧹 Deploying clean UI fixes - removing inner circles and symbols..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated CSS file
git add taskpane.css

# Commit the clean UI fixes
git commit -m "Clean UI fixes - remove white circles and checkmark symbols

🧹 Radio Button Clean-up:
- Removed white inner circles from selected radio buttons
- Clean solid blue fill when selected
- Maintains hover animations and shadows
- Simpler, more modern appearance

✨ Checkbox Clean-up:
- Removed checkmark symbols from selected checkboxes
- Clean solid blue fill when selected
- Maintains card-style hover effects and animations
- Cleaner, more minimalist design

🎨 Visual Improvements:
- Cleaner selected states without inner elements
- More consistent with modern UI design patterns
- Maintains all hover and transition effects
- Professional, minimalist appearance

The interface now has clean, solid-fill selections without
distracting inner symbols or circles.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Clean UI fixes deployed successfully!"
echo ""
echo "🧹 UI Improvements:"
echo "• Removed white circles from radio button centers"
echo "• Removed checkmark symbols from checkbox centers" 
echo "• Clean solid blue fill when selected"
echo "• Maintains all hover animations and effects"
echo "• More minimalist, modern appearance"
echo ""
echo "🎨 The result:"
echo "• Radio buttons: Clean blue circles when selected"
echo "• Checkboxes: Clean blue squares when selected"
echo "• No inner symbols or distracting elements"
echo "• Professional, Apple-style interface"