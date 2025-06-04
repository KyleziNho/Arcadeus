#!/bin/bash

# Navigate to the correct directory
cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

echo "Current directory: $(pwd)"

# Check git status
echo "=== Git Status ==="
git status

# Stage the taskpane.js file
echo "=== Staging taskpane.js ==="
git add taskpane.js

# Commit with message about collapsible functionality
echo "=== Creating commit ==="
git commit -m "Add collapsible Deal Assumptions functionality

- Added initializeCollapsibleSections() method to MAModelingAddin class
- Added event listener for minimize button to toggle 'collapsed' class  
- Added accessibility support with dynamic aria-label updates
- CSS animations for collapsible sections were implemented in previous sessions

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to origin main
echo "=== Pushing to origin main ==="
git push origin main

echo "=== Deployment complete ==="