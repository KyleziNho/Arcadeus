#!/usr/bin/env bash
set -e

echo "Starting deployment from $(pwd)"

# Change to the correct directory
cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

echo "Changed to directory: $(pwd)"

# Check status
echo "=== Git Status ==="
/usr/bin/git status

# Add the file
echo "=== Adding taskpane.js ==="
/usr/bin/git add taskpane.js

# Commit
echo "=== Creating commit ==="
/usr/bin/git commit -m "Add collapsible Deal Assumptions functionality

- Added initializeCollapsibleSections() method to MAModelingAddin class
- Added event listener for minimize button to toggle 'collapsed' class  
- Added accessibility support with dynamic aria-label updates
- CSS animations for collapsible sections were implemented in previous sessions

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push
echo "=== Pushing to origin main ==="
/usr/bin/git push origin main

echo "=== Deployment completed successfully ==="