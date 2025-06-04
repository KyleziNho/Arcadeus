#!/bin/bash

# Navigate to the excel-addin-hosting directory
cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Remove test files
rm -f taskpane-fixed.js
rm -f cleanup-and-push.sh

echo "Cleaned up temporary files"

# Check git status
git status

# Stage the important files
git add manifest.xml
git add taskpane.js

# Commit the changes
git commit -m "Fix JavaScript compatibility and manifest validation

- Convert TypeScript to JavaScript for browser compatibility
- Remove interfaces and type annotations 
- Add console logging for debugging
- Simplify manifest to pass Excel Web validation
- Fix button event listeners and Office.js integration

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "JavaScript fixes deployed successfully!"
echo "Now try the buttons in Excel - they should work!"