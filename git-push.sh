#!/bin/bash

echo "🚀 Pushing robust Excel Web fixes to main..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Show status
echo "Git status:"
git status

# Stage changes
git add taskpane.js

# Remove temp files
rm -f fix-and-deploy.sh
rm -f push-to-main.sh
rm -f deploy-robust-fixes.sh

# Commit
git commit -m "Add robust Excel Web compatibility and error handling

- Add comprehensive Office.js availability checking
- Implement fallback initialization for Excel Web environment  
- Add DOM readiness checks before element access
- Include global error handlers for better debugging
- Add initialization state tracking to prevent double-init
- Improve Excel API error handling with user feedback
- Add detailed console logging for troubleshooting

Fixes Excel Web platform compatibility issues and provides
better user feedback when features are unavailable.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Changes pushed to main successfully!"