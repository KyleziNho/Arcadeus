#!/bin/bash

# Navigate to the excel-addin-hosting directory
cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Remove test XML files
rm -f manifest-debug.xml
rm -f manifest-simple.xml
rm -f manifest-test1.xml
rm -f commit-changes.sh

echo "Removed test files"

# Check git status
git status

# Stage only the manifest.xml changes
git add manifest.xml

# Commit the changes
git commit -m "Fix manifest validation for Excel Web compatibility

- Remove VersionOverrides section causing validation failures
- Clean up HTML entities in text fields
- Generate new unique ID to avoid conflicts  
- Fix support URL from placeholder to actual domain

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "Changes pushed to main successfully!"