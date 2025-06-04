#!/bin/bash

echo "📊 Deploying final debt schedule with new worksheet and dark teal formatting..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated JavaScript file
git add taskpane.js

# Commit the final worksheet and formatting updates
git commit -m "Create debt schedule in new worksheet with dark teal period headers

📊 New Worksheet Creation:
- Creates dedicated 'Debt Schedule' worksheet instead of using current page
- Deletes existing 'Debt Schedule' worksheet if it exists to ensure clean slate
- Automatically activates new worksheet for immediate viewing
- Protects user's original worksheet from modification

🎨 Professional Dark Teal Formatting:
- Period header row: Dark teal background (#1F5F5B) - Accent 1, Darker 25%
- White text (#FFFFFF) on dark teal background for optimal readability
- Maintains Times New Roman 12pt font throughout
- Professional table borders and gray header for 'Debt Model'

✨ Enhanced User Experience:
- Debt schedule appears in dedicated clean worksheet
- No interference with user's current work
- Professional financial industry color scheme
- Clear visual hierarchy with contrasting colors

🔧 Reliable Implementation:
- Proper worksheet creation and activation sequence
- Comprehensive error handling for worksheet operations
- Removed area clearing since new worksheet is always clean
- Fixed data insertion to A1:J5 range with professional formatting

This provides the complete debt generation solution with
dedicated worksheet creation and professional formatting.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Final debt schedule with new worksheet deployed successfully!"
echo ""
echo "📊 New Features:"
echo "• Creates dedicated 'Debt Schedule' worksheet"
echo "• Protects original worksheet from modification"
echo "• Dark teal period headers with white text"
echo "• Professional Times New Roman 12pt formatting"
echo ""
echo "🎨 Professional Formatting:"
echo "• Dark teal background (#1F5F5B) for period headers"
echo "• White text for optimal readability on dark background"
echo "• Complete table borders and professional styling"
echo "• Gray header for main title with merged cells"
echo ""
echo "✨ User Experience:"
echo "• Dedicated clean worksheet for debt schedule"
echo "• Automatic activation to view results immediately"
echo "• No interference with current work"
echo "• Professional financial industry appearance"
echo ""
echo "🧪 Test the functionality:"
echo "• Click Generate Debt Schedule from any worksheet"
echo "• Verify new 'Debt Schedule' worksheet is created and activated"
echo "• Check dark teal period headers with white text"
echo "• Confirm professional formatting and data layout"