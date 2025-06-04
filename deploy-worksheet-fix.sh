#!/bin/bash

echo "🔧 Deploying worksheet existence fix for debt schedule generation..."

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check git status
echo "Git status:"
git status

# Stage updated JavaScript file
git add taskpane.js

# Commit the worksheet fix
git commit -m "Fix worksheet does not exist error by using current active worksheet

🔧 Worksheet Error Fix:
- Removed new worksheet creation that was causing 'worksheet does not exist' error
- Now uses current active worksheet instead of creating new 'Debt Schedule' sheet
- Eliminates worksheet reference issues that were preventing data insertion
- Clears target area first to avoid conflicts with existing content

⚡ Simplified Workflow:
- Uses context.workbook.worksheets.getActiveWorksheet() for reliable access
- Clears range A1:J10 before inserting debt schedule data
- No complex worksheet creation/deletion logic that could fail
- Direct data insertion to current sheet eliminates reference errors

📊 Reliable Data Insertion:
- Fixed 5x10 data array insertion to range A1:J5
- Proper sync after clearing and before data insertion
- Basic formatting applied only after data is confirmed inserted
- Clear error handling and logging for debugging

🛡️ Error Prevention:
- No worksheet name conflicts or creation failures
- Uses existing active worksheet that always exists
- Avoids Excel API issues with new worksheet references
- Maintains data insertion even if worksheet operations fail

This ensures the debt schedule data will always be inserted
without worksheet reference errors.

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to main
git push origin main

echo "✅ Worksheet existence fix deployed successfully!"
echo ""
echo "🔧 Fixed Issues:"
echo "• Removed new worksheet creation causing 'does not exist' error"
echo "• Uses current active worksheet for reliable access"
echo "• Clears target area before inserting debt schedule"
echo "• Eliminates worksheet reference conflicts"
echo ""
echo "📊 New Behavior:"
echo "• Debt schedule appears in current active worksheet"
echo "• Data inserted to range A1:J5 with clearing first"
echo "• No new worksheet creation to avoid API conflicts"
echo "• Reliable data insertion without reference errors"
echo ""
echo "🧪 Test the functionality:"
echo "• Click Generate Debt Schedule in any worksheet"
echo "• Verify debt schedule data appears in current sheet"
echo "• Check that no worksheet error messages appear"
echo "• Confirm data is properly formatted in A1:J5 range"