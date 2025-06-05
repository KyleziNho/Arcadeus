#!/bin/bash

echo "🔄 Rolling back to commit: Added Revenue Assumption Fix"
echo "Commit hash: 07c4bb0146b511274adefce8ba884a6b7146a5d2"

cd "/Users/kylezinho/Desktop/M&A Plugin 2/excel-addin-hosting"

# Check current status
echo "📋 Current git status:"
git status

echo ""
echo "📋 Current commit:"
git log --oneline -1

echo ""
echo "🔍 Searching for target commit..."
git log --oneline | grep -i "revenue assumption fix" || echo "❌ Commit not found in recent history"

echo ""
echo "⚠️  WARNING: This will reset your working directory to the specified commit."
echo "⚠️  All changes after commit 07c4bb0146b511274adefce8ba884a6b7146a5d2 will be lost!"
echo ""
read -p "Are you sure you want to proceed? (y/N): " -n 1 -r
echo ""

if [[ $REPLY =~ ^[Yy]$ ]]; then
    echo "🔄 Performing hard reset to commit 07c4bb0146b511274adefce8ba884a6b7146a5d2..."
    
    # Hard reset to the specific commit
    git reset --hard 07c4bb0146b511274adefce8ba884a6b7146a5d2
    
    if [ $? -eq 0 ]; then
        echo "✅ Successfully rolled back to commit: Added Revenue Assumption Fix"
        echo ""
        echo "📋 Current commit after rollback:"
        git log --oneline -1
        echo ""
        echo "📁 Files have been restored to their state at that commit."
        echo ""
        echo "🚀 Next steps:"
        echo "1. Verify the files are in the correct state"
        echo "2. If you want to push this rollback, run: git push --force-with-lease origin main"
        echo "   (⚠️  WARNING: This will overwrite the remote repository)"
    else
        echo "❌ Failed to rollback. The commit hash might not exist."
        echo "💡 Try running: git log --oneline to see available commits"
    fi
else
    echo "❌ Rollback cancelled."
fi

echo ""
echo "💡 If you need to see all commits, run: git log --oneline"
echo "💡 If you need to find a specific commit, run: git log --grep=\"Revenue\""