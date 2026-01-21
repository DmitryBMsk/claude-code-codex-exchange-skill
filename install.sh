#!/bin/bash
#
# Claude Exchange Mail - Installation Script
#
# This script installs the Exchange Mail skill for Claude Code.
#
# Usage: ./install.sh
#

set -e

echo "📧 Installing Claude Exchange Mail..."

# Check Python version
PYTHON_VERSION=$(python3 --version 2>&1 | cut -d' ' -f2 | cut -d'.' -f1,2)
REQUIRED_VERSION="3.10"

if [ "$(printf '%s\n' "$REQUIRED_VERSION" "$PYTHON_VERSION" | sort -V | head -n1)" != "$REQUIRED_VERSION" ]; then
    echo "❌ Python $REQUIRED_VERSION+ is required. Found: $PYTHON_VERSION"
    exit 1
fi

echo "✓ Python version: $PYTHON_VERSION"

# Install dependencies
echo "Installing Python dependencies..."
pip3 install -q exchangelib

# Create directories
mkdir -p ~/.claude/skills/exchange-mail/scripts

# Copy files
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

cp "$SCRIPT_DIR/skills/exchange-mail/scripts/exchange_mail.py" ~/.claude/exchange_mail.py
cp "$SCRIPT_DIR/skills/exchange-mail/SKILL.md" ~/.claude/skills/exchange-mail/SKILL.md

chmod +x ~/.claude/exchange_mail.py

echo "✓ Files installed"

# Check environment variables
MISSING_VARS=""
[ -z "$EXCHANGE_SERVER" ] && MISSING_VARS="$MISSING_VARS EXCHANGE_SERVER"
[ -z "$EXCHANGE_EMAIL" ] && MISSING_VARS="$MISSING_VARS EXCHANGE_EMAIL"
[ -z "$EXCHANGE_USERNAME" ] && MISSING_VARS="$MISSING_VARS EXCHANGE_USERNAME"
[ -z "$EXCHANGE_PASSWORD" ] && MISSING_VARS="$MISSING_VARS EXCHANGE_PASSWORD"

if [ -n "$MISSING_VARS" ]; then
    echo ""
    echo "⚠️  Missing environment variables:$MISSING_VARS"
    echo ""
    echo "Add to your ~/.zshrc or ~/.bashrc:"
    echo ""
    echo '  export EXCHANGE_SERVER="mail.yourcompany.com"'
    echo '  export EXCHANGE_EMAIL="your.email@company.com"'
    echo '  export EXCHANGE_USERNAME="your_username"'
    echo '  export EXCHANGE_PASSWORD="your_password"'
    echo ""
    echo "Then run: source ~/.zshrc"
else
    echo "✓ Environment variables configured"
fi

echo ""
echo "✅ Installation complete!"
echo ""
echo "Usage in Claude Code:"
echo "  /exchange-mail              # List unread emails"
echo "  /exchange-mail read <id>    # Read email"
echo "  /exchange-mail archive --external  # Archive external emails"
echo ""
