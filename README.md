# Claude Code & Codex MS Exchange Skill

A skill for **Claude Code** and **OpenAI Codex CLI** to manage Microsoft Exchange/Outlook emails directly from your terminal.

> 📧 Check emails, reply, archive — all without leaving your coding environment!

## Compatibility

| Tool | Status |
|------|--------|
| [Claude Code](https://claude.ai/code) | ✅ Full support |
| [OpenAI Codex CLI](https://github.com/openai/codex) | ✅ Full support |

## Features

- 📧 **List** unread emails with smart filtering (To/CC recipients only)
- 📖 **Read** full email content with metadata
- ✉️ **Reply** to emails with quoted original message
- ✅ **Mark as read** - single email or batch operations
- 📁 **Archive** - single email or batch (external/internal/all)
- 🏷️ **Stable IDs** - 8-character hex IDs persist across sessions

## Quick Start

### 1. Install

```bash
git clone https://github.com/DmitryBMsk/claude-code-codex-exchange-skill.git
cd claude-code-codex-exchange-skill
./install.sh
```

### 2. Configure

Add to `~/.zshrc` or `~/.bashrc`:

```bash
export EXCHANGE_SERVER="mail.yourcompany.com"
export EXCHANGE_EMAIL="your.email@company.com"
export EXCHANGE_USERNAME="your_username"
export EXCHANGE_PASSWORD="your_password"
```

### 3. Use

```bash
# In Claude Code or Codex CLI
/exchange-mail                    # List today's unread
/exchange-mail read abc123        # Read specific email
/exchange-mail archive --external # Archive all external emails
```

## Commands

### List Unread Emails

```bash
/exchange-mail                       # Today's unread where you're in To/CC
/exchange-mail list --days 3         # Last 3 days
/exchange-mail list --all            # All unread (not just To/CC)
/exchange-mail list --json           # JSON output for scripting
```

### Read Email

```bash
/exchange-mail read <email_id>
```

### Reply to Email

```bash
/exchange-mail reply <email_id> "Your reply message here"
```

### Mark as Read

```bash
/exchange-mail mark-read <email_id>      # Single email
/exchange-mail mark-read --external      # All external emails
/exchange-mail mark-read --internal      # All internal (company) emails
/exchange-mail mark-read --all           # All unread emails
/exchange-mail mark-read --external --days 7  # External from last week
```

### Archive Emails

```bash
/exchange-mail archive <email_id>        # Single email
/exchange-mail archive --external        # All external emails
/exchange-mail archive --internal        # All internal emails
/exchange-mail archive --all             # All unread emails
```

## Output Example

```
📧 9 unread emails today:

━━━ Internal (4) ━━━
[b7bc8d99] [13:57] John Smith
        Re: Project Discussion
        Hi team, here's what we agreed on...

[50a05833] [12:57] Jane Doe
        Re: Security Policies
        Thanks for the meeting. Here's the link...

━━━ External (5) ━━━
[43e56cc9] [09:50] newsletter@company.com
        Weekly Update

──────────────────────────────────────────────────
Commands: read <id>, reply <id> "text", mark-read <id>, archive <id>
```

## Workflow Examples

### Morning Routine

```bash
# 1. Check what's new
/exchange-mail

# 2. Read the important one
/exchange-mail read b7bc8d99

# 3. Quick reply
/exchange-mail reply b7bc8d99 "Thanks, I'll review this today!"

# 4. Clean up spam
/exchange-mail archive --external
```

### Weekly Inbox Cleanup

```bash
# Archive all external from last 7 days
/exchange-mail archive --external --days 7

# Mark all internal FYI emails as read
/exchange-mail mark-read --internal --days 7
```

### Scripting with JSON

```bash
# Get emails as JSON for further processing
/exchange-mail list --json | jq '.[] | select(.is_internal) | .subject'
```

## Installation Details

### Prerequisites

- Python 3.10 or higher
- Microsoft Exchange account with EWS (Exchange Web Services) access
- Claude Code CLI or OpenAI Codex CLI

### Manual Installation

```bash
# Install Python dependency
pip install exchangelib

# Create directories
mkdir -p ~/.claude/skills/exchange-mail

# Copy files
cp src/exchange_mail.py ~/.claude/
cp skills/exchange-mail/SKILL.md ~/.claude/skills/exchange-mail/

# Make executable
chmod +x ~/.claude/exchange_mail.py
```

### Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `EXCHANGE_SERVER` | Exchange server hostname (e.g., `mail.company.com`) | ✅ |
| `EXCHANGE_EMAIL` | Your email address | ✅ |
| `EXCHANGE_USERNAME` | Your username | ✅ |
| `EXCHANGE_PASSWORD` | Your password | ✅ |
| `EXCHANGE_DOMAIN` | Windows domain (optional) | ❌ |

## Security Notes

⚠️ **Important**: Never commit credentials to version control!

- Store credentials in environment variables only
- Add `.env` files to `.gitignore`
- Consider using a secrets manager for team environments
- The script disables SSL verification for corporate certificates

## Troubleshooting

### "exchangelib not installed"

```bash
pip install exchangelib
# or
pip3 install exchangelib
```

### SSL Certificate Errors

The script automatically disables SSL verification for corporate certificates. If you need strict verification, modify `exchange_mail.py`:

```python
# Remove or comment out these lines:
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
urllib3.disable_warnings()
```

### Connection Refused

1. Verify your Exchange server supports EWS
2. Check if your account has EWS access enabled
3. Try connecting via Outlook to verify credentials
4. Check firewall/VPN settings

### Emails Not Showing

- By default, only emails where you're in To/CC are displayed
- Use `--all` flag to see all unread emails
- Check `--days` parameter for older emails

## License

MIT License - see [LICENSE](LICENSE)

## Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Author

**DmitryBMsk** - [GitHub](https://github.com/DmitryBMsk)

## Acknowledgments

- [Claude Code](https://claude.ai/code) by Anthropic
- [OpenAI Codex](https://github.com/openai/codex)
- [exchangelib](https://github.com/ecederstrand/exchangelib) - Python client for Microsoft Exchange
