# Claude Exchange Mail

A Claude Code skill for managing Microsoft Exchange/Outlook emails directly from your terminal.

## Features

- 📧 **List** unread emails with smart filtering (To/CC only)
- 📖 **Read** full email content
- ✉️ **Reply** to emails with quoted original
- ✅ **Mark as read** - single or batch
- 📁 **Archive** - single or batch (external/internal/all)

## Installation

### Prerequisites

- Python 3.10+
- Claude Code CLI
- Microsoft Exchange account with EWS access

### Install Dependencies

```bash
pip install exchangelib
```

### Install the Skill

```bash
# Clone the repository
git clone https://github.com/DmitryBMsk/claude-exchange-mail.git

# Copy skill to Claude Code skills directory
cp -r claude-exchange-mail/skills/exchange-mail ~/.claude/skills/

# Copy the script
cp claude-exchange-mail/src/exchange_mail.py ~/.claude/
```

### Configure Credentials

Add to your `~/.zshrc` or `~/.bashrc`:

```bash
export EXCHANGE_SERVER="mail.yourcompany.com"
export EXCHANGE_EMAIL="your.email@company.com"
export EXCHANGE_USERNAME="your_username"
export EXCHANGE_PASSWORD="your_password"
```

Then reload:
```bash
source ~/.zshrc
```

## Usage

### List Unread Emails

```bash
/exchangeMail                    # Today's unread where you're in To/CC
/exchangeMail list --days 3      # Last 3 days
/exchangeMail list --all         # All unread (not just To/CC)
/exchangeMail list --json        # JSON output
```

### Read Email

```bash
/exchangeMail read <email_id>
```

### Reply to Email

```bash
/exchangeMail reply <email_id> "Your reply message"
```

### Mark as Read

```bash
/exchangeMail mark-read <email_id>      # Single email
/exchangeMail mark-read --external      # All external emails
/exchangeMail mark-read --internal      # All internal (company) emails
/exchangeMail mark-read --all           # All unread emails
```

### Archive Emails

```bash
/exchangeMail archive <email_id>        # Single email
/exchangeMail archive --external        # All external emails
/exchangeMail archive --internal        # All internal emails
/exchangeMail archive --all             # All unread emails
/exchangeMail archive --external --days 7  # External from last 7 days
```

## Output Example

```
📧 9 unread emails today:

━━━ Internal (4) ━━━
[b7bc8d99] [13:57] John Smith
        Re: Project Discussion
        Hi team, here's what we agreed on...

━━━ External (5) ━━━
[43e56cc9] [09:50] newsletter@company.com
        Weekly Update

──────────────────────────────────────────────────
Commands: read <id>, reply <id> "text", mark-read <id>, archive <id>
```

## Email IDs

Each email gets a stable 8-character hex ID (e.g., `b7bc8d99`).
This ID is consistent across script runs and can be used for all commands.

## Workflow Examples

### Morning Routine

```bash
# 1. Check unread emails
/exchangeMail

# 2. Read important one
/exchangeMail read b7bc8d99

# 3. Reply if needed
/exchangeMail reply b7bc8d99 "Thanks, I'll look into this!"

# 4. Archive all external spam
/exchangeMail archive --external
```

### Weekly Cleanup

```bash
# Archive all external from last 7 days
/exchangeMail archive --external --days 7

# Mark all internal as read (if just FYI)
/exchangeMail mark-read --internal --days 7
```

## Configuration

### Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `EXCHANGE_SERVER` | Exchange server hostname | Yes |
| `EXCHANGE_EMAIL` | Your email address | Yes |
| `EXCHANGE_USERNAME` | Your username | Yes |
| `EXCHANGE_PASSWORD` | Your password | Yes |

### Internal Domain Detection

By default, emails from `@yourcompany.com` are classified as "internal".
To customize, modify the `is_internal` check in `exchange_mail.py`:

```python
is_internal = sender.lower().endswith('@yourcompany.com')
```

## Security Notes

⚠️ **Important**: Never commit your credentials to version control!

- Store credentials in environment variables
- Add `.env` files to `.gitignore`
- Consider using a secrets manager for production

## Troubleshooting

### SSL Certificate Errors

The script disables SSL verification for corporate certificates. If you need strict verification, remove these lines:

```python
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
urllib3.disable_warnings()
```

### Connection Issues

1. Verify your Exchange server supports EWS
2. Check if your account has EWS access enabled
3. Try connecting with Outlook to verify credentials

### Missing Emails

- By default, only emails where you're in To/CC are shown
- Use `--all` flag to see all unread emails
- Check `--days` parameter if looking for older emails

## License

MIT License - see [LICENSE](LICENSE)

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Author

[DmitryBMsk](https://github.com/DmitryBMsk)

## Related Projects

- [Claude Code](https://claude.ai/code) - AI-powered coding assistant
- [exchangelib](https://github.com/ecederstrand/exchangelib) - Python client for Microsoft Exchange
