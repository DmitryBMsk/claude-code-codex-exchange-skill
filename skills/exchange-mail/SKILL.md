---
name: exchange-mail
description: Get unread Exchange/Outlook emails from today where you are a direct recipient (To/CC)
---

# Exchange Mail CLI

Full email management for Microsoft Exchange/Outlook.

## Quick Commands

```bash
# List unread emails from today
source ~/.zshrc && python3 ~/.claude/exchange_mail.py list

# Read specific email
source ~/.zshrc && python3 ~/.claude/exchange_mail.py read <id>

# Reply to email
source ~/.zshrc && python3 ~/.claude/exchange_mail.py reply <id> "Your message"

# Mark as read
source ~/.zshrc && python3 ~/.claude/exchange_mail.py mark-read <id>
source ~/.zshrc && python3 ~/.claude/exchange_mail.py mark-read --external
source ~/.zshrc && python3 ~/.claude/exchange_mail.py mark-read --internal
source ~/.zshrc && python3 ~/.claude/exchange_mail.py mark-read --all

# Archive emails
source ~/.zshrc && python3 ~/.claude/exchange_mail.py archive <id>
source ~/.zshrc && python3 ~/.claude/exchange_mail.py archive --external
source ~/.zshrc && python3 ~/.claude/exchange_mail.py archive --internal
source ~/.zshrc && python3 ~/.claude/exchange_mail.py archive --all
```

## Commands

### list - List unread emails
```bash
list [--days N] [--all] [--json]
```
- `--days N` - Look back N days (default: 0 = today only)
- `--all` - Show all unread (not just where you're in To/CC)
- `--json` - Output as JSON

### read - Read full email
```bash
read <email_id>
```
Shows full email content with sender, recipients, date, and body.

### reply - Reply to email
```bash
reply <email_id> "Your reply message"
```
Sends reply with quoted original message.

### mark-read - Mark as read
```bash
mark-read <email_id>           # Single email
mark-read --external           # All external emails
mark-read --internal           # All internal emails
mark-read --all                # All unread emails
mark-read --external --days 3  # External from last 3 days
```

### archive - Archive emails
```bash
archive <email_id>             # Single email
archive --external             # All external emails
archive --internal             # All internal emails
archive --all                  # All unread emails
archive --external --days 7    # External from last 7 days
```

## Output Format

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
# 1. Check unread
/exchange-mail

# 2. Read important email
/exchange-mail read 33aa2016

# 3. Reply if needed
/exchange-mail reply 33aa2016 "Thanks, I'll look into this!"

# 4. Archive all spam
/exchange-mail archive --external
```

### Quick Cleanup
```bash
# Archive all external spam
/exchange-mail archive --external

# Mark internal as read (if just FYI)
/exchange-mail mark-read --internal
```

### Weekly Cleanup
```bash
# Archive all external from last 7 days
/exchange-mail archive --external --days 7
```

## Environment Variables

Required in `~/.zshrc` or `~/.bashrc`:
```bash
export EXCHANGE_SERVER="mail.company.com"
export EXCHANGE_EMAIL="user@company.com"
export EXCHANGE_USERNAME="username"
export EXCHANGE_PASSWORD="password"
```

## Related Commands

- `/slack-unread-messages` - Get unread Slack messages
- `/slack-reply` - Reply to Slack message
- `/slack-eyes` - Mark Slack message as seen
