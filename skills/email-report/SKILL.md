---
name: email-report
description: >
  Generate email inbox reports, summaries, and digests from any connected IMAP
  mailbox. Use when the user asks: "summarize my inbox", "what emails do I have
  today", "give me a report of my unread messages", "digest of my emails this
  week", "what's new across my mailboxes", "inbox overview", "email summary",
  or "show me what needs my attention".
---

# Email Report Skill

Generate clear, scannable email reports from one or all IMAP accounts.

## Workflow

1. Call `list_accounts` to confirm which accounts are available.
2. Call `get_recent_emails` with `account: "all"` and an appropriate `since_days`
   (default 1 for "today", 7 for "this week"). Use `limit: 25` to get enough coverage.
3. If the user asked about a specific topic or sender, also call `search_emails`
   with a relevant `query`.
4. Group results by account. Within each account, sort newest first.
5. Present the report in the format below.

## Report Format

Use this structure — keep it scannable, no walls of text:

```
## Email report — [date range]

### [Account name] ([email address])
**[N] unread** · [M] total in period

| # | From | Subject | Date | Status |
|---|------|---------|------|--------|
| 1 | … | … | … | 🔴 Unread / ✅ Read |

### [Next account]
…

---
**Summary:** [1–2 sentences on overall volume and anything urgent]
**Needs attention:** [bullet list of emails that look time-sensitive, contain action words like "urgent", "deadline", "invoice", "approval", "RSVP", or are from known important senders]
```

## Guidelines

- Flag emails as needing attention if: subject contains urgent/deadline/invoice/
  approval/action required/RSVP/overdue, or if the email is older than 3 days
  and still unread.
- If no emails are found for an account, say so briefly — don't omit the account.
- If an account returns an error, report the error and continue with others.
- After presenting the report, ask: "Would you like me to open any of these,
  draft a reply, or set a reminder?"
