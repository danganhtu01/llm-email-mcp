---
name: email-reminders
description: >
  Scan emails for follow-ups, deadlines, and action items, then surface them as
  reminders. Use when the user asks: "what emails need follow-up", "do I have
  any pending email replies", "remind me to follow up with [person]", "what's
  overdue in my inbox", "emails I haven't responded to", "find action items in
  my emails", "what deadlines are coming up from emails", or "set a reminder
  for this email".
---

# Email Reminders Skill

Surface action items, deadlines, and pending follow-ups buried in email.

## Workflow

### Mode A — Scan for reminders across all accounts

1. Call `get_recent_emails` with `account: "all"`, `since_days: 14`, `limit: 30`.
2. For each email that looks action-oriented (see detection rules below), call
   `read_email` to get the full body — do this in batches of up to 5 at a time
   to avoid slowness.
3. Extract action items and deadlines from the body.
4. Present the Reminders Report (format below).

### Mode B — Reminder for a specific email

1. If the user named a sender or subject, call `search_emails` to find the UID.
2. Call `read_email` to get the full content.
3. Extract the action item or deadline.
4. Confirm the reminder with the user: "Got it — I'll remind you to [action]
   regarding '[subject]'. When would you like the reminder?"

## Detection Rules

Flag an email as action-oriented if any of the following appear in subject or body:

- **Deadline keywords:** due, deadline, by [date], expires, RSVP by, respond by,
  action required, must complete, submit by
- **Follow-up keywords:** following up, checking in, as discussed, per our call,
  waiting on your, pending your, awaiting approval
- **Request keywords:** please review, please confirm, please send, can you,
  could you, would you, let me know, your thoughts
- **Unanswered emails:** sent more than 3 days ago and still unread, or sent
  by the user (check sent folder) with no reply thread

## Reminders Report Format

```
## Pending follow-ups & action items

### Overdue (no response > 3 days)
- 📬 **[Subject]** from [Sender] · [Account] · [Date]
  Action: [what needs to happen]

### Deadlines coming up
- ⏰ **[Subject]** · Due: [extracted date or "soon"]
  Action: [what needs to happen]

### Awaiting your reply
- 💬 **[Subject]** from [Sender] · [Date received]
  They asked: [1-line summary of the ask]

### General action items
- 📋 **[Subject]** · [Account]
  Action: [extracted task]
```

## Guidelines

- Extract the specific ask or deadline from the email body — don't just list
  the subject. One line is enough.
- If no actionable emails are found, say so clearly and suggest the user check
  back after more emails arrive.
- Do not invent deadlines. Only surface dates or urgency language explicitly
  present in the email.
- After presenting the report, offer: "Want me to draft a reply for any of
  these, or download any attachments?"
