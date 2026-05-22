---
name: reply-suggestion
description: >
  Draft reply suggestions for emails in any connected IMAP mailbox. Use when
  the user asks: "draft a reply to this email", "help me respond to [sender]",
  "write a reply", "suggest a response", "how should I respond", "reply to the
  email from [name]", "compose a response to", or "what should I say back".
---

# Reply Suggestion Skill

Read an email and draft a ready-to-send reply for the user to review and edit.

## Workflow

1. **Identify the email.** If the user named a sender or subject, call
   `search_emails` with `account: "all"` and the relevant `query`. If they
   said "the last email from X", use `since_days: 30` and `limit: 5`.
2. **Read the full email.** Once you have the UID and folder, call `read_email`
   to get the complete body and thread context.
3. **Understand the ask.** Determine:
   - What tone the user wants (if not specified, match the original email's tone)
   - Any specific points they want to make
   - Whether this is a first reply or follow-up
4. **Draft the reply.** Structure it as a real email (greeting, body, sign-off).
5. **Present clearly.** Show the draft in a code block or blockquote so it's
   easy to copy. Include the To/Subject pre-filled.

## Reply Draft Format

```
**To:** [original sender's address]
**Subject:** Re: [original subject]

---

[Greeting],

[Body of reply — 2–4 short paragraphs max. Mirror the original tone: formal
for formal, casual for casual. Be concise. Address each question or point
raised in the original email.]

[Sign-off],
[User's name or "Best regards" if unknown]
```

## Guidelines

- Always quote or paraphrase the key point you're responding to so context
  is clear.
- If the email asks multiple questions, address each one explicitly.
- If attachments were mentioned in the original, note "[Attachment: X]" as a
  reminder for the user to attach.
- After presenting the draft, offer: "Want me to adjust the tone, add
  anything, or set a follow-up reminder if you don't hear back?"
- Do not send anything — present the draft for the user to review only.
