---
name: outlook-set-automatic-replies
description: 'Set Outlook automatic replies (Out of Office) for a selected mailbox with internal and external bilingual text templates.'
argument-hint: 'Specify mailbox, on/off state, and date range text to include (e.g., 22.07.2026-12.08.2026)'
user-invocable: true
---

# Setting Outlook Automatic Replies By Mailbox

## Overview

This skill explains how to set Automatic Replies (Out of Office) for a selected mailbox in Outlook.
Use it when you need to enable/disable OOF and set internal/external texts.

Supported mailbox targets:
- andreas.bossert@itk-engineering.de
- andreas.bossert2@de.bosch.com

## Tool To Use

Use MCP tool `set_automatic_replies` with:
- `enabled`: true or false
- `store_index`: mailbox store index
- `internal_text`: message for internal recipients
- `external_text`: message for external recipients

Then verify with MCP tool `get_automatic_replies`.

## Mailbox Selection

### Fast path (known profile order)

In this profile, mailbox stores are typically:
- `store_index: 1` -> andreas.bossert@itk-engineering.de
- `store_index: 2` -> andreas.bossert2@de.bosch.com

### Safe path (always verify index first)

```python
import pythoncom
import win32com.client

pythoncom.CoInitialize()
try:
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    for i in range(1, ns.Stores.Count + 1):
        print(i, ns.Stores.Item(i).DisplayName)
finally:
    pythoncom.CoUninitialize()
```

## Recommended Execution Pattern

1. Resolve target mailbox -> `store_index`.
2. Call `set_automatic_replies` with `enabled: true` and both texts.
3. Call `get_automatic_replies` for confirmation.
4. If text fields return `(not set)`, keep OOF enabled and provide the text for manual paste in Outlook UI.

## Text Template Example

Use the following exact texts.

### INTERN

Vielen Dank für
Deine Nachricht.

Ich bin vom
22.07.2026 bis einschließlich 12.08.2026 nicht im Büro und habe in dieser Zeit
keinen Zugriff auf meine E-Mails.

Thank you for your
message.

I am out of the
office from 22 July 2026 through 12 August 2026 (inclusive) and will have no
access to email during this period.

In dringenden Fällen
wende Dich bitte an Torsten Breitel bzw. das V&V Team.

For urgent matters,
please contact my backup colleague or team.

Freundliche Grüße /
Kind regards

Andreas Bossert

### EXTERN

Vielen Dank für
Ihre Nachricht.

Ich bin vom
22.07.2026 bis einschließlich 12.08.2026 nicht im Büro und habe in dieser Zeit
keinen Zugriff auf meine E-Mails.

Thank you for your
message.

I am out of the
office from 22 July 2026 through 12 August 2026 (inclusive) and will have no
access to email during this period.

In dringenden Fällen
wenden Sie sich bitte an torsten.breitel@itk-engineering.de bzw. das
V&V-Team (test-strategy@itk.engineering.de).

For urgent matters,
please contact my backup colleague or team.

Freundliche Grüße /
Kind regards

Andreas Bossert

## Example MCP Call

```json
{
  "enabled": true,
  "store_index": 1,
  "internal_text": "<INTERN text above>",
  "external_text": "<EXTERN text above>"
}
```

## Notes

- Date ranges are part of the message text in this server implementation.
- If your Outlook profile blocks text-property writes, the OOF state can still be toggled and texts may need to be pasted manually in Outlook.
