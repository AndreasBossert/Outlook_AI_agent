---
name: outlook-read-automatic-replies
description: 'Read Outlook automatic replies (Out of Office) for a selected mailbox. Use when checking OOF status and internal/external reply text for specific accounts such as andreas.bossert@itk-engineering.de and andreas.bossert2@de.bosch.com.'
argument-hint: 'Specify mailbox email: andreas.bossert@itk-engineering.de or andreas.bossert2@de.bosch.com'
user-invocable: true
---

# Reading Outlook Automatic Replies By Mailbox

## Overview

This skill explains how to read Automatic Replies (Out of Office) from a specific mailbox store in Outlook.
Use it when the user wants OOF status and configured internal/external texts for one of multiple accounts.

Supported mailbox targets:
- andreas.bossert@itk-engineering.de
- andreas.bossert2@de.bosch.com

## Tool To Use

Use MCP tool `get_automatic_replies` with `store_index`.

Example:

```json
{
  "store_index": 1
}
```

## Mailbox Selection

### Fast path (known profile order)

In this workspace/profile, mailbox stores are typically:
- `store_index: 1` -> andreas.bossert@itk-engineering.de
- `store_index: 2` -> andreas.bossert2@de.bosch.com

If this order is stable, use the matching index directly.

### Safe path (always verify index first)

Store order can change between Outlook profiles. Verify before reading:

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

Then call `get_automatic_replies` with the discovered index.

## Response Template

Return data in this structure:
- Store display name
- Store index
- Automatic replies enabled/disabled
- Internal reply text
- External reply text

## Example Requests

- "Read automatic replies for andreas.bossert@itk-engineering.de"
- "Check OOF text for andreas.bossert2@de.bosch.com"
- "Show OOF status for both mailboxes"

## Notes

- This skill is for reading only.
- If text shows `(not set)`, automatic replies may still be enabled but no retrievable message body is stored for that profile.
