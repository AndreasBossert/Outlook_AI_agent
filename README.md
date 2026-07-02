# Outlook AI Assistant — MCP Server

Connect AI assistants (GitHub Copilot, Claude, etc.) to Microsoft Outlook via the **Windows COM interface** using the [Model Context Protocol (MCP)](https://modelcontextprotocol.io/).

## Features

| Tool | Description |
| ------ | ------------- |
| `list_folders` | List all mail folders (optionally recursive) |
| `get_emails` | Retrieve emails from any folder (newest first, with unread filter) |
| `get_email_details` | Full email details including body and attachment list |
| `search_emails` | Search by subject, body, sender, or all fields |
| `send_email` | Compose and send a new email |
| `reply_to_email` | Reply (or reply-all) to an existing email |
| `forward_email` | Forward an email to new recipients |
| `move_email` | Move an email to a different folder |
| `delete_email` | Move an email to Deleted Items |
| `mark_email` | Mark as read / unread |
| `get_calendar_events` | Retrieve calendar events in a date range |
| `create_calendar_event` | Create a new appointment |
| `get_contacts` | List / search contacts |
| `create_contact` | Add a new contact |
| `get_tasks` | List tasks (open or including completed) |
| `get_automatic_replies` | Show automatic replies status and internal/external text |
| `set_automatic_replies` | Enable/disable automatic replies and optionally set texts |

## Requirements

- Windows with **Microsoft Outlook** installed and a configured mailbox
- Python 3.10+

## Installation

```powershell
pip install -r requirements.txt
```

## Running the server

```powershell
python server.py
```

The server communicates over **stdio** (standard MCP transport).

## VS Code / Copilot integration

Add the following entry to your VS Code MCP configuration (`.vscode/mcp.json` or user settings):

```json
{
  "servers": {
    "outlook": {
      "type": "stdio",
      "command": "python",
      "args": ["${workspaceFolder}/server.py"]
    }
  }
}
```

## Usage examples

Once connected, you can ask your AI assistant things like:

- *"Show me my unread emails from the last 3 days."*
- *"Search my inbox for emails about project X."*
- *"Reply to the email with subject 'Budget Review' and tell them I'll attend."*
- *"Move all newsletter emails to the 'Newsletter' subfolder."*
- *"What meetings do I have this week?"*
- *"Create a reminder appointment for tomorrow at 9:00 AM."*
- *"Are my automatic replies enabled, and what text is configured?"*
- *"Enable automatic replies and set internal/external text."*
