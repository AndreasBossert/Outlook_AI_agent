---
name: outlook-create-meeting
description: 'Create Outlook meetings with correct timezone handling. Use when scheduling new appointments, adding events to calendar. Handles Windows COM API date/time UTC conversion requirements.'
argument-hint: 'Provide meeting details (subject, date/time, attendees, location)'
user-invocable: true
---

# Creating Outlook Meetings

## Overview

This skill provides guidance for creating new Outlook meetings via the Windows COM API while correctly handling the timezone/UTC quirks specific to the COM interface.

## Critical Timezone Issue: UTC vs Local Time

### The Problem When Creating

The Windows COM API has asymmetric timezone behavior:

- **On Read**: COM automatically converts UTC to local time (what Outlook stores internally)
- **On Write**: COM expects local time input BUT interprets it based on system timezone
- **Ambiguity**: If you pass a datetime without timezone info, COM assumes local time
- **Risk**: Passing UTC directly without conversion results in times offset by your timezone

### Solution: Always Use Local Time for Creation

When creating meetings, **ALWAYS provide times in your system's local timezone**:

```python
from datetime import datetime

# Correct: Use local time
local_now = datetime.now()  # This is local time
appointment.Start = local_now  # COM expects local, stores as UTC internally

# Wrong: Using UTC directly
import pytz
utc_now = datetime.now(pytz.UTC)  # This is UTC
appointment.Start = utc_now  # COM will offset it incorrectly!
```

## Procedure: Creating Meetings

### 1-2. Connect and Create Simple Meeting

Use [create_meeting()](./scripts/outlook_create.py) function to handle connection and creation:

```python
from datetime import datetime
from scripts.outlook_create import create_meeting

# Simple meeting - minimal required parameters
start = datetime(2025, 7, 15, 10, 0, 0)  # LOCAL TIME
end = datetime(2025, 7, 15, 10, 30, 0)   # LOCAL TIME

meeting = create_meeting(
    subject="Team Standup",
    start_time=start,  # Must be LOCAL timezone!
    end_time=end,      # Must be LOCAL timezone!
)
```

### 3. Full Meeting Creation with All Details

```python
from datetime import datetime
from scripts.outlook_create import create_meeting, validate_meeting_times

# Validate before creating
start = datetime(2025, 7, 15, 10, 0, 0)  # LOCAL
end = datetime(2025, 7, 15, 10, 30, 0)   # LOCAL
is_valid, msg = validate_meeting_times(start, end)
if not is_valid:
    print(f"Invalid times: {msg}")
    exit()

# Create with all options
meeting = create_meeting(
    subject="Team Sync",
    start_time=start,           # Must be LOCAL!
    end_time=end,               # Must be LOCAL!
    location="Conference Room A",
    body="Weekly team synchronization meeting",
    required_attendees=["alice@company.com", "bob@company.com"],
    optional_attendees=["charlie@company.com"],
)

if meeting:
    print(f"Meeting created: {meeting.Subject}")
else:
    print("Failed to create meeting")
```

All parameters documented in [outlook_create.py](./scripts/outlook_create.py).

## Timezone Handling: Do's and Don'ts

### ✓ DO: Create with Local Time

```python
from datetime import datetime

# Correct
local_start = datetime(2025, 7, 15, 14, 0)  # 2:00 PM local
appointment.Start = local_start              # COM stores as UTC, retrieves as local
```

### ✗ DON'T: Create with UTC and Convert

```python
import pytz
from datetime import datetime

# Wrong - don't do this
utc_start = datetime(2025, 7, 15, 18, 0, tzinfo=pytz.UTC)  # 6:00 PM UTC
appointment.Start = utc_start  # COM will misinterpret timezone offset!
```

### ✓ DO: If You Must Convert from UTC

If incoming data is in UTC, convert to local BEFORE passing to COM:

```python
import pytz
from datetime import datetime

# Incoming UTC time from API or database
utc_time = datetime(2025, 7, 15, 18, 0, tzinfo=pytz.UTC)

# Get system timezone
local_tz = datetime.now().astimezone().tzinfo

# Convert UTC to local
local_time = utc_time.astimezone(local_tz).replace(tzinfo=None)

# NOW pass to COM
appointment.Start = local_time  # Correct!
```

## Creating Recurring Meetings

Use [create_recurring_meeting()](./scripts/outlook_create.py):

```python
from datetime import datetime
from scripts.outlook_create import create_recurring_meeting

start = datetime(2025, 7, 21, 10, 0, 0)  # LOCAL
end = datetime(2025, 7, 21, 10, 30, 0)   # LOCAL
end_date = datetime(2025, 12, 31)        # LOCAL

meeting = create_recurring_meeting(
    subject="Weekly Team Standup",
    start_time=start,
    end_time=end,
    recurrence_type=1,  # 0=daily, 1=weekly, 2=monthly, 3=yearly
    interval=1,         # Every 1 week
    end_date=end_date,  # Or use occurrences=26
    body="Recurring weekly meeting",
    attendees=["team@company.com"],
)
```

## Creating Online Meetings (Teams/Skype)

Use [create_online_meeting()](./scripts/outlook_create.py):

```python
from datetime import datetime
from scripts.outlook_create import create_online_meeting

start = datetime(2025, 7, 15, 15, 0, 0)  # LOCAL
end = datetime(2025, 7, 15, 16, 0, 0)    # LOCAL

meeting = create_online_meeting(
    subject="Virtual Meeting",
    start_time=start,
    end_time=end,
    body="Teams meeting with auto-generated URL",
    attendees=["team@company.com"],  # Required for URL generation
)

if meeting:
    print(f"Join at: {meeting.OnlineMeetingUrl}")
```

## Proper DateTime Handling Examples

### Scenario 1: Schedule for next Tuesday at 2 PM local time

```python
from datetime import datetime, timedelta
from scripts.outlook_create import create_meeting

# Calculate next Tuesday
today = datetime.now()
days_until_tuesday = (1 - today.weekday()) % 7
if days_until_tuesday == 0:
    days_until_tuesday = 7

next_tuesday = today + timedelta(days=days_until_tuesday)
meeting_time = next_tuesday.replace(hour=14, minute=0, second=0, microsecond=0)

create_meeting(
    subject="Tuesday Meeting",
    start_time=meeting_time,         # LOCAL
    end_time=meeting_time + timedelta(hours=1),
)
```

### Scenario 2: Schedule from database timestamp (ISO UTC format)

```python
from datetime import datetime
from scripts.timezone_utils import utc_to_local
from scripts.outlook_create import create_meeting

# From database: "2025-07-15T18:00:00Z" (UTC)
db_timestamp = "2025-07-15T18:00:00Z"
utc_time = datetime.fromisoformat(db_timestamp.replace('Z', '+00:00'))

# Convert to local BEFORE passing to create_meeting
local_time = utc_to_local(utc_time)
end_time = utc_to_local(utc_time.replace(hour=utc_time.hour + 1))

create_meeting(
    subject="Database Event",
    start_time=local_time,  # Now safe for COM
    end_time=end_time,
)
```

### Scenario 3: Schedule from user input (assume local)

```python
from datetime import datetime
from scripts.outlook_create import create_meeting

# User enters: "2025-07-15 14:00"
user_input = "2025-07-15 14:00"
local_time = datetime.strptime(user_input, "%Y-%m-%d %H:%M")

create_meeting(
    subject="User Scheduled Meeting",
    start_time=local_time,  # Treat as local
    end_time=local_time + timedelta(hours=1),
)
```

## Error Handling

```python
try:
    appointment = outlook.CreateItem(1)
    appointment.Subject = "New Meeting"
    appointment.Start = datetime(2025, 7, 15, 10, 0)
    appointment.End = datetime(2025, 7, 15, 11, 0)
    
    # Validate times
    if appointment.Start >= appointment.End:
        raise ValueError("Start time must be before end time")
    
    appointment.Save()
    print(f"Meeting created: {appointment.Subject}")
    
except Exception as e:
    print(f"Failed to create meeting: {e}")
    # Don't save if error occurred
```

## Validation Checklist

Before saving a created meeting:

- [ ] Start time is BEFORE end time
- [ ] Times are in LOCAL timezone (not UTC)
- [ ] Subject is not empty
- [ ] All attendee emails are resolved
- [ ] Duration is reasonable (not negative or 0)
- [ ] Date is in the future (unless intentional)

## References

- [Outlook AppointmentItem Object](https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem)
- [Outlook Recipients Collection](https://learn.microsoft.com/en-us/office/vba/api/outlook.recipients)
- [Python datetime documentation](https://docs.python.org/3/library/datetime.html)
