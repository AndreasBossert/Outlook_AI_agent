---
name: outlook-read-meetings
description: 'Read Outlook meetings with proper timezone handling. Use when retrieving calendar events, reading meeting details, iterating through appointments. Handles Windows COM API UTC/local time conversion issues.'
argument-hint: 'Specify date range or meeting criteria (e.g., "today", "this week", specific date range)'
user-invocable: true
---

# Reading Outlook Meetings

## Overview

This skill provides comprehensive guidance for reading/retrieving Outlook meetings via the Windows COM API while correctly handling timezone conversions. The COM API has known quirks with datetime formats that require special handling.

## Critical Timezone Issue: UTC vs Local Time

### The Problem

The Outlook Windows COM API stores all datetime values in **UTC internally**, but the `.Start` and `.End` properties return values in the **system's local timezone** when accessed from Python or VBScript. This creates confusion because:

1. Datetimes are stored as UTC in Outlook
2. When retrieved via COM, they're automatically converted to local time
3. If you convert again, you get wrong times
4. Timezone info is often lost in the conversion

### Solution Approach

**DO NOT convert COM-retrieved datetimes again** — they're already in local time. Use this pattern:

```python
# Correct: COM returns local time already
appointment = calendar_item
start_time = appointment.Start  # Already in LOCAL timezone
end_time = appointment.End      # Already in LOCAL timezone

# Store with timezone awareness if needed
import datetime
local_tz = datetime.datetime.now().astimezone().tzinfo
aware_start = start_time.replace(tzinfo=local_tz)
```

**DO NOT do this:**
```python
# WRONG: Double conversion
import pytz
utc_dt = appointment.Start  # Already converted from UTC to local!
eastern = pytz.timezone('US/Eastern')
wrong_time = utc_dt.replace(tzinfo=pytz.UTC).astimezone(eastern)  # INCORRECT!
```

## Procedure: Reading Meetings

### 1. Connect to Outlook

Use the [outlook_manager.py](./scripts/outlook_manager.py) module to establish a connection:

```python
from scripts.outlook_manager import OutlookManager, get_outlook_calendar

# Option A: Quick connection
calendar = get_outlook_calendar()

# Option B: Managed connection
manager = OutlookManager()
if manager.connect():
    calendar = manager.get_calendar()
```

### 2. Access Calendar Items

Use functions from [outlook_read.py](./scripts/outlook_read.py) to retrieve appointments:

**Get all meetings:**
```python
from scripts.outlook_read import iterate_all_appointments

appointments = iterate_all_appointments()
for appt in appointments:
    print(appt.Subject, appt.Start, appt.End)
```

**Get meetings in date range:**
```python
from datetime import datetime, timedelta
from scripts.outlook_read import get_meetings_in_range

start = datetime.now()
end = start + timedelta(days=7)
appointments = get_meetings_in_range(start, end)
```

**Get today's meetings:**
```python
from scripts.outlook_read import get_today_meetings

appts = get_today_meetings()
for appt in appts:
    print(f"{appt.Subject}: {appt.Start} to {appt.End}")
```

**Get upcoming meetings (next N days):**
```python
from scripts.outlook_read import get_upcoming_meetings

appts = get_upcoming_meetings(days=7)  # Next 7 days
```

### 3. Extract Meeting Details

Use [get_meeting_details()](./scripts/outlook_read.py) to extract comprehensive meeting information:

```python
from scripts.outlook_read import get_meeting_details, get_upcoming_meetings

appts = get_upcoming_meetings(days=1)
for appt in appts:
    details = get_meeting_details(appt)
    print(f"Subject: {details['subject']}")
    print(f"Time: {details['start']} to {details['end']}")
    print(f"Organizer: {details['organizer']}")
    print(f"Required: {details['required_attendees']}")
    print(f"Optional: {details['optional_attendees']}")
    print(f"Online: {details['is_online']} - {details['online_url']}")
```

Returned dict keys:
- `subject`, `start`, `end`, `location`, `body`
- `organizer`: Organizer name or None
- `required_attendees`, `optional_attendees`: Lists of attendee names
- `is_recurring`: Boolean
- `is_online`: Boolean
- `online_url`: Teams/Skype URL if available
- `response_status`: 0=pending, 1=accepted, 2=declined, 3=tentative

### 4. Handle Recurring Meetings

Use [get_meeting_recurrence_info()](./scripts/outlook_read.py) for recurring meeting details:

```python
from scripts.outlook_read import get_meeting_details, get_meeting_recurrence_info

details = get_meeting_details(appointment)
if details.get('is_recurring'):
    recurrence_info = get_meeting_recurrence_info(appointment)
    print(f"Pattern: {recurrence_info['pattern']}")
    print(f"Interval: {recurrence_info['interval']}")
    print(f"End date: {recurrence_info['end_date']}")
```

Recurrence pattern values:
- 0 = Daily
- 1 = Weekly
- 2 = Monthly
- 3 = Yearly

### 5. Search for Specific Meetings

Use [find_meetings()](./scripts/outlook_read.py) to search by subject:

```python
from scripts.outlook_read import find_meetings, get_meeting_details

result = find_meetings("Team Standup")
if result:
    details = get_meeting_details(result)
    print(f"Found: {details['subject']}")
```

## Timezone Best Practices

Use functions from [timezone_utils.py](./scripts/timezone_utils.py) for timezone conversions:

### When Storing Meeting Data

```python
from datetime import datetime
from scripts.timezone_utils import com_to_aware_local, get_system_timezone

meeting = {
    'subject': appointment.Subject,
    'start_local': appointment.Start.isoformat(),  # Keep as-is (local)
    'system_timezone': get_system_timezone().tzname(datetime.now()),
}
```

### When Comparing Times

```python
from datetime import datetime, timedelta

# All times from COM are local — compare directly
appointment_start = appointment.Start
now = datetime.now()

if appointment_start < now:
    print("Meeting is in the past")
elif appointment_start < now + timedelta(days=1):
    print("Meeting is today")
```

### When Converting to a Different Timezone

Only if you need a timezone OTHER than system local:

```python
from scripts.timezone_utils import com_to_aware_local, convert_to_timezone

# Convert COM time (local) to UTC-aware
local_aware = com_to_aware_local(appointment.Start)

# Then convert to target timezone
eastern_time = convert_to_timezone(local_aware, 'US/Eastern')
print(f"Eastern time: {eastern_time}")
```

## Common Pitfalls

| Pitfall | Problem | Solution |
|---------|---------|----------|
| Double timezone conversion | Wrong time results | Don't convert COM times, they're already local |
| Ignoring timezone in stored data | Can't reproduce correct time later | Store timezone name alongside local time |
| Comparing with UTC times directly | Will be off by offset | Convert local appointment times to UTC first if needed |
| Lost timezone info on JSON serialization | Timezone unknown when deserializing | Store timezone string separately |

## Error Handling

```python
try:
    items = calendar.Items
    items.Sort("[Start]")
    for appointment in items:
        try:
            start = appointment.Start
            # Process meeting
        except AttributeError:
            print(f"Skipping item: Missing required property")
        except Exception as e:
            print(f"Error processing appointment: {e}")
except Exception as e:
    print(f"Failed to access calendar: {e}")
```

## References

- [Outlook Object Model - AppointmentItem](https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem)
- [Outlook RecurrencePattern](https://learn.microsoft.com/en-us/office/vba/api/outlook.recurrencepattern)
- [Windows COM API timezone handling](https://learn.microsoft.com/en-us/windows/win32/com/time-zone-handling)
