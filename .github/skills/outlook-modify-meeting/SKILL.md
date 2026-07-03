---
name: outlook-modify-meeting
description: 'Modify Outlook meetings with proper timezone handling. Use when updating meeting times, attendees, location, or other details. Handles Windows COM API UTC/local time conversion issues on updates.'
argument-hint: 'Specify meeting identifier and changes (e.g., "reschedule to next Tuesday at 3 PM", "add attendees")'
user-invocable: true
---

# Modifying Outlook Meetings

## Overview

This skill provides comprehensive guidance for updating existing Outlook meetings via the Windows COM API while correctly handling the timezone conversion quirks specific to the COM interface.

## Critical Timezone Issue: UTC vs Local Time in Updates

### The Problem When Updating

Modifying meeting times has the same timezone risks as creating meetings:

- **On Read**: COM returns times in local timezone
- **On Write**: COM expects local timezone input
- **Risk**: Updating with UTC times will offset the meeting incorrectly
- **Complication**: You read local time, but must convert properly if the source is UTC

### Solution: Read Local, Update Local

```python
# Correct: Read returns local, update with local
appointment = find_appointment(...)
current_start = appointment.Start  # This is LOCAL
new_start = datetime(2025, 7, 20, 15, 0)  # NEW LOCAL time
appointment.Start = new_start  # Update with local
appointment.Save()

# Wrong: Converting already-local time to UTC
from datetime import datetime
import pytz
appointment.Start = current_start.replace(tzinfo=pytz.UTC)  # WRONG - double converts!
```

## Procedure: Modifying Meetings

### 1. Find the Appointment to Modify

Use [find_appointment()](./scripts/outlook_modify.py) to locate a meeting:

```python
from scripts.outlook_modify import find_appointment, update_meeting

# Simple search by subject
appointment = find_appointment("Team Standup")

# Or search within a date range
from datetime import datetime
start_date = datetime(2025, 7, 15)
appointment = find_appointment("Team Standup", start_date)

if appointment:
    print(f"Found: {appointment.Subject}")
    print(f"Time: {appointment.Start} to {appointment.End}")
else:
    print("Meeting not found")
```

### 2. Update Basic Meeting Properties

Use [update_meeting()](./scripts/outlook_modify.py) to modify appointment details:

```python
from scripts.outlook_modify import update_meeting

# Update subject, location, and body
meeting = update_meeting(
    subject="Team Standup",  # Meeting to find
    new_subject="Updated: Team Standup",
    new_location="Conference Room B",
    new_body="Updated meeting details and agenda",
    send_update=True,  # Notify attendees
)
```

### 3. Modify Meeting Time (With Timezone Care)

Use [reschedule_meeting()](./scripts/outlook_modify.py) or [move_meeting_by_offset()](./scripts/outlook_modify.py):

```python
from datetime import datetime
from scripts.outlook_modify import reschedule_meeting, move_meeting_by_offset

# Option A: Reschedule to specific times
new_start = datetime(2025, 7, 20, 15, 0, 0)  # 3 PM LOCAL
new_end = datetime(2025, 7, 20, 15, 30, 0)   # LOCAL

meeting = reschedule_meeting(
    subject="Team Standup",
    new_start=new_start,
    new_end=new_end,
    send_update=True,
)

# Option B: Move by offset (e.g., 30 minutes earlier)
meeting = move_meeting_by_offset(
    subject="Team Standup",
    minutes_offset=-30,  # Negative = earlier
    send_update=True,
)
```

### 4. Comprehensive Update Examples

All update functions documented in [outlook_modify.py](./scripts/outlook_modify.py):

```python
from datetime import datetime
from scripts.outlook_modify import (
    update_meeting,
    reschedule_meeting,
    add_attendee,
    remove_attendee,
    validate_meeting_update,
)

# Validate times before updating
new_start = datetime(2025, 7, 20, 15, 0)
new_end = datetime(2025, 7, 20, 15, 30)
is_valid, msg = validate_meeting_update(new_start, new_end)
if not is_valid:
    print(f"Invalid: {msg}")
    exit()

# Update multiple properties at once
meeting = update_meeting(
    subject="Team Standup",
    new_subject="Updated: Team Sync",
    new_start=new_start,  # LOCAL!
    new_end=new_end,      # LOCAL!
    new_location="Room 200",
    new_body="Updated agenda",
    send_update=True,
)
```

## Handling Different Timezone Scenarios

### Scenario 1: Reschedule to Same Time Next Week

```python
from datetime import timedelta
from scripts.outlook_modify import reschedule_meeting, find_appointment

appointment = find_appointment("Weekly Meeting")
current_start = appointment.Start  # LOCAL time
current_end = appointment.End

# Move to same time next week
new_start = current_start + timedelta(days=7)
new_end = current_end + timedelta(days=7)

meeting = reschedule_meeting(
    subject="Weekly Meeting",
    new_start=new_start,
    new_end=new_end,
)
```

### Scenario 2: Update from External UTC Timestamp

```python
from datetime import datetime
from scripts.timezone_utils import utc_to_local
from scripts.outlook_modify import reschedule_meeting

# External API provides UTC time
external_utc = datetime(2025, 7, 20, 19, 0, tzinfo=pytz.UTC)

# Convert to local BEFORE updating
local_time = utc_to_local(external_utc)
end_time = utc_to_local(external_utc.replace(hour=external_utc.hour + 1))

# Now safe to update
meeting = reschedule_meeting(
    subject="External Event",
    new_start=local_time,  # Correct!
    new_end=end_time,
)
```

### Scenario 3: Move Meeting Earlier by X Minutes

```python
from scripts.outlook_modify import move_meeting_by_offset

# Move 30 minutes earlier
meeting = move_meeting_by_offset(
    subject="Team Standup",
    minutes_offset=-30,
    send_update=True,
)
```

## Modifying Attendees

Use attendee functions from [outlook_modify.py](./scripts/outlook_modify.py):

### Add Attendees

```python
from scripts.outlook_modify import add_attendee

# Add required attendee
add_attendee(
    subject="Team Standup",
    email="newperson@company.com",
    role=1,  # 1=required, 2=optional
)

# Add optional attendee
add_attendee(
    subject="Team Standup",
    email="maybeattend@company.com",
    role=2,
)
```

### Remove Attendees

```python
from scripts.outlook_modify import remove_attendee

remove_attendee(
    subject="Team Standup",
    email="oldperson@company.com",
)
```

### Update Attendee Role

```python
from scripts.outlook_modify import update_attendee_role

update_attendee_role(
    subject="Team Standup",
    email="person@company.com",
    new_role=1,  # Change to required
)
```

## Modifying Recurring Meetings

### Modify All Occurrences

```python
appointment = find_appointment("Weekly Standup")

# Edit the recurring series
appointment.Subject = "Updated: Weekly Standup"
appointment.Location = "New Location"

# This updates all occurrences
appointment.Save()
appointment.SendUpdate()
```

### Modify Single Occurrence (Exception)

```python
from datetime import datetime

# Find the series
appointment = find_appointment("Weekly Standup")

# Get occurrence at specific date
occurrence_date = datetime(2025, 7, 22)  # LOCAL
occurrence = appointment.GetOccurrence(occurrence_date)

# Modify just this occurrence
occurrence.Subject = "CANCELED: Weekly Standup"
occurrence.Start = occurrence.Start + timedelta(hours=1)  # LOCAL + timedelta
occurrence.Save()
```

## Best Practices

### ✓ DO: Update Using Local Times

```python
from datetime import datetime

new_start = datetime(2025, 7, 20, 15, 0)  # LOCAL
appointment.Start = new_start
appointment.Save()
```

### ✗ DON'T: Convert Already-Local Times

```python
import pytz

# Wrong - double conversion
utc_tz = pytz.UTC
local_time = appointment.Start  # Already local from COM
wrong = local_time.replace(tzinfo=utc_tz)  # DO NOT DO THIS
```

### ✓ DO: Store Timezone Context

```python
from datetime import datetime

meeting_update = {
    'subject': appointment.Subject,
    'start_local': appointment.Start.isoformat(),
    'system_timezone': datetime.now().astimezone().tzinfo.tzname(datetime.now()),
    'updated_at': datetime.utcnow().isoformat(),
}
```

### ✓ DO: Notify Attendees of Changes

```python
appointment.Subject = "Updated Meeting Time"
appointment.Start = new_start
appointment.End = new_end

# Save the changes
appointment.Save()

# Send update notification if meeting has attendees
if appointment.Recipients.Count > 0:
    appointment.SendUpdate()
```

### ✗ DON'T: Forget to Call Save()

```python
# Changes won't persist without Save()
appointment.Start = new_start
# WRONG - forgot appointment.Save()

# Correct
appointment.Start = new_start
appointment.Save()  # Required!
```

## Validation Before Update

```python
from datetime import datetime

def validate_meeting_update(appointment, new_start, new_end):
    """Validate update before saving"""
    errors = []
    
    if new_start >= new_end:
        errors.append("Start time must be before end time")
    
    if new_start < datetime.now():
        errors.append("Cannot reschedule to past time")
    
    if (new_end - new_start).total_seconds() > 8 * 3600:  # 8 hours max
        errors.append("Meeting duration too long (max 8 hours)")
    
    if (new_end - new_start).total_seconds() < 300:  # 5 minutes min
        errors.append("Meeting duration too short (min 5 minutes)")
    
    return errors

# Usage
errors = validate_meeting_update(appointment, new_start, new_end)
if errors:
    print("Validation errors:")
    for error in errors:
        print(f"  - {error}")
else:
    appointment.Start = new_start
    appointment.End = new_end
    appointment.Save()
```

## Error Handling

```python
try:
    appointment = find_appointment("Team Standup")
    
    if not appointment:
        print("Meeting not found")
        return
    
    # Update with validation
    new_start = datetime(2025, 7, 20, 15, 0)
    new_end = datetime(2025, 7, 20, 15, 30)
    
    if new_start >= new_end:
        raise ValueError("Invalid time range")
    
    appointment.Start = new_start
    appointment.End = new_end
    appointment.Save()
    
    if appointment.Recipients.Count > 0:
        appointment.SendUpdate()
    
    print("Meeting updated successfully")
    
except Exception as e:
    print(f"Failed to update meeting: {e}")
```

## References

- [Outlook AppointmentItem](https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem)
- [Outlook Recipients Collection](https://learn.microsoft.com/en-us/office/vba/api/outlook.recipients)
- [Outlook RecurrencePattern](https://learn.microsoft.com/en-us/office/vba/api/outlook.recurrencepattern)
- [Python datetime arithmetic](https://docs.python.org/3/library/datetime.html)
