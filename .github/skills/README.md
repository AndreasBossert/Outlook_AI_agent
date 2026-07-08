# Outlook Meeting Skills - Script Refactoring

## Overview

Three independent skills for managing Outlook meetings have been refactored to extract reusable code into dedicated scripts. Each skill contains:

1. **SKILL.md** - Comprehensive documentation with references to scripts
2. **scripts/** - General-purpose, reusable Python modules

Additionally, one mailbox utility skill is available for reading Automatic Replies (Out of Office) text and status.

## Structure

### outlook-read-meetings
- **SKILL.md** - Guide for reading meetings with timezone safety
- **scripts/**
  - `outlook_manager.py` - Connection management (OutlookManager class)
  - `timezone_utils.py` - Timezone conversion utilities
  - `outlook_read.py` - Functions for retrieving meetings:
    - `get_meeting_details()` - Extract comprehensive meeting info
    - `find_meetings()` - Search by subject
    - `get_meetings_in_range()` - Date range queries
    - `get_today_meetings()`, `get_upcoming_meetings()`
    - `get_meeting_recurrence_info()` - Recurring meeting details
    - `iterate_all_appointments()` - Process all meetings

### outlook-create-meeting
- **SKILL.md** - Guide for creating meetings with timezone awareness
- **scripts/**
  - `outlook_manager.py` - Connection management
  - `timezone_utils.py` - Timezone conversion utilities
  - `outlook_create.py` - Functions for creating meetings:
    - `create_meeting()` - Full-featured meeting creation
    - `create_recurring_meeting()` - Recurring meetings
    - `create_online_meeting()` - Teams/Skype meetings
    - `validate_meeting_times()` - Time validation

### outlook-modify-meeting
- **SKILL.md** - Guide for updating meetings safely
- **scripts/**
  - `outlook_manager.py` - Connection management
  - `timezone_utils.py` - Timezone conversion utilities
  - `outlook_modify.py` - Functions for updating meetings:
    - `find_appointment()` - Locate meetings
    - `update_meeting()` - Update multiple properties
    - `reschedule_meeting()`, `move_meeting_by_offset()`
    - `add_attendee()`, `remove_attendee()`, `update_attendee_role()`
    - `validate_meeting_update()` - Update validation

### outlook-read-automatic-replies
- **SKILL.md** - Guide for reading OOF status/text for a selected mailbox
- Supports mailbox selection for:
  - `andreas.bossert@itk-engineering.de`
  - `andreas.bossert2@de.bosch.com`
- Uses `get_automatic_replies` with mailbox `store_index`

### outlook-set-automatic-replies
- **SKILL.md** - Guide for setting OOF status/text for a selected mailbox
- Supports mailbox selection for:
  - `andreas.bossert@itk-engineering.de`
  - `andreas.bossert2@de.bosch.com`
- Includes bilingual internal/external template text example
- Uses `set_automatic_replies` and verification via `get_automatic_replies`

## Key Design Principles

### General vs Concrete
- **General**: Functions accept parameters for flexibility (subject, date ranges, attendee lists, etc.)
- **Concrete**: Default values and sensible assumptions (e.g., system local timezone, current calendar)

### Timezone Safety
All three skills emphasize the critical COM API quirk:
- **COM stores**: UTC internally
- **COM returns on read**: Local time automatically (don't convert again!)
- **COM expects on write**: Local time input (pass UTC directly causes offset errors)

Functions include `IMPORTANT` notes about timezone handling.

### Error Handling
All functions include:
- Try/except blocks with user-friendly error messages
- Return `None` on failure (not exceptions)
- Validation functions before modifying

### Reusability
Common utilities are **duplicated across skills** (not shared) to ensure:
- Each skill is **self-contained** and independent
- No cross-skill dependencies
- Can be used standalone in different projects
- Files in `scripts/` work together within each skill

## Usage Examples

### Reading Meetings
```python
from scripts.outlook_read import get_meeting_details, get_today_meetings

appts = get_today_meetings()
for appt in appts:
    details = get_meeting_details(appt)
    print(f"{details['subject']}: {details['start']} to {details['end']}")
```

### Creating Meetings
```python
from datetime import datetime
from scripts.outlook_create import create_meeting

start = datetime(2025, 7, 15, 10, 0, 0)  # LOCAL TIME
end = datetime(2025, 7, 15, 10, 30, 0)   # LOCAL TIME

meeting = create_meeting(
    subject="Team Sync",
    start_time=start,
    end_time=end,
    location="Conference Room A",
    required_attendees=["alice@company.com"],
)
```

### Modifying Meetings
```python
from scripts.outlook_modify import reschedule_meeting, add_attendee

# Reschedule
reschedule_meeting(
    subject="Team Standup",
    new_start=datetime(2025, 7, 20, 15, 0),  # LOCAL
    new_end=datetime(2025, 7, 20, 15, 30),   # LOCAL
    send_update=True,
)

# Add attendee
add_attendee(
    subject="Team Standup",
    email="newperson@company.com",
    role=1,  # required
)
```

## Timezone Conversion Examples

### UTC to Local (from external API)
```python
from scripts.timezone_utils import utc_to_local
import pytz

external_utc = datetime(2025, 7, 15, 18, 0, tzinfo=pytz.UTC)
local_time = utc_to_local(external_utc)  # Safe for COM!
```

### COM Local to UTC (for storage)
```python
from scripts.timezone_utils import com_to_aware_local, aware_local_to_utc

# COM returns local time
appt_time_local = appointment.Start  # Already local

# Make timezone-aware
aware_local = com_to_aware_local(appt_time_local)

# Convert to UTC
utc_time = aware_local_to_utc(aware_local)
```

### Convert to Different Timezone
```python
from scripts.timezone_utils import com_to_aware_local, convert_to_timezone

appt_time = appointment.Start  # Local
aware = com_to_aware_local(appt_time)
eastern = convert_to_timezone(aware, 'US/Eastern')
```

## Benefits of Script Extraction

| Aspect | Before | After |
|--------|--------|-------|
| Documentation size | ~350 lines per skill | ~150 lines per skill |
| Reusability | Code examples only | Executable functions |
| Maintenance | Duplicated patterns | DRY modules |
| Testing | N/A | Can unit test scripts |
| IDE support | Manual | Full autocomplete |
| Error handling | Examples only | Built-in validation |
| Type hints | Documented | Can add to scripts |

## Integration with AI Assistant

These skills are designed for use with the Outlook AI Assistant:

1. Skills appear as `/` commands in chat
2. Agent auto-loads relevant skill based on task description
3. User can reference scripts by name in requests
4. All code is production-ready with error handling

## Future Enhancements

Possible improvements to scripts:

- Add `@dataclass` for meeting objects (instead of dicts)
- Add logging instead of print statements
- Add async support for batch operations
- Add type hints to all functions
- Create shared script repository (central utilities)
- Add retry logic for network issues
- Cache calendar for repeated reads
