"""
Outlook Meeting Reading Functions

General-purpose functions for reading, retrieving, and searching Outlook meetings.
All returned datetimes from COM are already in local time - do not convert again.
"""

from datetime import datetime, timedelta
from outlook_manager import OutlookManager, get_outlook_calendar


def get_meeting_details(appointment):
    """
    Extract comprehensive meeting information from an appointment.
    
    Args:
        appointment: Outlook appointment COM object
        
    Returns:
        dict: Meeting details with keys:
            - subject: Meeting title
            - start: Start time (LOCAL timezone from COM)
            - end: End time (LOCAL timezone from COM)
            - location: Meeting location
            - body: Meeting description
            - organizer: Organizer name or None
            - required_attendees: List of required attendee names
            - optional_attendees: List of optional attendee names
            - is_recurring: Boolean
            - is_online: Boolean
            - online_url: Teams/Skype URL if online meeting
            - response_status: 0=pending, 1=accepted, 2=declined, 3=tentative
    """
    try:
        required = [a.Name for a in appointment.Recipients if a.Role == 1]
        optional = [a.Name for a in appointment.Recipients if a.Role == 2]
        
        details = {
            'subject': appointment.Subject,
            'start': appointment.Start,  # LOCAL TIME - don't convert!
            'end': appointment.End,      # LOCAL TIME - don't convert!
            'location': appointment.Location,
            'body': appointment.Body,
            'organizer': appointment.Organizer.Name if appointment.Organizer else None,
            'required_attendees': required,
            'optional_attendees': optional,
            'is_recurring': appointment.IsRecurring,
            'is_online': appointment.IsOnlineMeeting,
            'online_url': appointment.OnlineMeetingUrl if appointment.IsOnlineMeeting else None,
            'response_status': appointment.ResponseStatus,
        }
        return details
    except Exception as e:
        print(f"Error extracting meeting details: {e}")
        return None


def find_meetings(search_term, calendar=None):
    """
    Search for meetings by subject (exact or partial match).
    
    Args:
        search_term: Text to search for in meeting subject
        calendar: Calendar object (optional, uses default if not provided)
        
    Returns:
        Appointment object if found (first match), None if not found
    """
    if calendar is None:
        calendar = get_outlook_calendar()
    
    if not calendar:
        return None
    
    items = calendar.Items
    items.Sort("[Start]")
    
    # Try exact match first
    result = items.Find(f"[Subject] = '{search_term}'")
    if result:
        return result
    
    # Fall back to partial match
    result = items.Restrict(f"[Subject] LIKE '%{search_term}%'")
    return result


def get_meetings_in_range(start_date, end_date, calendar=None):
    """
    Get all meetings within a date range.
    
    Args:
        start_date: datetime or date object (inclusive)
        end_date: datetime or date object (exclusive)
        calendar: Calendar object (optional)
        
    Returns:
        List of appointment objects in the range, sorted by start time
    """
    if calendar is None:
        calendar = get_outlook_calendar()
    
    if not calendar:
        return []
    
    # Format dates for Outlook restriction
    start_str = start_date.strftime('%m/%d/%Y')
    end_str = end_date.strftime('%m/%d/%Y')
    
    items = calendar.Items
    items.Sort("[Start]")
    
    restriction = f"[Start] >= '{start_str}' AND [Start] < '{end_str}'"
    filtered_items = items.Restrict(restriction)
    
    # Convert to list to preserve results
    return list(filtered_items)


def iterate_all_appointments(calendar=None, callback=None):
    """
    Iterate through all calendar appointments with optional callback.
    
    Args:
        calendar: Calendar object (optional)
        callback: Function to call for each appointment (optional)
                 Signature: callback(appointment) -> None
                 
    Returns:
        List of all appointments if no callback provided, None if callback used
    """
    if calendar is None:
        calendar = get_outlook_calendar()
    
    if not calendar:
        return []
    
    items = calendar.Items
    items.Sort("[Start]")
    
    if callback:
        for appointment in items:
            try:
                callback(appointment)
            except Exception as e:
                print(f"Error processing appointment: {e}")
        return None
    else:
        return list(items)


def get_today_meetings(calendar=None):
    """
    Get all meetings for today.
    
    Args:
        calendar: Calendar object (optional)
        
    Returns:
        List of today's appointments
    """
    today = datetime.now()
    tomorrow = today + timedelta(days=1)
    return get_meetings_in_range(today, tomorrow, calendar)


def get_upcoming_meetings(days=7, calendar=None):
    """
    Get meetings for the next N days.
    
    Args:
        days: Number of days to look ahead (default 7)
        calendar: Calendar object (optional)
        
    Returns:
        List of upcoming appointments
    """
    start = datetime.now()
    end = start + timedelta(days=days)
    return get_meetings_in_range(start, end, calendar)


def get_meeting_recurrence_info(appointment):
    """
    Extract recurrence pattern information from a recurring meeting.
    
    Args:
        appointment: Outlook appointment COM object
        
    Returns:
        dict with recurrence info if recurring, None if not recurring:
            - pattern: 0=daily, 1=weekly, 2=monthly, 3=yearly
            - interval: Recurrence interval
            - occurrences: Number of occurrences (0 if unlimited)
            - end_date: Pattern end date (LOCAL TIME)
    """
    if not appointment.IsRecurring:
        return None
    
    try:
        recurrence = appointment.RecurrencePattern
        return {
            'pattern': recurrence.RecurrenceType,
            'interval': recurrence.Interval,
            'occurrences': recurrence.Occurrences,
            'end_date': recurrence.PatternEndDate,  # LOCAL TIME
        }
    except Exception as e:
        print(f"Error getting recurrence info: {e}")
        return None
