"""
Outlook Meeting Modification Functions

General-purpose functions for updating existing Outlook meetings with proper timezone handling.
All times must be in system LOCAL timezone when updating - do not use UTC.
"""

from datetime import datetime, timedelta
from outlook_manager import OutlookManager, get_outlook_calendar


def find_appointment(subject, start_date=None, calendar=None):
    """
    Find an appointment by subject and optional date.
    
    Args:
        subject: Meeting subject to search for
        start_date: datetime or date to restrict search (optional)
        calendar: Calendar object (optional)
        
    Returns:
        Appointment object if found, None otherwise
    """
    if calendar is None:
        calendar = get_outlook_calendar()
    
    if not calendar:
        return None
    
    items = calendar.Items
    items.Sort("[Start]")
    
    # Simple search by subject
    result = items.Find(f"[Subject] = '{subject}'")
    if result:
        return result
    
    # If date provided, restrict search
    if start_date:
        start_str = start_date.strftime('%m/%d/%Y')
        restriction = f"[Start] >= '{start_str}'"
        filtered = items.Restrict(restriction)
        result = filtered.Find(f"[Subject] = '{subject}'")
        if result:
            return result
    
    return None


def update_meeting(
    subject,
    new_subject=None,
    new_start=None,
    new_end=None,
    new_location=None,
    new_body=None,
    send_update=True,
):
    """
    Update an existing meeting with new details.
    
    IMPORTANT: new_start and new_end must be in system LOCAL timezone (not UTC).
    
    Args:
        subject: Current meeting subject to find
        new_subject: New subject text (optional)
        new_start: New start time in LOCAL timezone (optional, not UTC)
        new_end: New end time in LOCAL timezone (optional, not UTC)
        new_location: New location (optional)
        new_body: New description (optional)
        send_update: Send update notification to attendees (default True)
        
    Returns:
        Updated appointment object if successful, None if failed
    """
    appointment = find_appointment(subject)
    if not appointment:
        print(f"Meeting '{subject}' not found")
        return None
    
    try:
        # Update subject
        if new_subject:
            appointment.Subject = new_subject
        
        # Update times (use LOCAL time directly)
        if new_start:
            appointment.Start = new_start
        
        if new_end:
            appointment.End = new_end
        
        # If only start time provided, keep duration
        if new_start and not new_end:
            original_duration = appointment.End - appointment.Start
            appointment.End = new_start + original_duration
        
        # Update location and body
        if new_location:
            appointment.Location = new_location
        
        if new_body:
            appointment.Body = new_body
        
        # Save changes
        appointment.Save()
        
        # Notify attendees if meeting has recipients
        if send_update and appointment.Recipients.Count > 0:
            appointment.SendUpdate()
        
        return appointment
        
    except Exception as e:
        print(f"Error updating meeting: {e}")
        return None


def reschedule_meeting(subject, new_start, new_end, send_update=True):
    """
    Reschedule a meeting to new date/time.
    
    IMPORTANT: new_start and new_end must be in system LOCAL timezone.
    
    Args:
        subject: Meeting subject to find
        new_start: New start time in LOCAL timezone (required)
        new_end: New end time in LOCAL timezone (required)
        send_update: Notify attendees (default True)
        
    Returns:
        Updated appointment if successful, None if failed
    """
    return update_meeting(
        subject=subject,
        new_start=new_start,
        new_end=new_end,
        send_update=send_update,
    )


def move_meeting_by_offset(subject, minutes_offset, send_update=True):
    """
    Move a meeting earlier or later by a specified offset.
    
    Args:
        subject: Meeting subject to find
        minutes_offset: Minutes to move (negative = earlier, positive = later)
        send_update: Notify attendees (default True)
        
    Returns:
        Updated appointment if successful, None if failed
    """
    appointment = find_appointment(subject)
    if not appointment:
        print(f"Meeting '{subject}' not found")
        return None
    
    try:
        offset = timedelta(minutes=minutes_offset)
        new_start = appointment.Start + offset
        new_end = appointment.End + offset
        
        appointment.Start = new_start
        appointment.End = new_end
        appointment.Save()
        
        if send_update and appointment.Recipients.Count > 0:
            appointment.SendUpdate()
        
        return appointment
        
    except Exception as e:
        print(f"Error moving meeting: {e}")
        return None


def add_attendee(subject, email, role=1):
    """
    Add an attendee to an existing meeting.
    
    Args:
        subject: Meeting subject to find
        email: Attendee email address
        role: 1=required, 2=optional (default 1)
        
    Returns:
        Updated appointment if successful, None if failed
    """
    appointment = find_appointment(subject)
    if not appointment:
        print(f"Meeting '{subject}' not found")
        return None
    
    try:
        recipient = appointment.Recipients.Add(email)
        recipient.Role = role
        appointment.Recipients.ResolveAll()
        appointment.Save()
        appointment.SendUpdate()
        return appointment
        
    except Exception as e:
        print(f"Error adding attendee: {e}")
        return None


def remove_attendee(subject, email):
    """
    Remove an attendee from an existing meeting.
    
    Args:
        subject: Meeting subject to find
        email: Attendee email to remove
        
    Returns:
        Updated appointment if successful, None if failed
    """
    appointment = find_appointment(subject)
    if not appointment:
        print(f"Meeting '{subject}' not found")
        return None
    
    try:
        # Iterate in reverse to avoid index issues when removing
        for i in range(appointment.Recipients.Count, 0, -1):
            recipient = appointment.Recipients(i)
            if recipient.Address == email:
                appointment.Recipients.Remove(i)
                break
        
        appointment.Save()
        appointment.SendUpdate()
        return appointment
        
    except Exception as e:
        print(f"Error removing attendee: {e}")
        return None


def update_attendee_role(subject, email, new_role):
    """
    Change an attendee's role in a meeting.
    
    Args:
        subject: Meeting subject to find
        email: Attendee email
        new_role: 1=required, 2=optional
        
    Returns:
        Updated appointment if successful, None if failed
    """
    appointment = find_appointment(subject)
    if not appointment:
        print(f"Meeting '{subject}' not found")
        return None
    
    try:
        for recipient in appointment.Recipients:
            if recipient.Address == email:
                recipient.Role = new_role
                break
        
        appointment.Save()
        appointment.SendUpdate()
        return appointment
        
    except Exception as e:
        print(f"Error updating attendee role: {e}")
        return None


def validate_meeting_update(new_start, new_end):
    """
    Validate meeting update parameters.
    
    Args:
        new_start: Start datetime
        new_end: End datetime
        
    Returns:
        tuple: (is_valid, error_message)
    """
    if new_start >= new_end:
        return False, "Start time must be before end time"
    
    if new_start < datetime.now():
        return False, "Cannot reschedule to past time"
    
    duration = new_end - new_start
    
    if duration.total_seconds() < 300:
        return False, "Meeting duration too short (minimum 5 minutes)"
    
    if duration.total_seconds() > 28800:
        return False, "Meeting duration too long (maximum 8 hours)"
    
    return True, ""
