"""
Outlook Meeting Creation Functions

General-purpose functions for creating Outlook meetings with proper timezone handling.
Always pass LOCAL time to COM API - do not pass UTC directly.
"""

from datetime import datetime, timedelta
from outlook_manager import OutlookManager


def create_meeting(
    subject,
    start_time,
    end_time,
    location=None,
    body=None,
    required_attendees=None,
    optional_attendees=None,
    is_online=False,
    online_url=None,
):
    """
    Create a new Outlook meeting with comprehensive options.
    
    IMPORTANT: All times must be in system LOCAL timezone (not UTC).
    
    Args:
        subject: Meeting title (required)
        start_time: datetime in LOCAL timezone (required, not UTC)
        end_time: datetime in LOCAL timezone (required, not UTC)
        location: Meeting location string (optional)
        body: Meeting description/body text (optional)
        required_attendees: List of email addresses (optional)
        optional_attendees: List of email addresses (optional)
        is_online: Enable online meeting (Teams/Skype) (default False)
        online_url: Custom online meeting URL (optional)
        
    Returns:
        Appointment object if successful, None if failed
    """
    manager = OutlookManager()
    if not manager.connect():
        return None
    
    try:
        # Create appointment item
        appointment = manager.create_appointment()
        
        # Set required properties
        appointment.Subject = subject
        appointment.Start = start_time  # COM expects LOCAL time
        appointment.End = end_time      # COM expects LOCAL time
        
        # Set optional properties
        if location:
            appointment.Location = location
        
        if body:
            appointment.Body = body
        
        # Set online meeting properties
        if is_online:
            appointment.IsOnlineMeeting = True
            if online_url:
                appointment.OnlineMeetingUrl = online_url
        
        # Add attendees
        if required_attendees:
            for email in required_attendees:
                recipient = appointment.Recipients.Add(email)
                recipient.Role = 1  # 1 = required
        
        if optional_attendees:
            for email in optional_attendees:
                recipient = appointment.Recipients.Add(email)
                recipient.Role = 2  # 2 = optional
        
        # Resolve all attendee names
        if required_attendees or optional_attendees:
            appointment.Recipients.ResolveAll()
        
        # Save to calendar
        appointment.Save()
        return appointment
        
    except Exception as e:
        print(f"Error creating meeting: {e}")
        return None


def create_recurring_meeting(
    subject,
    start_time,
    end_time,
    recurrence_type,
    interval=1,
    end_date=None,
    occurrences=None,
    location=None,
    body=None,
    attendees=None,
):
    """
    Create a recurring Outlook meeting.
    
    IMPORTANT: start_time and end_time must be in system LOCAL timezone.
    
    Args:
        subject: Meeting title
        start_time: datetime in LOCAL timezone (first occurrence)
        end_time: datetime in LOCAL timezone (first occurrence)
        recurrence_type: 0=daily, 1=weekly, 2=monthly, 3=yearly (required)
        interval: Repeat interval (e.g., 2 for every 2 weeks) (default 1)
        end_date: When to stop recurring (datetime, LOCAL time) (optional)
        occurrences: Max number of occurrences (optional, use either this or end_date)
        location: Meeting location (optional)
        body: Meeting description (optional)
        attendees: List of email addresses (optional)
        
    Returns:
        Appointment object if successful, None if failed
    """
    manager = OutlookManager()
    if not manager.connect():
        return None
    
    try:
        appointment = manager.create_appointment()
        
        # Set basic properties
        appointment.Subject = subject
        appointment.Start = start_time
        appointment.End = end_time
        
        if location:
            appointment.Location = location
        
        if body:
            appointment.Body = body
        
        # Configure recurrence
        recurrence = appointment.GetRecurrencePattern()
        recurrence.RecurrenceType = recurrence_type
        recurrence.Interval = interval
        
        if end_date:
            recurrence.PatternEndDate = end_date
        
        if occurrences:
            recurrence.Occurrences = occurrences
        
        # Add attendees if provided
        if attendees:
            for email in attendees:
                appointment.Recipients.Add(email)
            appointment.Recipients.ResolveAll()
        
        # Save to calendar
        appointment.Save()
        return appointment
        
    except Exception as e:
        print(f"Error creating recurring meeting: {e}")
        return None


def create_online_meeting(
    subject,
    start_time,
    end_time,
    location=None,
    body=None,
    attendees=None,
):
    """
    Create an online meeting (Teams/Skype) in Outlook.
    
    The Teams URL is auto-generated after saving if attendees are present.
    
    IMPORTANT: Times must be in system LOCAL timezone (not UTC).
    
    Args:
        subject: Meeting title
        start_time: datetime in LOCAL timezone
        end_time: datetime in LOCAL timezone
        location: Physical location if hybrid (optional)
        body: Meeting description (optional)
        attendees: List of email addresses (recommended for URL generation)
        
    Returns:
        Appointment object with OnlineMeetingUrl populated, None if failed
    """
    manager = OutlookManager()
    if not manager.connect():
        return None
    
    try:
        appointment = manager.create_appointment()
        
        # Set basic properties
        appointment.Subject = subject
        appointment.Start = start_time
        appointment.End = end_time
        
        if location:
            appointment.Location = location
        
        if body:
            appointment.Body = body
        
        # Enable online meeting
        appointment.IsOnlineMeeting = True
        
        # Add attendees (required for Teams URL generation)
        if attendees:
            for email in attendees:
                appointment.Recipients.Add(email)
            appointment.Recipients.ResolveAll()
        
        # Save first so Teams URL is generated
        appointment.Save()
        
        # URL is now available
        meeting_url = appointment.OnlineMeetingUrl
        if meeting_url:
            print(f"Online meeting created. Join at: {meeting_url}")
        
        return appointment
        
    except Exception as e:
        print(f"Error creating online meeting: {e}")
        return None


def validate_meeting_times(start_time, end_time):
    """
    Validate meeting time parameters before creation.
    
    Args:
        start_time: Start datetime
        end_time: End datetime
        
    Returns:
        tuple: (is_valid, error_message)
    """
    if start_time >= end_time:
        return False, "Start time must be before end time"
    
    duration = end_time - start_time
    
    if duration.total_seconds() < 300:  # 5 minutes minimum
        return False, "Meeting duration too short (minimum 5 minutes)"
    
    if duration.total_seconds() > 28800:  # 8 hours maximum
        return False, "Meeting duration too long (maximum 8 hours)"
    
    return True, ""
