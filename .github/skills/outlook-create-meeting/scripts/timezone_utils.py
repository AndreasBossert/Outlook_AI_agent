"""
Timezone utilities for Outlook COM API

Handles conversion between UTC and local time with proper timezone awareness.
The Windows COM API stores times as UTC internally but returns them as local time.
"""

from datetime import datetime
import pytz


def get_system_timezone():
    """Get the system's local timezone."""
    return datetime.now().astimezone().tzinfo


def utc_to_local(utc_datetime):
    """
    Convert UTC datetime to system local time.
    
    Args:
        utc_datetime: datetime object with UTC timezone info
        
    Returns:
        datetime object in local timezone (naive)
    """
    if utc_datetime.tzinfo is None:
        utc_datetime = utc_datetime.replace(tzinfo=pytz.UTC)
    
    local_tz = get_system_timezone()
    local_dt = utc_datetime.astimezone(local_tz)
    return local_dt.replace(tzinfo=None)


def local_to_utc(local_datetime):
    """
    Convert system local time to UTC.
    
    Args:
        local_datetime: datetime object (naive, assumed to be local)
        
    Returns:
        datetime object in UTC timezone (aware)
    """
    local_tz = get_system_timezone()
    aware_dt = local_datetime.replace(tzinfo=local_tz)
    utc_dt = aware_dt.astimezone(pytz.UTC)
    return utc_dt


def make_timezone_aware(naive_datetime, tz=None):
    """
    Add timezone info to a naive datetime (assumes local time if tz not specified).
    
    Args:
        naive_datetime: datetime object without timezone info
        tz: timezone to apply (defaults to system local timezone)
        
    Returns:
        timezone-aware datetime object
    """
    if tz is None:
        tz = get_system_timezone()
    return naive_datetime.replace(tzinfo=tz)


def convert_to_timezone(aware_datetime, target_tz_name):
    """
    Convert timezone-aware datetime to a different timezone.
    
    Args:
        aware_datetime: datetime object with timezone info
        target_tz_name: timezone name string (e.g., 'US/Eastern', 'Europe/London')
        
    Returns:
        datetime object in target timezone (naive)
    """
    target_tz = pytz.timezone(target_tz_name)
    converted = aware_datetime.astimezone(target_tz)
    return converted.replace(tzinfo=None)


# COM API special handling
def com_to_aware_local(com_datetime):
    """
    Convert COM-returned datetime (already local) to timezone-aware format.
    
    The COM API returns local time automatically, so this just adds timezone info.
    
    Args:
        com_datetime: datetime from Outlook COM (naive local time)
        
    Returns:
        timezone-aware datetime in local timezone
    """
    local_tz = get_system_timezone()
    return com_datetime.replace(tzinfo=local_tz)


def aware_local_to_utc(aware_local):
    """
    Convert timezone-aware local datetime to UTC.
    
    Args:
        aware_local: timezone-aware datetime in local timezone
        
    Returns:
        timezone-aware datetime in UTC
    """
    return aware_local.astimezone(pytz.UTC)
