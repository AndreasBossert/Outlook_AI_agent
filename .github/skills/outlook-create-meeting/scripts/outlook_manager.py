"""
Outlook Connection Manager

General-purpose module for establishing and managing Outlook COM connections.
"""

import win32com.client


class OutlookManager:
    """Manages Outlook COM connection and provides access to calendar."""
    
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.calendar = None
    
    def connect(self):
        """Establish connection to Outlook."""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.calendar = self.namespace.GetDefaultFolder(9)  # 9 = Calendar folder
            return True
        except Exception as e:
            print(f"Failed to connect to Outlook: {e}")
            return False
    
    def disconnect(self):
        """Close Outlook connection."""
        self.outlook = None
        self.namespace = None
        self.calendar = None
    
    def get_calendar(self):
        """Get calendar object (connects if needed)."""
        if not self.calendar:
            self.connect()
        return self.calendar
    
    def create_appointment(self):
        """Create a new appointment item."""
        if not self.outlook:
            self.connect()
        return self.outlook.CreateItem(1)  # 1 = olAppointmentItem
    
    def get_calendar_items(self):
        """Get all calendar items sorted by start time."""
        calendar = self.get_calendar()
        items = calendar.Items
        items.Sort("[Start]")
        return items


# Convenience function for simple one-off connections
def get_outlook_calendar():
    """Quick helper: get calendar without managing connection object."""
    manager = OutlookManager()
    if manager.connect():
        return manager.calendar
    return None
