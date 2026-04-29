"""Time parsing and formatting utilities for pywerMeter."""
from datetime import datetime
from typing import Union

# Try to import pytz for timezone support
try:
    import pytz
    PYTZ_AVAILABLE = True
except ImportError:
    PYTZ_AVAILABLE = False


def parse_time_value(time_value: Union[int, float, str, None]) -> float:
    """
    Parse time value from config.
    
    Args:
        time_value: Numeric (minutes) or string "M:SS" format
        
    Returns:
        Time in minutes as float
        
    Examples:
        >>> parse_time_value(1.5)
        1.5
        >>> parse_time_value("1:30")
        1.5
        >>> parse_time_value("5:00")
        5.0
    """
    if time_value is None:
        return 0.0
    
    if isinstance(time_value, (int, float)):
        return float(time_value)
    
    if isinstance(time_value, str) and ':' in time_value:
        parts = time_value.split(':')
        if len(parts) == 2:
            try:
                minutes = int(parts[0])
                seconds = int(parts[1])
                return minutes + (seconds / 60.0)
            except ValueError:
                print(f"Warning: Invalid time format '{time_value}'. Using 0.")
                return 0.0
    
    print(f"Warning: Unrecognized time format '{time_value}'. Using 0.")
    return 0.0


def format_time_minutes(minutes: float) -> str:
    """
    Format minutes as M:SS string.
    
    Args:
        minutes: Time in minutes
        
    Returns:
        Formatted string like "5:30"
        
    Examples:
        >>> format_time_minutes(5.5)
        '5:30'
        >>> format_time_minutes(1.25)
        '1:15'
    """
    m = int(minutes)
    s = int((minutes - m) * 60)
    return f"{m}:{s:02d}"


def get_formatted_start_time(elapsed_minutes: float) -> str:
    """
    Get formatted start time with clock time and elapsed minutes.
    
    Args:
        elapsed_minutes: Elapsed time in minutes from global timer start
        
    Returns:
        Formatted string like "14:30:15 / 5.50 min"
    """
    if PYTZ_AVAILABLE:
        try:
            est = pytz.timezone('US/Eastern')
            actual_start_time = datetime.now(est)
        except:
            actual_start_time = datetime.now()
    else:
        actual_start_time = datetime.now()
    
    return f"{actual_start_time.strftime('%H:%M:%S')} / {elapsed_minutes:.2f} min"


__all__ = ['parse_time_value', 'format_time_minutes', 'get_formatted_start_time', 'PYTZ_AVAILABLE']