"""Tests for timeUtils module."""
import pytest
from pywerHelper.timeUtils import parse_time_value, format_time_minutes, get_formatted_start_time, PYTZ_AVAILABLE


class TestParseTimeValue:
    """Test parse_time_value function."""
    
    def test_parse_integer(self):
        """Test parsing integer minutes."""
        assert parse_time_value(5) == 5.0
        assert parse_time_value(0) == 0.0
        assert parse_time_value(120) == 120.0
    
    def test_parse_float(self):
        """Test parsing float minutes."""
        assert parse_time_value(1.5) == 1.5
        assert parse_time_value(2.75) == 2.75
        assert parse_time_value(0.5) == 0.5
    
    def test_parse_mss_format(self):
        """Test parsing M:SS string format."""
        assert parse_time_value("1:30") == 1.5
        assert parse_time_value("5:00") == 5.0
        assert parse_time_value("0:30") == 0.5
        assert parse_time_value("10:15") == 10.25
        assert parse_time_value("2:45") == 2.75
    
    def test_parse_none(self):
        """Test parsing None value."""
        assert parse_time_value(None) == 0.0
    
    def test_parse_invalid_format(self):
        """Test parsing invalid formats."""
        assert parse_time_value("invalid") == 0.0
        assert parse_time_value("1:2:3") == 0.0
        assert parse_time_value("abc:def") == 0.0
    
    def test_parse_edge_cases(self):
        """Test edge cases."""
        assert parse_time_value("0:00") == 0.0
        assert parse_time_value("60:00") == 60.0
        assert parse_time_value("0:01") == pytest.approx(0.0167, abs=0.001)


class TestFormatTimeMinutes:
    """Test format_time_minutes function."""
    
    def test_format_whole_minutes(self):
        """Test formatting whole minutes."""
        assert format_time_minutes(0) == "0:00"
        assert format_time_minutes(1) == "1:00"
        assert format_time_minutes(5) == "5:00"
        assert format_time_minutes(60) == "60:00"
    
    def test_format_with_seconds(self):
        """Test formatting minutes with seconds."""
        assert format_time_minutes(1.5) == "1:30"
        assert format_time_minutes(2.75) == "2:45"
        assert format_time_minutes(0.5) == "0:30"
        assert format_time_minutes(10.25) == "10:15"
    
    def test_format_negative(self):
        """Test formatting negative values."""
        assert format_time_minutes(-1) == "-1:00"
        assert format_time_minutes(-1.5) == "-1:30"
    
    def test_format_large_values(self):
        """Test formatting large time values."""
        assert format_time_minutes(120) == "120:00"
        assert format_time_minutes(99.5) == "99:30"
    
    def test_format_small_fractions(self):
        """Test formatting small fractions of minutes."""
        result = format_time_minutes(0.1)
        assert result == "0:06"
        
        result = format_time_minutes(0.25)
        assert result == "0:15"


class TestGetFormattedStartTime:
    """Test get_formatted_start_time function."""
    
    def test_format_zero_elapsed(self):
        """Test formatting with zero elapsed time."""
        result = get_formatted_start_time(0)
        assert "0:00" in result or "0.00" in result
    
    def test_format_positive_elapsed(self):
        """Test formatting with positive elapsed time."""
        result = get_formatted_start_time(5.5)
        
        # Should contain both time format and minutes
        assert ":" in result  # HH:MM:SS format
        assert "5.50" in result or "5.5" in result  # minutes
    
    def test_format_large_elapsed(self):
        """Test formatting with large elapsed time."""
        result = get_formatted_start_time(125.75)
        
        # Should handle hours correctly
        assert ":" in result
        assert "125.75" in result
    
    @pytest.mark.skipif(not PYTZ_AVAILABLE, reason="pytz not available")
    def test_format_with_pytz(self):
        """Test formatting when pytz is available."""
        result = get_formatted_start_time(10.0)
        
        # Should contain timestamp in Eastern time if pytz available
        assert "/" in result or "min" in result.lower()
    
    def test_format_without_pytz(self, monkeypatch):
        """Test formatting when pytz is not available."""
        # Temporarily disable pytz
        import pywerHelper.timeUtils as time_utils
        monkeypatch.setattr(time_utils, "PYTZ_AVAILABLE", False)
        
        result = get_formatted_start_time(10.0)
        
        # Should still return valid format
        assert "/" in result or "min" in result.lower()


class TestRoundTripConversion:
    """Test round-trip conversion between parse and format."""
    
    def test_round_trip_whole_minutes(self):
        """Test parsing and formatting whole minutes."""
        original = 5.0
        formatted = format_time_minutes(original)
        parsed = parse_time_value(formatted)
        assert parsed == original
    
    def test_round_trip_with_seconds(self):
        """Test parsing and formatting with seconds."""
        test_values = [1.5, 2.25, 3.75, 10.5]
        
        for value in test_values:
            formatted = format_time_minutes(value)
            parsed = parse_time_value(formatted)
            assert parsed == pytest.approx(value, abs=0.01)
    
    def test_consistency(self):
        """Test consistency between parse and format."""
        # These should be consistent
        assert parse_time_value("1:30") == 1.5
        assert format_time_minutes(1.5) == "1:30"
        
        assert parse_time_value("5:00") == 5.0
        assert format_time_minutes(5.0) == "5:00"
