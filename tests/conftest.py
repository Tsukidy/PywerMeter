"""Pytest configuration and shared fixtures."""
import pytest
import os
import sys
import logging

# Add project root to Python path for imports
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)


@pytest.fixture(autouse=True)
def suppress_logging():
    """Suppress logging output during tests."""
    logging.disable(logging.CRITICAL)
    yield
    logging.disable(logging.NOTSET)


@pytest.fixture
def sample_config_dict():
    """Provide a sample configuration dictionary for testing."""
    return {
        "log_file_location": "logs/",
        "test_1_start_time": 0,
        "test_1_duration": 1.5,
        "test_2_start_time": 5,
        "test_2_duration": 2.0,
        "test_3_start_time": "10:30",
        "test_3_duration": "3:00",
        "test_4_start_time": 15,
        "test_4_duration": 4.0,
        "test_fast_start_1": "On",
        "test_fast_start_2": "Off",
        "test_fast_start_3": "On",
        "test_fast_start_4": "Off",
        "adjust_time_without_off": "On",
        "serial_port": "COM3",
        "baud_rate": 9600,
        "timeout": 1,
        "command_send_interval": 1,
        "command_read_delay": 0.1,
    }


@pytest.fixture
def mock_serial_data():
    """Provide mock serial data samples."""
    return [
        b"10.5\r\n",
        b"11.2\r\n",
        b"10.8\r\n",
        b"11.0\r\n",
        b"10.9\r\n",
    ]
