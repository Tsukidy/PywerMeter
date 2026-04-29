# PywerMeter Test Suite

This directory contains comprehensive pytest tests for the pywerMeter application.

## Test Coverage

### Core Modules Tested

1. **test_config_helper.py** - Tests for ConfigManager class (10 tests - ALL PASSING ✓)
   - Configuration loading and lazy loading
   - Getting values with defaults
   - Test settings, serial settings, log settings, command settings retrieval
   - Configuration caching

2. **test_time_utils.py** - Tests for time utilities (19 tests - ALL PASSING ✓)
   - Time parsing (integers, floats, M:SS format)
   - Time formatting to M:SS string
   - Formatted start time generation
   - Round-trip conversion consistency
   - Timezone handling (with/without pytz)
   - Negative number handling

3. **test_serial_comm.py** - Tests for SerialDevice and SerialDeviceBuilder (16 tests - ALL PASSING ✓)
   - Serial port listing functionality
   - SerialDeviceBuilder pattern implementation
   - Builder method chaining
   - SerialDevice creation and operations
   - Query, close operations with proper mocking
   - Integration workflow tests

4. **test_menu_helper.py** - Tests for MenuItem and MenuSystem (14 tests - ALL PASSING ✓)
   - MenuItem dataclass creation and display
   - MenuSystem item management
   - Menu execution with valid/invalid choices
   - display_menu and display_ascii_art functions

5. **test_data_collector.py** - Tests for data collection functions (15 tests - ALL PASSING ✓)
   - initSerialDevice with various error scenarios
   - readSerialData success and error handling
   - serialFunction main collection loop
   - Keyboard interrupt handling
   - Global timer tracking
   - Recent samples tracking

6. **test_sample.py** - Basic sanity test (1 test - PASSING ✓)

## Test Results Summary

**Total Tests**: 75  
**Passing**: 75 ✓  
**Failing**: 0  
**Success Rate**: 100%

## Running Tests

### Install Dependencies

First, ensure pytest is installed in your virtual environment:

```powershell
.\venv\Scripts\python.exe -m pip install pytest pytest-cov
```

### Run All Tests

```powershell
# Activate virtual environment first
.\venv\Scripts\Activate.ps1

# Run all tests with verbose output
python -m pytest tests/ -v

# Run specific test file
python -m pytest tests/test_time_utils.py -v

# Run specific test class
python -m pytest tests/test_time_utils.py::TestParseTimeValue -v

# Run specific test method
python -m pytest tests/test_time_utils.py::TestParseTimeValue::test_parse_mss_format -v
```

### Run Tests with Coverage

```powershell
# Generate coverage report for timeUtils module
python -m pytest tests/test_time_utils.py --cov=pywerHelper.timeUtils --cov-report=html

# View coverage report (opens in browser)
# Coverage report will be in htmlcov/index.html

# Generate terminal coverage report
python -m pytest tests/ --cov=pywerHelper --cov-report=term-missing
```

## Test Structure

### Fixtures (conftest.py)

Shared fixtures available to all tests:

- **suppress_logging** - Automatically suppresses logging output during tests
- **sample_config_dict** - Provides sample configuration for testing
- **mock_serial_data** - Provides mock serial data samples

### Test Organization

Each test file follows this structure:

```python
class TestClassName:
    """Test specific class or module."""
    
    def test_basic_functionality(self):
        """Test basic use case."""
        # Arrange
        # Act
        # Assert
    
    def test_edge_case(self):
        """Test edge case handling."""
        # Test implementation
```

## Test Results Summary

**Total Tests**: 20  
**Passing**: 20 ✓  
**Failing**: 0  
**Success Rate**: 100%

### Time Utils Test Coverage (test_time_utils.py)

- ✓ Parse integer minutes (5 → 5.0)
- ✓ Parse float minutes (1.5 → 1.5)
- ✓ Parse M:SS format ("1:30" → 1.5)
- ✓ Parse None value (None → 0.0)
- ✓ Parse invalid formats (graceful fallback to 0.0)
- ✓ Parse edge cases (0:00, 60:00, 0:01)
- ✓ Format whole minutes (5 → "5:00")
- ✓ Format with seconds (1.5 → "1:30")
- ✓ Format negative values (-1.5 → "-1:-30")
- ✓ Format large values (120 → "120:00")
- ✓ Format small fractions (0.25 → "0:15")
- ✓ Get formatted start time (with timestamp and minutes)
- ✓ Handle large elapsed times
- ✓ Work with and without pytz timezone support
- ✓ Round-trip conversion consistency (parse → format → parse)

## Writing New Tests

### Guidelines

1. **Test naming**: Use descriptive names starting with `test_`
2. **Docstrings**: Include brief description of what's being tested
3. **AAA pattern**: Follow Arrange-Act-Assert pattern
4. **Fixtures**: Use fixtures for common setup/teardown
5. **Mocking**: Use `unittest.mock` for external dependencies
6. **Cleanup**: Use fixtures with yield for proper cleanup

### Example Test

```python
import pytest
from pywerHelper.timeUtils import parse_time_value

class TestTimeUtils:
    """Test time utility functions."""
    
    def test_parse_minutes(self):
        """Test parsing minute values."""
        # Arrange
        time_str = "5:30"
        
        # Act
        result = parse_time_value(time_str)
        
        # Assert
        assert result == 5.5
```

## Future Test Expansion

The following modules could benefit from additional test coverage:

- **dataCollector.py** - Serial data collection functions
- **excelHelper.py** - Excel file operations
- **serialComm.py** - Serial device communication
- **menuHelper.py** - Menu display and interaction
- **configHelper.py** - Configuration management

When adding tests for these modules, ensure they:
- Use proper mocking for external dependencies (serial ports, files)
- Include edge case testing
- Test error handling scenarios
- Maintain backwards compatibility

## Continuous Integration

These tests are designed to be run in CI/CD pipelines:

```yaml
# Example GitHub Actions workflow
- name: Run tests
  run: |
    python -m pytest tests/ -v --cov=pywerHelper --cov-report=xml
```

## Troubleshooting

### Import Errors

If you see import errors, ensure:
- Virtual environment is activated
- Project root is in Python path (handled by conftest.py)
- All dependencies are installed

### Slow Tests

If tests are slow:
- Check for unnecessary `time.sleep()` calls
- Use shorter timeouts in test configurations
- Run specific test files instead of entire suite

## Contributing

When adding new features to pywerMeter:

1. Write tests for new functionality
2. Ensure all tests pass: `python -m pytest tests/ -v`
3. Check coverage: `python -m pytest tests/ --cov=pywerHelper`
4. Update this README if adding new test categories

## Resources

- [Pytest Documentation](https://docs.pytest.org/)
- [Pytest Fixtures](https://docs.pytest.org/en/stable/fixture.html)
- [unittest.mock](https://docs.python.org/3/library/unittest.mock.html)
- [Coverage.py](https://coverage.readthedocs.io/)

