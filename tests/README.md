# PywerMeter Test Suite

This directory contains comprehensive pytest tests for the pywerMeter application.

## Test Coverage

### Core Modules Tested

1. **test_config_helper.py** - Tests for ConfigManager class
   - Configuration loading and lazy loading
   - Getting values with defaults
   - Test settings, serial settings, log settings retrieval
   - Error handling (missing files, invalid YAML)
   - Configuration caching

2. **test_time_utils.py** - Tests for time utilities
   - Time parsing (integers, floats, M:SS format)
   - Time formatting to M:SS string
   - Formatted start time generation
   - Round-trip conversion consistency
   - Timezone handling (with/without pytz)

3. **test_excel_structure.py** - Tests for ExcelFileStructure class
   - Excel file structure detection (averages, start times, regular)
   - Column letter conversion with caching
   - Header retrieval and column finding
   - Data row management
   - Column data clearing and writing
   - Average formula updates
   - Start time setting

4. **test_menu_helper.py** - Tests for MenuItem and MenuSystem
   - MenuItem dataclass creation
   - MenuSystem item management
   - Menu display formatting
   - Menu execution with valid/invalid choices
   - Method chaining and callable actions

5. **test_serial_comm.py** - Tests for SerialDevice and SerialDeviceBuilder
   - SerialDevice creation, open, close operations
   - Reading and writing serial data
   - SerialDeviceBuilder pattern implementation
   - Builder method chaining
   - Configuration-based builder creation

6. **test_test_runner.py** - Tests for TestRunner class
   - Test execution orchestration
   - Elapsed time calculation with/without pauses
   - Wait until functionality
   - Pause handler integration
   - Global timer adjustment
   - Full test workflow integration

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

# Run all tests with detailed output
python -m pytest tests/ -v --tb=short

# Run specific test file
python -m pytest tests/test_config_helper.py -v

# Run specific test class
python -m pytest tests/test_time_utils.py::TestParseTimeValue -v

# Run specific test method
python -m pytest tests/test_menu_helper.py::TestMenuSystem::test_execute_valid_choice -v
```

### Run Tests with Coverage

```powershell
# Generate coverage report
python -m pytest tests/ --cov=pywerHelper --cov-report=html

# View coverage report (opens in browser)
# Coverage report will be in htmlcov/index.html

# Generate terminal coverage report
python -m pytest tests/ --cov=pywerHelper --cov-report=term-missing
```

### Run Tests in Watch Mode

For continuous testing during development:

```powershell
# Install pytest-watch if needed
pip install pytest-watch

# Run in watch mode
ptw tests/ -- -v
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
    
    def test_error_handling(self):
        """Test error scenarios."""
        # Test implementation
```

## Writing New Tests

### Guidelines

1. **Test naming**: Use descriptive names starting with `test_`
2. **Docstrings**: Include brief description of what's being tested
3. **AAA pattern**: Follow Arrange-Act-Assert pattern
4. **Fixtures**: Use fixtures for common setup/teardown
5. **Mocking**: Use `unittest.mock` for external dependencies (serial, files)
6. **Cleanup**: Use fixtures with yield for proper cleanup

### Example Test

```python
import pytest
from pywerHelper.mymodule import MyClass

class TestMyClass:
    """Test MyClass functionality."""
    
    def test_my_feature(self):
        """Test that my feature works correctly."""
        # Arrange
        obj = MyClass()
        
        # Act
        result = obj.my_method(input_data)
        
        # Assert
        assert result == expected_output
    
    def test_error_handling(self):
        """Test that errors are handled properly."""
        obj = MyClass()
        
        with pytest.raises(ValueError):
            obj.my_method(invalid_input)
```

## Test Categories

### Unit Tests
- Test individual functions and methods in isolation
- Use mocking for external dependencies
- Fast execution

### Integration Tests
- Test interaction between multiple components
- May use temporary files or mock resources
- Slightly slower execution

### Edge Cases
- Test boundary conditions
- Test with invalid inputs
- Test error handling

## Continuous Integration

These tests are designed to be run in CI/CD pipelines:

```yaml
# Example GitHub Actions workflow
- name: Run tests
  run: |
    python -m pytest tests/ -v --cov=pywerHelper --cov-report=xml
```

## Coverage Goals

Target coverage metrics:

- **Line Coverage**: > 80%
- **Branch Coverage**: > 70%
- **Critical paths**: 100%

## Known Test Limitations

1. **Serial communication**: Tests use mocks, not real hardware
2. **Excel files**: Tests use temporary files, not production data
3. **Time-dependent tests**: May have slight timing variations
4. **Timezone tests**: Require pytz to be installed

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

### Flaky Tests

If tests fail intermittently:
- Check for timing-dependent assertions
- Increase tolerance in float comparisons
- Use proper mocking for time-dependent code

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
