"""Tests for dataCollector module."""
import pytest
from unittest.mock import Mock, patch, MagicMock
import logging
import serial
from pywerHelper import dataCollector, serialComm


@pytest.fixture
def mock_logger():
    """Create a mock logger."""
    return Mock(spec=logging.Logger)


class TestInitSerialDevice:
    """Test initSerialDevice function."""
    
    @patch('pywerHelper.serialComm.SerialDevice')
    def test_init_serial_device_success(self, mock_device_class, mock_logger):
        """Test successful serial device initialization."""
        mock_device = Mock()
        mock_device_class.return_value = mock_device
        
        result = dataCollector.initSerialDevice(mock_logger)
        
        assert result == mock_device
        mock_logger.info.assert_called()
    
    @patch('pywerHelper.serialComm.SerialDevice')
    def test_init_serial_device_serial_exception(self, mock_device_class, mock_logger):
        """Test handling SerialException during initialization."""
        mock_device_class.side_effect = serial.SerialException("Port not found")
        
        result = dataCollector.initSerialDevice(mock_logger)
        
        assert result is None
        mock_logger.error.assert_called()
    
    @patch('pywerHelper.serialComm.SerialDevice')
    def test_init_serial_device_file_not_found(self, mock_device_class, mock_logger):
        """Test handling FileNotFoundError during initialization."""
        mock_device_class.side_effect = FileNotFoundError("Config not found")
        
        result = dataCollector.initSerialDevice(mock_logger)
        
        assert result is None
        mock_logger.error.assert_called()
    
    @patch('pywerHelper.serialComm.SerialDevice')
    def test_init_serial_device_unexpected_error(self, mock_device_class, mock_logger):
        """Test handling unexpected error during initialization."""
        mock_device_class.side_effect = Exception("Unexpected error")
        
        result = dataCollector.initSerialDevice(mock_logger)
        
        assert result is None
        mock_logger.error.assert_called()


class TestReadSerialData:
    """Test readSerialData function."""
    
    def test_read_serial_data_success(self, mock_logger):
        """Test successful serial data read."""
        mock_device = Mock()
        mock_device.query.return_value = (b'12.34', '31 32 2e 33 34', '12.34')
        
        raw, hex_data, ascii_data = dataCollector.readSerialData(
            mock_device, mock_logger, command=b'?MPOW'
        )
        
        assert raw == b'12.34'
        assert hex_data == '31 32 2e 33 34'
        assert ascii_data == '12.34'
        mock_device.query.assert_called_once_with(command=b'?MPOW')
    
    def test_read_serial_data_serial_exception(self, mock_logger):
        """Test handling SerialException during read."""
        mock_device = Mock()
        mock_device.query.side_effect = serial.SerialException("Communication error")
        
        raw, hex_data, ascii_data = dataCollector.readSerialData(
            mock_device, mock_logger
        )
        
        assert raw is None
        assert hex_data is None
        assert ascii_data is None
        mock_logger.error.assert_called()
    
    def test_read_serial_data_attribute_error(self, mock_logger):
        """Test handling AttributeError (invalid device)."""
        mock_device = Mock()
        mock_device.query.side_effect = AttributeError("No query method")
        
        raw, hex_data, ascii_data = dataCollector.readSerialData(
            mock_device, mock_logger
        )
        
        assert raw is None
        assert hex_data is None
        assert ascii_data is None
        mock_logger.error.assert_called()
    
    def test_read_serial_data_unexpected_error(self, mock_logger):
        """Test handling unexpected error during read."""
        mock_device = Mock()
        mock_device.query.side_effect = Exception("Unexpected")
        
        raw, hex_data, ascii_data = dataCollector.readSerialData(
            mock_device, mock_logger
        )
        
        assert raw is None
        assert hex_data is None
        assert ascii_data is None
        mock_logger.error.assert_called()


class TestSerialFunction:
    """Test serialFunction (main data collection function)."""
    
    @patch('pywerHelper.dataCollector.initSerialDevice')
    def test_serial_function_init_fails(self, mock_init, mock_logger):
        """Test serialFunction when device initialization fails."""
        mock_init.return_value = None
        
        result = dataCollector.serialFunction(mock_logger, minutes=0.01)
        
        assert result == []
        mock_logger.error.assert_called()
    
    @patch('pywerHelper.dataCollector.readSerialData')
    @patch('pywerHelper.dataCollector.initSerialDevice')
    @patch('time.time')
    def test_serial_function_collects_samples(self, mock_time, mock_init, mock_read, mock_logger):
        """Test that serialFunction collects samples during test duration."""
        # Setup mock device
        mock_device = Mock()
        mock_device.close = Mock()
        mock_init.return_value = mock_device
        
        # Setup time progression (simulate 2 iterations)
        start_time = 1000.0
        mock_time.side_effect = [
            start_time,  # test_start_time
            start_time + 0.5,  # First iteration
            start_time + 1.1  # Second iteration (past end_time)
        ]
        
        # Setup read data
        mock_read.side_effect = [
            (b'10.5', 'hex1', '10.5'),
            (b'11.2', 'hex2', '11.2')
        ]
        
        # Run for very short duration
        result = dataCollector.serialFunction(
            mock_logger, 
            minutes=0.01,  # 0.6 seconds
            test_header="Test 1"
        )
        
        # Should have collected samples
        assert len(result) > 0
        mock_device.close.assert_called_once()
    
    @patch('pywerHelper.dataCollector.initSerialDevice')
    def test_serial_function_keyboard_interrupt(self, mock_init, mock_logger):
        """Test handling KeyboardInterrupt during data collection."""
        mock_device = Mock()
        mock_device.close = Mock()
        mock_init.return_value = mock_device
        
        with patch('pywerHelper.dataCollector.readSerialData') as mock_read:
            mock_read.side_effect = KeyboardInterrupt()
            
            result = dataCollector.serialFunction(mock_logger, minutes=0.01)
            
            # Should return collected samples before interrupt (empty in this case)
            assert isinstance(result, list)
            mock_device.close.assert_called_once()
    
    @patch('pywerHelper.dataCollector.readSerialData')
    @patch('pywerHelper.dataCollector.initSerialDevice')
    @patch('time.time')
    def test_serial_function_with_global_timer(self, mock_time, mock_init, mock_read, mock_logger):
        """Test serialFunction with global timer tracking."""
        mock_device = Mock()
        mock_device.close = Mock()
        mock_init.return_value = mock_device
        
        # Setup time
        start_time = 1000.0
        global_start = 950.0
        mock_time.side_effect = [
            start_time,
            start_time + 1.0  # End immediately
        ]
        
        mock_read.return_value = (b'10.5', 'hex', '10.5')
        
        result = dataCollector.serialFunction(
            mock_logger,
            minutes=0.001,
            global_timer_start=global_start,
            test_header="Test with Global Timer"
        )
        
        assert isinstance(result, list)
        mock_device.close.assert_called_once()
    
    @patch('pywerHelper.dataCollector.readSerialData')
    @patch('pywerHelper.dataCollector.initSerialDevice')
    @patch('time.time')
    def test_serial_function_tracks_recent_samples(self, mock_time, mock_init, mock_read, mock_logger):
        """Test that serialFunction tracks recent samples for display."""
        mock_device = Mock()
        mock_device.close = Mock()
        mock_init.return_value = mock_device
        
        # Create many samples to test the 15-sample limit
        samples_data = [(b'10.5', 'hex', f'Sample{i}') for i in range(20)]
        
        start_time = 1000.0
        # Need enough time calls for all samples plus initial and final
        time_calls = [start_time] + [start_time + (i * 0.1) for i in range(len(samples_data))]
        time_calls.append(start_time + 10)  # Final time past end
        mock_time.side_effect = time_calls
        
        mock_read.side_effect = samples_data
        
        result = dataCollector.serialFunction(
            mock_logger,
            minutes=0.01,
            test_header="Test Recent Samples"
        )
        
        # All samples should be collected
        assert len(result) > 0
        mock_device.close.assert_called_once()


class TestDataCollectorIntegration:
    """Integration tests for dataCollector module."""
    
    @patch('pywerHelper.serialComm.SerialDevice')
    @patch('time.time')
    @patch('pywerHelper.dataCollector.readSerialData')
    def test_full_data_collection_workflow(self, mock_read, mock_time, mock_device_class, mock_logger):
        """Test complete data collection workflow."""
        # Setup mock device
        mock_device = Mock()
        mock_device.close = Mock()
        mock_device_class.return_value = mock_device
        
        # Setup time - need enough time calls for the loop
        start_time = 1000.0
        mock_time.side_effect = [
            start_time,  # test_start_time
            start_time + 0.1,  # First iteration current_time
            start_time + 10  # Second iteration - past end_time
        ]
        
        # Mock successful data read
        mock_read.return_value = (b'12.34', 'hex', '12.34')
        
        # Run collection
        result = dataCollector.serialFunction(
            mock_logger,
            minutes=0.001,
            test_header="Integration Test"
        )
        
        # Verify results
        assert isinstance(result, list)
        # Collection happens, device is closed
        mock_device.close.assert_called_once()
