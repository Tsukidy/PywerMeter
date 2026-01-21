"""Tests for serialComm module."""
import pytest
from unittest.mock import Mock, patch, MagicMock
from pywerHelper.serialComm import SerialDevice, SerialDeviceBuilder


class TestSerialDevice:
    """Test SerialDevice class."""
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_create_serial_device(self, mock_serial):
        """Test creating a SerialDevice."""
        mock_instance = MagicMock()
        mock_serial.return_value = mock_instance
        
        device = SerialDevice("COM3", 9600, 1)
        
        assert device.port == "COM3"
        assert device.baudrate == 9600
        assert device.timeout == 1
        mock_serial.assert_called_once_with("COM3", 9600, timeout=1)
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_serial_device_open(self, mock_serial):
        """Test opening serial connection."""
        mock_instance = MagicMock()
        mock_instance.is_open = False
        mock_serial.return_value = mock_instance
        
        device = SerialDevice("COM3", 9600, 1)
        device.open()
        
        mock_instance.open.assert_called_once()
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_serial_device_close(self, mock_serial):
        """Test closing serial connection."""
        mock_instance = MagicMock()
        mock_instance.is_open = True
        mock_serial.return_value = mock_instance
        
        device = SerialDevice("COM3", 9600, 1)
        device.close()
        
        mock_instance.close.assert_called_once()
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_serial_device_write(self, mock_serial):
        """Test writing to serial device."""
        mock_instance = MagicMock()
        mock_serial.return_value = mock_instance
        
        device = SerialDevice("COM3", 9600, 1)
        device.write(b"TEST")
        
        mock_instance.write.assert_called_once_with(b"TEST")
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_serial_device_read(self, mock_serial):
        """Test reading from serial device."""
        mock_instance = MagicMock()
        mock_instance.readline.return_value = b"response\n"
        mock_serial.return_value = mock_instance
        
        device = SerialDevice("COM3", 9600, 1)
        data = device.read()
        
        assert data == b"response\n"
        mock_instance.readline.assert_called_once()
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_serial_device_is_open(self, mock_serial):
        """Test checking if serial device is open."""
        mock_instance = MagicMock()
        mock_instance.is_open = True
        mock_serial.return_value = mock_instance
        
        device = SerialDevice("COM3", 9600, 1)
        
        assert device.is_open() is True


class TestSerialDeviceBuilder:
    """Test SerialDeviceBuilder class."""
    
    def test_builder_default_values(self):
        """Test builder with default values."""
        builder = SerialDeviceBuilder()
        
        assert builder._port is None
        assert builder._baudrate == 9600
        assert builder._timeout == 1
    
    def test_builder_set_port(self):
        """Test setting port."""
        builder = SerialDeviceBuilder()
        result = builder.set_port("COM5")
        
        assert builder._port == "COM5"
        assert result is builder  # Should return self for chaining
    
    def test_builder_set_baudrate(self):
        """Test setting baudrate."""
        builder = SerialDeviceBuilder()
        result = builder.set_baudrate(115200)
        
        assert builder._baudrate == 115200
        assert result is builder
    
    def test_builder_set_timeout(self):
        """Test setting timeout."""
        builder = SerialDeviceBuilder()
        result = builder.set_timeout(5)
        
        assert builder._timeout == 5
        assert result is builder
    
    def test_builder_chaining(self):
        """Test method chaining."""
        builder = (SerialDeviceBuilder()
                   .set_port("COM7")
                   .set_baudrate(19200)
                   .set_timeout(2))
        
        assert builder._port == "COM7"
        assert builder._baudrate == 19200
        assert builder._timeout == 2
    
    @patch('pywerHelper.serialComm.SerialDevice')
    def test_builder_build(self, mock_device_class):
        """Test building SerialDevice."""
        mock_device_class.return_value = Mock()
        
        builder = (SerialDeviceBuilder()
                   .set_port("COM3")
                   .set_baudrate(9600)
                   .set_timeout(1))
        
        device = builder.build()
        
        mock_device_class.assert_called_once_with("COM3", 9600, 1)
    
    def test_builder_build_without_port(self):
        """Test building without setting port."""
        builder = SerialDeviceBuilder()
        
        with pytest.raises(ValueError, match="Port must be set"):
            builder.build()
    
    def test_from_config(self):
        """Test creating builder from config dict."""
        config = {
            "port": "COM4",
            "baudrate": 115200,
            "timeout": 3
        }
        
        builder = SerialDeviceBuilder.from_config(config)
        
        assert builder._port == "COM4"
        assert builder._baudrate == 115200
        assert builder._timeout == 3
    
    def test_from_config_partial(self):
        """Test creating builder from partial config."""
        config = {
            "port": "COM4"
            # baudrate and timeout will use defaults
        }
        
        builder = SerialDeviceBuilder.from_config(config)
        
        assert builder._port == "COM4"
        assert builder._baudrate == 9600  # Default
        assert builder._timeout == 1  # Default
    
    def test_from_config_empty(self):
        """Test creating builder from empty config."""
        config = {}
        
        builder = SerialDeviceBuilder.from_config(config)
        
        assert builder._port is None
        assert builder._baudrate == 9600
        assert builder._timeout == 1
    
    @patch('pywerHelper.serialComm.SerialDevice')
    def test_from_config_to_build(self, mock_device_class):
        """Test full workflow from config to build."""
        mock_device_class.return_value = Mock()
        
        config = {
            "port": "COM8",
            "baudrate": 19200,
            "timeout": 2
        }
        
        device = SerialDeviceBuilder.from_config(config).build()
        
        mock_device_class.assert_called_once_with("COM8", 19200, 2)
