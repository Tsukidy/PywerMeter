"""Tests for serialComm module."""
import pytest
from unittest.mock import Mock, patch, MagicMock
import serial
from pywerHelper.serialComm import SerialDevice, SerialDeviceBuilder, returnSerialPorts


class TestReturnSerialPorts:
    """Test returnSerialPorts function."""
    
    @patch('pywerHelper.serialComm.serial.tools.list_ports.comports')
    def test_list_ports_with_available_ports(self, mock_comports):
        """Test listing ports when ports are available."""
        # Mock port objects
        mock_port1 = Mock()
        mock_port1.device = "COM3"
        mock_port1.description = "USB Serial Port"
        
        mock_port2 = Mock()
        mock_port2.device = "COM4"
        mock_port2.description = "Arduino"
        
        mock_comports.return_value = [mock_port1, mock_port2]
        
        ports = returnSerialPorts()
        
        assert len(ports) == 2
        assert ports[0].device == "COM3"
        assert ports[1].device == "COM4"
    
    @patch('pywerHelper.serialComm.serial.tools.list_ports.comports')
    def test_list_ports_no_ports_available(self, mock_comports):
        """Test listing ports when no ports are available."""
        mock_comports.return_value = []
        
        ports = returnSerialPorts()
        
        assert ports == []
    
    @patch('pywerHelper.serialComm.serial.tools.list_ports.comports')
    def test_list_ports_exception(self, mock_comports):
        """Test handling exception during port enumeration."""
        mock_comports.side_effect = Exception("USB error")
        
        ports = returnSerialPorts()
        
        assert ports == []


class TestSerialDeviceBuilder:
    """Test SerialDeviceBuilder class."""
    
    def test_builder_default_values(self):
        """Test builder with default values."""
        builder = SerialDeviceBuilder()
        
        assert builder.port == "COM9"
        assert builder.baudrate == 9600
        assert builder.timeout == 1
    
    def test_builder_set_port(self):
        """Test setting port."""
        builder = SerialDeviceBuilder()
        result = builder.set_port("COM5")
        
        assert builder.port == "COM5"
        assert result is builder  # Should return self for chaining
    
    def test_builder_set_baudrate(self):
        """Test setting baudrate."""
        builder = SerialDeviceBuilder()
        result = builder.set_baudrate(115200)
        
        assert builder.baudrate == 115200
        assert result is builder
    
    def test_builder_set_timeout(self):
        """Test setting timeout."""
        builder = SerialDeviceBuilder()
        result = builder.set_timeout(5)
        
        assert builder.timeout == 5
        assert result is builder
    
    def test_builder_chaining(self):
        """Test method chaining."""
        builder = (SerialDeviceBuilder()
                   .set_port("COM7")
                   .set_baudrate(19200)
                   .set_timeout(2))
        
        assert builder.port == "COM7"
        assert builder.baudrate == 19200
        assert builder.timeout == 2
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_builder_build(self, mock_serial_class):
        """Test building SerialDevice."""
        mock_serial_instance = MagicMock()
        mock_serial_class.return_value = mock_serial_instance
        
        builder = (SerialDeviceBuilder()
                   .set_port("COM3")
                   .set_baudrate(9600)
                   .set_timeout(1))
        
        device = builder.build()
        
        # Verify SerialDevice was created
        assert device is not None
        mock_serial_class.assert_called_once()


class TestSerialDevice:
    """Test SerialDevice class."""
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_create_serial_device(self, mock_serial_class):
        """Test creating a SerialDevice."""
        mock_instance = MagicMock()
        mock_serial_class.return_value = mock_instance
        
        device = SerialDevice("COM3", 9600, timeout=1)
        
        assert device.ser is not None
        mock_serial_class.assert_called_once()
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_serial_device_init_failure(self, mock_serial_class):
        """Test handling serial device initialization failure."""
        mock_serial_class.side_effect = serial.SerialException("Port not found")
        
        with pytest.raises(serial.SerialException):
            device = SerialDevice("COM999")
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_query_success(self, mock_serial_class):
        """Test successful query operation."""
        mock_instance = MagicMock()
        mock_instance.is_open = True
        mock_instance.read.return_value = b'12.34\r\n'
        mock_serial_class.return_value = mock_instance
        
        device = SerialDevice("COM3")
        raw, hex_str, ascii_str = device.query(b'?MPOW')
        
        assert raw == b'12.34\r\n'
        assert ascii_str == '12.34'
        assert hex_str is not None
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_query_no_response(self, mock_serial_class):
        """Test query with no response from device."""
        mock_instance = MagicMock()
        mock_instance.is_open = True
        mock_instance.read.return_value = b''
        mock_serial_class.return_value = mock_instance
        
        device = SerialDevice("COM3")
        raw, hex_str, ascii_str = device.query(b'?MPOW')
        
        assert raw is None
        assert hex_str is None
        assert ascii_str is None
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_query_port_not_open(self, mock_serial_class):
        """Test query when port is not open."""
        mock_instance = MagicMock()
        mock_instance.is_open = False
        mock_serial_class.return_value = mock_instance
        
        device = SerialDevice("COM3")
        device.ser = mock_instance
        
        with pytest.raises(serial.SerialException):
            device.query(b'?MPOW')
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_close_device(self, mock_serial_class):
        """Test closing serial device."""
        mock_instance = MagicMock()
        mock_instance.is_open = True
        mock_serial_class.return_value = mock_instance
        
        device = SerialDevice("COM3")
        device.close()
        
        mock_instance.close.assert_called_once()
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_close_already_closed(self, mock_serial_class):
        """Test closing device that's already closed."""
        mock_instance = MagicMock()
        mock_instance.is_open = False
        mock_serial_class.return_value = mock_instance
        
        device = SerialDevice("COM3")
        device.ser = mock_instance
        device.close()
        
        # Should not call close on already closed port
        mock_instance.close.assert_not_called()


class TestSerialDeviceIntegration:
    """Integration tests for SerialDevice workflow."""
    
    @patch('pywerHelper.serialComm.serial.Serial')
    def test_builder_to_device_workflow(self, mock_serial_class):
        """Test complete workflow from builder to device operation."""
        mock_instance = MagicMock()
        mock_instance.is_open = True
        mock_instance.read.return_value = b'10.5\r\n'
        mock_serial_class.return_value = mock_instance
        
        # Build device
        device = (SerialDeviceBuilder()
                  .set_port("COM3")
                  .set_baudrate(9600)
                  .build())
        
        # Query device
        raw, hex_str, ascii_str = device.query(b'?MPOW')
        
        assert ascii_str == '10.5'
        
        # Close device
        device.close()
        mock_instance.close.assert_called_once()
