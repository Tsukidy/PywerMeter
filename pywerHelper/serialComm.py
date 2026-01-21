# Python script for handling serial communication with a device.
import serial
import serial.tools.list_ports
import logging
import os
from typing import Optional, Tuple

# Configure logging
logger = logging.getLogger(__name__)

# Ensure log directory exists
try:
    logPath = os.path.join(os.path.dirname(os.path.dirname(__file__)), "logs")
    logName = "serial_communication.log"
    fullLogPath = os.path.join(logPath, logName)
    
    if not os.path.exists(logPath):
        os.makedirs(logPath)
    
    if not os.path.exists(fullLogPath):
        open(fullLogPath, 'a').close()
    
    if not logger.handlers:
        handler = logging.FileHandler(fullLogPath, encoding='utf-8')
        handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
        logger.addHandler(handler)
        logger.setLevel(logging.DEBUG)
        logger.info("Serial communication logger initialized")
except Exception as e:
    print(f"WARNING: Failed to setup logging: {e}")
    logger.addHandler(logging.NullHandler())


def returnSerialPorts():
    """List all available serial ports."""
    try:
        ports = serial.tools.list_ports.comports()
        if not ports:
            print("No serial ports found.")
            logger.warning("No serial ports detected")
            return []
        
        print("Available serial ports:")
        for port in ports:
            print(f"{port.device}: {port.description}")
        logger.info(f"Found {len(ports)} serial port(s)")
        return ports
    except Exception as e:
        print(f"ERROR: Failed to enumerate serial ports: {e}")
        logger.error(f"Failed to enumerate serial ports: {e}", exc_info=True)
        return []


class SerialDeviceBuilder:
    """Builder pattern for SerialDevice configuration."""
    
    def __init__(self):
        self.port = "COM9"
        self.baudrate = 9600
        self.bytesize = serial.EIGHTBITS
        self.stopbits = serial.STOPBITS_ONE
        self.timeout = 1
    
    def from_config(self, config_manager=None):
        """
        Load settings from ConfigManager or config file.
        
        Args:
            config_manager: ConfigManager instance. If None, loads from default location.
        """
        if config_manager is None:
            # Load config directly if no manager provided
            from pywerHelper.configHelper import ConfigManager
            config_manager = ConfigManager()
        
        try:
            settings = config_manager.get_serial_settings()
            
            self.port = settings.get('port', self.port)
            self.baudrate = settings.get('baudrate', self.baudrate)
            self.bytesize = settings.get('bytesize', self.bytesize)
            self.stopbits = settings.get('stopbits', self.stopbits)
            self.timeout = settings.get('timeout', self.timeout)
            
            logger.info(f"Loaded settings: port={self.port}, baudrate={self.baudrate}")
        except Exception as e:
            logger.warning(f"Could not load config: {e}. Using defaults.")
        
        return self
    
    def set_port(self, port: str):
        """Set serial port."""
        self.port = port
        return self
    
    def set_baudrate(self, baudrate: int):
        """Set baud rate."""
        self.baudrate = baudrate
        return self
    
    def set_timeout(self, timeout: float):
        """Set timeout."""
        self.timeout = timeout
        return self
    
    def build(self):
        """Build SerialDevice instance."""
        return SerialDevice(
            port=self.port,
            baudrate=self.baudrate,
            bytesize=self.bytesize,
            stopbits=self.stopbits,
            timeout=self.timeout
        )


class SerialDevice:
    """Serial communication device handler."""
    
    def __init__(
        self,
        port: str = "COM9",
        baudrate: int = 9600,
        bytesize: int = serial.EIGHTBITS,
        stopbits: int = serial.STOPBITS_ONE,
        timeout: float = 1
    ):
        """Initialize serial device."""
        self.ser: Optional[serial.Serial] = None
        
        try:
            logger.info(f"Opening serial port: {port}")
            self.ser = serial.Serial(
                port=port,
                baudrate=baudrate,
                bytesize=bytesize,
                stopbits=stopbits,
                timeout=timeout
            )
            logger.info(f"Serial port {port} opened successfully")
            print(f"Connected to {port}")
        except serial.SerialException as e:
            print(f"ERROR: Failed to open serial port: {e}")
            logger.error(f"SerialException: {e}", exc_info=True)
            print(f"\nTroubleshooting:")
            print(f"  - Check device is connected")
            print(f"  - Verify port '{port}' (use returnSerialPorts())")
            print(f"  - Ensure port is not in use")
            self.ser = None
            raise
        except ValueError as e:
            print(f"ERROR: Invalid parameters: {e}")
            logger.error(f"ValueError: {e}", exc_info=True)
            self.ser = None
            raise

    def query(self, command: bytes = b'?MPOW') -> Tuple[Optional[bytes], Optional[str], Optional[str]]:
        """
        Query device with command.
        
        Args:
            command: Byte command to send
            
        Returns:
            Tuple of (raw_response, hex_string, ascii_string)
        """
        if not self.ser or not self.ser.is_open:
            error_msg = "Serial port is not open"
            logger.error(error_msg)
            raise serial.SerialException(error_msg)
        
        try:
            logger.debug(f"Sending command: {command}")
            self.ser.write(command)
            
            response = self.ser.read(64)
            
            if not response:
                logger.warning("No response from device")
                return None, None, None
            
            # Decode response
            ascii_str = response.decode(errors='ignore').strip()
            
            # Convert to hex
            hex_str = response.hex()
            spaced_hex = ' '.join([hex_str[i:i+2] for i in range(0, len(hex_str), 2)])
            
            logger.debug(f"Received: {ascii_str} | Hex: {spaced_hex}")
            return response, spaced_hex, ascii_str
            
        except (serial.SerialTimeoutException, serial.SerialException, OSError) as e:
            logger.error(f"Communication error: {e}", exc_info=True)
            return None, None, None
        except Exception as e:
            logger.error(f"Unexpected error: {e}", exc_info=True)
            return None, None, None

    def close(self):
        """Close serial port."""
        if self.ser and self.ser.is_open:
            try:
                logger.info("Closing serial port")
                self.ser.close()
                logger.info("Serial port closed")
            except Exception as e:
                logger.error(f"Error closing port: {e}", exc_info=True)


if __name__ == "__main__":
    print("Serial Communication Module - Port Listing")
    print("="*50)
    returnSerialPorts()
