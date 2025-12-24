# Python script for handling serial communication wtih a device.
import serial
import serial.tools.list_ports
import logging
import os
import yaml
import sys

# Start Logging with error handling
try:
    logPath = os.path.join(os.path.dirname(os.path.dirname(__file__)), "logs")
    logName = "serial_communication.log"
    fullLogPath = os.path.join(logPath, logName)
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
    
    # Create log directory if it doesn't exist & log file. Python will not do this itself.
    if not os.path.exists(logPath):
        try:
            os.makedirs(logPath)
        except PermissionError:
            print(f"ERROR: Permission denied creating log directory: {logPath}")
        except OSError as e:
            print(f"ERROR: Failed to create log directory: {e}")
    
    if not os.path.exists(fullLogPath):
        try:
            open(fullLogPath, 'a').close()
        except PermissionError:
            print(f"ERROR: Permission denied creating log file: {fullLogPath}")
        except OSError as e:
            print(f"ERROR: Failed to create log file: {e}")
    
    logging.basicConfig(filename=fullLogPath, encoding='utf-8', level=logging.DEBUG, 
                       format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    logger.info("Serial communication logger initialized")
except Exception as e:
    print(f"WARNING: Failed to setup logging for serial communication: {e}")
    # Create a null logger if file logging fails
    logger = logging.getLogger(__name__)
    logger.addHandler(logging.NullHandler())

# Return available serial ports
def returnSerialPorts():
    """List all available serial ports with error handling."""
    try:
        ports = serial.tools.list_ports.comports()
        if not ports:
            print("No serial ports found.")
            logger.warning("No serial ports detected on system")
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

# Primary class for serial communication
class SerialDevice:
    def __init__(
        self,
        port="COM9",
        baudrate=9600,
        bytesize=serial.EIGHTBITS,
        stopbits=serial.STOPBITS_ONE,
        timeout=1
    ):
        """Initialize serial device with comprehensive error handling."""
        # Load configuration from YAML if available
        config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config.yaml")
        
        try:
            if not os.path.exists(config_path):
                print("Config file not found. Using default serial settings.")
                logger.warning("Config file not found. Using default serial settings.")
            else:
                print("Config file found. Loading serial settings from config.")
                logger.info("Loading serial settings from config.yaml")
                
                try:
                    with open(config_path, 'r') as file:
                        config = yaml.safe_load(file)
                        
                    if config and 'connection_settings' in config:
                        port = config['connection_settings'].get('port', port)
                        baudrate = config['connection_settings'].get('baudrate', baudrate)
                        bytesize = config['connection_settings'].get('bytesize', bytesize)
                        stopbits = config['connection_settings'].get('stopbits', stopbits)
                        timeout = config['connection_settings'].get('timeout', timeout)
                        logger.info(f"Loaded settings: port={port}, baudrate={baudrate}, timeout={timeout}")
                    else:
                        logger.warning("Config file missing 'connection_settings' section")
                        
                except yaml.YAMLError as e:
                    print(f"ERROR: Invalid YAML in config file: {e}")
                    logger.error(f"YAML parsing error in config: {e}", exc_info=True)
                    print("Using default serial settings.")
                except PermissionError:
                    print(f"ERROR: Permission denied reading config file: {config_path}")
                    logger.error(f"Permission denied reading config: {config_path}")
                    print("Using default serial settings.")
        except Exception as e:
            print(f"ERROR: Unexpected error loading config: {e}")
            logger.error(f"Unexpected error loading config: {e}", exc_info=True)
            print("Using default serial settings.")
        
        # Attempt serial port connection
        self.ser = None
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
            logger.error(f"SerialException opening port {port}: {e}", exc_info=True)
            logger.error(f"Port: {port}, Baudrate: {baudrate}, Bytesize: {bytesize}, Stopbits: {stopbits}, Timeout: {timeout}")
            print(f"\nTroubleshooting:")
            print(f"  - Check that the device is connected")
            print(f"  - Verify port '{port}' is correct (use returnSerialPorts() to list available ports)")
            print(f"  - Ensure no other application is using the port")
            self.ser = None
            raise
        except ValueError as e:
            print(f"ERROR: Invalid serial port parameters: {e}")
            logger.error(f"ValueError with serial parameters: {e}", exc_info=True)
            self.ser = None
            raise
        except Exception as e:
            print(f"ERROR: Unexpected error opening serial port: {e}")
            logger.error(f"Unexpected error opening serial port: {e}", exc_info=True)
            self.ser = None
            raise

    def query(
        self,
        command=b'?MPOW'
    ):
        """Query device with comprehensive error handling."""
        if not self.ser or not self.ser.is_open:
            error_msg = "Serial port is not open"
            logger.error(error_msg)
            raise serial.SerialException(error_msg)
        
        try:
            logger.debug(f"Sending command: {command}")
            self.ser.write(command)
            
            logger.debug("Reading response from device")
            response = self.ser.read(64)
            
            if not response:
                logger.warning("No response received from device")
                return None, None, None
            
            # Decode response
            try:
                ascii_str = response.decode(errors='ignore')
                trimmed_ascii = ascii_str.strip()
            except UnicodeDecodeError as e:
                logger.warning(f"Unicode decode error: {e}. Using 'ignore' error handler.")
                ascii_str = response.decode(errors='ignore')
                trimmed_ascii = ascii_str.strip()
            
            # Convert to hex
            hex_str = response.hex()
            spaced_hex = ' '.join([hex_str[i:i+2] for i in range(0, len(hex_str), 2)])
            
            logger.debug(f"Received: {trimmed_ascii} | Hex: {spaced_hex}")
            return response, spaced_hex, trimmed_ascii
            
        except serial.SerialTimeoutException as e:
            print(f"ERROR: Serial timeout: {e}")
            logger.error(f"Serial timeout during query: {e}", exc_info=True)
            return None, None, None
        except serial.SerialException as e:
            print(f"ERROR: Serial communication error: {e}")
            logger.error(f"Serial exception during query: {e}", exc_info=True)
            return None, None, None
        except OSError as e:
            print(f"ERROR: OS error during communication: {e}")
            logger.error(f"OS error during serial query: {e}", exc_info=True)
            return None, None, None
        except Exception as e:
            print(f"ERROR: Unexpected error during query: {e}")
            logger.error(f"Unexpected error during query: {e}", exc_info=True)
            return None, None, None

    def close(self):
        """Close serial port with proper error handling."""
        if self.ser and self.ser.is_open:
            try:
                logger.info("Closing serial port")
                self.ser.close()
                logger.info("Serial port closed successfully")
            except serial.SerialException as e:
                print(f"ERROR: Serial exception while closing port: {e}")
                logger.error(f"Serial exception while closing port: {e}", exc_info=True)
            except Exception as e:
                print(f"ERROR: Unexpected error closing serial port: {e}")
                logger.error(f"Unexpected error closing serial port: {e}", exc_info=True)
        else:
            logger.debug("Serial port already closed or not initialized")

# Lists available serial ports when run as main
if __name__ == "__main__":
    print("Serial Communication Module - Port Listing")
    print("="*50)
    try:
        returnSerialPorts()
    except Exception as e:
        print(f"ERROR: Failed to list serial ports: {e}")
        logger.error(f"Failed to list serial ports in main: {e}", exc_info=True)
        sys.exit(1)
