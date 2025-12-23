# Python script for handling serial communication wtih a device.
import serial, logging, os, yaml

# Start Logging
logPath = "../logs"
logName = "serial_communication.log"
fullLogPath = os.path.join(logPath, logName)
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
if not os.path.exists(logPath):
    os.makedirs(logPath)
if not os.path.exists(fullLogPath):
    open(fullLogPath, 'a').close()
logging.basicConfig(filename=fullLogPath, encoding='utf-8', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Return available serial ports
def returnSerialPorts():
    import serial.tools.list_ports
    ports = serial.tools.list_ports.comports()
    for port in ports:
        print(f"{port.device}: {port.description}")

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
        # Load configuration from YAML if available
        if not os.path.exists("../config.yaml"):
            print("Config file not found. Using default serial settings.")
            logger.warning("Config file not found. Using default serial settings.")
        else:
            print("Config file found. Loading serial settings from config.")
            logger.info("Config file found. Loading serial settings from config.")
            with open("../config.yaml", 'r') as file:
                config = yaml.safe_load(file)
                port = config['connection_settings'].get('port', port)
                baudrate = config['connection_settings'].get('baudrate', baudrate)
                bytesize = config['connection_settings'].get('bytesize', bytesize)
                stopbits = config['connection_settings'].get('stopbits', stopbits)
                timeout = config['connection_settings'].get('timeout', timeout)
        
        # Attempt serial port connection
        try:
            logger.info("Opening serial port.")
            self.ser = serial.Serial(
                port=port,
                baudrate=baudrate,
                bytesize=bytesize,
                stopbits=stopbits,
                timeout=timeout
            )
        except serial.SerialException as e:
            print(f"Error opening serial port: {e}")
            logger.error(f"Error opening serial port: {e}")
            logger.info("Serial Port" + port + " could not be opened.")
            logger.info("Current Baudrate: " + str(baudrate))
            logger.info("Current Bytesize: " + str(bytesize))
            logger.info("Current Stopbits: " + str(stopbits))
            logger.info("Current Timeout: " + str(timeout))
            self.ser = None
            raise

    # Query the device with a command and return response
    def query(
        self,
        command=b'?MPOW'
    ):
        if not self.ser or not self.ser.is_open:
            raise serial.SerialException("Serial port is not open.")
            logging.error("Serial port is not open.")
            return none, None, None
        try:
            logger.info(f"Sending command: {command}")
            self.ser.write(command)
            logger.info("Reading response from device.")
            response = self.ser.read(64)
            ascii_str = response.decode(errors='ignore')
            trimmed_ascii = ascii_str.strip()
            hex_str = response.hex()
            spaced_hex = ' '.join([hex_str[i:i+2] for i in range(0, len(hex_str), 2)])
            logger.info(f"Received response: {trimmed_ascii} | Hex: {spaced_hex}")
            logger.info("Query completed successfully.")
            return response, spaced_hex, trimmed_ascii
        except serial.SerialException as e:
            print(f"Error during serial communication: {e}")
            logger.error(f"Error during serial communication: {e}")
            return None, None, None

    # Close the serial port
    def close(self):
        if self.ser and self.ser.is_open:
            try:
                logger.info("Closing serial port.")
                self.ser.close()
            except serial.SerialException as e:
                print(f"Error closing serial port: {e}")
                logger.error(f"Error closing serial port: {e}")

# Lists available serial ports when run as main
if __name__ == "__main__":
    returnSerialPorts()
