# Python script for handling serial communication wtih a device.
import serial, logging, os

# Start Logging
logName = "../logs/serial_communication.log"
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
if not os.path.exists("../logs"):
    os.makedirs("../logs")
if not os.path.exists("../logs/serial_communication.log"):
    open("../logs/serial_communication.log", 'a').close()
logging.basicConfig(filename=logName, encoding='utf-8', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

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

    def close(self):
        if self.ser and self.ser.is_open:
            try:
                logger.info("Closing serial port.")
                self.ser.close()
            except serial.SerialException as e:
                print(f"Error closing serial port: {e}")
                logger.error(f"Error closing serial port: {e}")

if __name__ == "__main__":
    returnSerialPorts()
