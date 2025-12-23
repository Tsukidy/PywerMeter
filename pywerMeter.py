from pywerHelper import serialComm, excelHelper
import time, logging, os, yaml

# Check config file for settings
configFilePath = "./config.yaml"
with open(configFilePath, 'r') as file:
    config = yaml.safe_load(file)


# Setup Logging
logPath = config['log_settings']['log_dir']
logName = "pywerMeter.log"
fullLogPath = os.path.join(logPath, logName)
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
if not os.path.exists(logPath):
    os.makedirs(logPath)
if not os.path.exists(fullLogPath):
    open(fullLogPath, 'a').close()
logging.basicConfig(filename=fullLogPath, encoding='utf-8', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')


def initSerialDevice():
    try:
        logger.info("Initializing serial device.")
        dev = serialComm.SerialDevice()
        return dev
    except Exception as e:
        print(f"Failed to initialize serial device: {e}")
        logger.error(f"Failed to initialize serial device: {e}")
        return None


def readSerialData(dev, command=b'?MPOW'):
    try:
        logger.info("Reading serial data.")
        unformattedData, retHexData, retAsciiData  = dev.query(command=command)
        logger.info(f"Recieved data: {retAsciiData}")
        return unformattedData, retHexData, retAsciiData
    except Exception as e:
        print(f"Error reading serial data: {e}")
        logger.error(f"Error reading serial data: {e}")
        return None, None, None


# Main execution block
if __name__ == "__main__":

    # Initialize serial device & variables
    dev = initSerialDevice()
    dataLog = ""
    minutes = 0.25
    end_time = time.time() + minutes * 60
    print("Attempting to read serial data.")

    # Read serial data for specified duration
    while time.time() < end_time:
        unformattedData, retHexData, retAsciiData  = readSerialData(dev, command=b'?MPOW')
        if retAsciiData:
            dataLog += retAsciiData + "\n"
            print(f"Received serial data: {retAsciiData}")
        #try:
        #    unformattedData, retHexData, retAsciiData  = dev.query(command=b'?MPOW')
        #    dataLog += retAsciiData + "\n"
        #    print(f"Received serial data: {retAsciiData}")
        #except Exception as e:
        #    print(f"Error during serial communication: {e}")
        #    logger.error(f"Error during serial communication: {e}")
    
    # Close serial device and write log to file
    try:
        logger.info("Closing serial device.")
        dev.close()
    except Exception as e:
        print(f"Error closing serial device: {e}")
        logger.error(f"Error closing serial device: {e}")

    # Write data log to file
    with open ("serialData.txt", "w") as file:
        try:
            file.write(dataLog)
        except Exception as e:
            print(f"Error writing to log file: {e}")
            logger.error(f"Error writing to log file: {e}")

    print("Serial data logging complete.")
    logger.info("Serial data logging complete.")
