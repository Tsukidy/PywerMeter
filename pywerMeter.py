from pywerHelper import serialComm, excelHelper
import time, logging, os, yaml

# Check config file for settings
configFilePath = "./config.yaml"
with open(configFilePath, 'r') as file:
    config = yaml.safe_load(file)

def loggingSetup():
    # Setup Logging
    logPath = config['log_settings']['log_dir']
    logName = "pywerMeter.log"
    fullLogPath = os.path.join(logPath, logName)
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
    # Create log directory if it doesn't exist & log file. Python will not do this itself.
    if not os.path.exists(logPath):
        os.makedirs(logPath)
    if not os.path.exists(fullLogPath):
        open(fullLogPath, 'a').close()
    logging.basicConfig(filename=fullLogPath, encoding='utf-8', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
    return logger


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

def serialFunction(minutes=0.25, filename="serialData.txt"):
    # Initialize serial device & variables
    dev = initSerialDevice()
    if dev is None:
        print("Failed to initialize serial device. Exiting.")
        logger.error("Failed to initialize serial device. Exiting.")
        exit(1)
    dataLog = ""
    end_time = time.time() + minutes * 60
    print(f"Attempting to read serial data for {minutes} minutes. Output: {filename}")

    # Read serial data for specified duration
    while time.time() < end_time:
        unformattedData, retHexData, retAsciiData  = readSerialData(dev, command=b'?MPOW')
        if retAsciiData:
            dataLog += retAsciiData + "\n"
            print(f"Received serial data: {retAsciiData}")
    
    # Close serial device and write log to file
    try:
        logger.info("Closing serial device.")
        dev.close()
    except Exception as e:
        print(f"Error closing serial device: {e}")
        logger.error(f"Error closing serial device: {e}")

    # Write data log to file
    with open (filename, "w") as file:
        try:
            file.write(dataLog)
            print(f"Data saved to {filename}")
            logger.info(f"Data saved to {filename}")
        except Exception as e:
            print(f"Error writing to log file: {e}")
            logger.error(f"Error writing to log file: {e}")

    print("Serial data logging complete.")
    logger.info("Serial data logging complete.")


# Main execution block
if __name__ == "__main__":

    # Setup logging
    logger = loggingSetup()
    logger.info("Starting pywerMeter serial data logging.")
    
    # Get test settings from config
    test_settings = config.get('test_settings', {})
    
    # Find all test configurations by looking for test_filename_x keys
    test_numbers = []
    for key in test_settings.keys():
        if key.startswith('test_filename_'):
            test_num = key.split('_')[-1]
            test_numbers.append(test_num)
    
    test_numbers.sort()
    
    # Run each test
    for test_num in test_numbers:
        filename = test_settings.get(f'test_filename_{test_num}')
        duration = test_settings.get(f'test_duration_{test_num}')
        pre_delay = test_settings.get(f'test_pre_delay_{test_num}', 0)
        pause_after = test_settings.get(f'after_test_pause_{test_num}')
        
        if filename and duration:
            print(f"\n=== Starting Test {test_num} ===")
            logger.info(f"Starting Test {test_num}: {filename} for {duration} minutes")
            
            # Apply pre-test delay if configured
            if pre_delay > 0:
                print(f"Pre-test delay: waiting {pre_delay} seconds...")
                logger.info(f"Pre-test delay: waiting {pre_delay} seconds")
                time.sleep(pre_delay)
            
            serialFunction(minutes=duration, filename=filename)
            if pause_after:
                input("Press Enter to continue to the next test...")
            print(f"=== Test {test_num} Complete ===\n")
        else:
            print(f"Warning: Missing configuration for test {test_num}")
            logger.warning(f"Missing configuration for test {test_num}")
    
    print("All tests complete.")
    logger.info("All tests complete.")

