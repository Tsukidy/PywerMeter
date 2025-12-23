from pywerHelper import serialComm, excelHelper
import time, logging, os, yaml

# Check config file for settings
configFilePath = "./config.yaml"
with open(configFilePath, 'r') as file:
    config = yaml.safe_load(file)

def parse_time_value(time_value):
    """
    Parse time value from config. Accepts:
    - Numeric (int/float): treated as minutes (e.g., 1.5 = 1.5 minutes)
    - String "M:SS": parsed as minutes:seconds (e.g., "1:30" = 1.5 minutes)
    Returns time in minutes as a float.
    """
    if isinstance(time_value, (int, float)):
        return float(time_value)
    elif isinstance(time_value, str) and ':' in time_value:
        parts = time_value.split(':')
        if len(parts) == 2:
            try:
                minutes = int(parts[0])
                seconds = int(parts[1])
                return minutes + (seconds / 60.0)
            except ValueError:
                print(f"Warning: Invalid time format '{time_value}'. Using 0.")
                return 0.0
    print(f"Warning: Unrecognized time format '{time_value}'. Using 0.")
    return 0.0

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
    
    # Initialize global timer
    global_start_time = time.time()
    elapsed_time = 0  # in minutes
    
    print("\n========== Starting Test Sequence ==========")
    print("Global timer started. Tests will run based on scheduled start times.\n")
    logger.info("Global timer started for test sequence")
    
    # Run each test
    for test_num in test_numbers:
        filename = test_settings.get(f'test_filename_{test_num}')
        start_time_raw = test_settings.get(f'test_start_time_{test_num}')
        duration_raw = test_settings.get(f'test_duration_{test_num}')
        pause_after = test_settings.get(f'after_test_pause_{test_num}')
        
        # Parse time values
        start_time = parse_time_value(start_time_raw) if start_time_raw is not None else None
        duration = parse_time_value(duration_raw) if duration_raw is not None else None
        
        if filename and start_time is not None and duration:
            # Wait until the global timer reaches the start time
            elapsed_time = (time.time() - global_start_time) / 60
            while elapsed_time < start_time:
                elapsed_time = (time.time() - global_start_time) / 60
                remaining = start_time - elapsed_time
                if remaining > 0:
                    print(f"\rGlobal Timer: {elapsed_time:.2f} min | Waiting for Test {test_num} (starts at {start_time} min, {remaining:.2f} min remaining)...", end="", flush=True)
                    time.sleep(1)
            
            # Update elapsed time one more time before starting test
            elapsed_time = (time.time() - global_start_time) / 60
            print(f"\n\n=== Starting Test {test_num} at {elapsed_time:.2f} minutes ===")
            logger.info(f"Starting Test {test_num}: {filename} for {duration} minutes (started at {elapsed_time:.2f} min)")
            
            serialFunction(minutes=duration, filename=filename)
            
            # Update elapsed time after test to include test duration
            elapsed_time = (time.time() - global_start_time) / 60
            print(f"Test completed. Global timer now at: {elapsed_time:.2f} minutes")
            
            # Pause the timer if requested
            if pause_after:
                print(f"\nGlobal timer paused at {elapsed_time:.2f} minutes.")
                logger.info(f"Global timer paused at {elapsed_time:.2f} minutes for user input")
                input("Press Enter to continue to the next test...")
                # Adjust the global start time to account for the pause
                global_start_time = time.time() - (elapsed_time * 60)
                print(f"Global timer resumed.\n")
                logger.info("Global timer resumed")
            
            print(f"=== Test {test_num} Complete ===\n")
        else:
            print(f"Warning: Missing configuration for test {test_num}")
            logger.warning(f"Missing configuration for test {test_num}")
    
    final_elapsed = (time.time() - global_start_time) / 60
    print(f"\n========== All Tests Complete ==========")
    print(f"Total elapsed time: {final_elapsed:.2f} minutes\n")
    logger.info(f"All tests complete. Total elapsed time: {final_elapsed:.2f} minutes")

