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

def serialFunction(minutes=0.25, filename="serialData.txt", global_timer_start=None):
    # Initialize serial device & variables
    dev = initSerialDevice()
    if dev is None:
        print("Failed to initialize serial device. Exiting.")
        logger.error("Failed to initialize serial device. Exiting.")
        exit(1)
    
    test_start_time = time.time()
    end_time = test_start_time + minutes * 60
    sample_count = 0
    recent_samples = []  # Keep track of last 15 samples
    print(f"Reading serial data for {minutes:.2f} minutes. Output: {filename}")

    # Open file for writing and write data as it comes in
    try:
        with open(filename, "w") as file:
            # Read serial data for specified duration
            while time.time() < end_time:
                unformattedData, retHexData, retAsciiData  = readSerialData(dev, command=b'?MPOW')
                if retAsciiData:
                    # Write data immediately to file
                    file.write(retAsciiData + "\n")
                    file.flush()  # Ensure data is written to disk immediately
                    sample_count += 1
                    # Add to recent samples and keep only last 15
                    recent_samples.append(retAsciiData)
                    if len(recent_samples) > 15:
                        recent_samples.pop(0)
                
                # Calculate time values
                current_time = time.time()
                test_elapsed = (current_time - test_start_time) / 60
                test_remaining = minutes - test_elapsed
                
                # Calculate global elapsed time if global timer was provided
                if global_timer_start:
                    global_elapsed = (current_time - global_timer_start) / 60
                    status = f"Test Progress: {test_elapsed:.2f}/{minutes:.2f} min | Remaining: {test_remaining:.2f} min | Global Timer: {global_elapsed:.2f} min | Samples: {sample_count}"
                else:
                    status = f"Test Progress: {test_elapsed:.2f}/{minutes:.2f} min | Remaining: {test_remaining:.2f} min | Samples: {sample_count}"
                
                # Clear previous lines (1 status line + up to 15 sample lines)
                print("\033[K", end="")  # Clear current line
                for _ in range(15):
                    print("\033[1B\033[K", end="")  # Move down and clear
                print(f"\033[16A\r{status}", end="")  # Move back up and print status
                
                # Print recent samples below status line
                for i, sample in enumerate(recent_samples):
                    print(f"\n  [{i+1:2d}] {sample}", end="")
                
                # Fill remaining lines if less than 15 samples
                for _ in range(15 - len(recent_samples)):
                    print("\n", end="")
                
                print("", flush=True)
        
        # Clear the sample lines and print newline after loop completes
        print("\033[K", end="")
        for _ in range(15):
            print("\033[1B\033[K", end="")
        print("\033[16A")
        print()  # Print newline after loop completes
        print(f"Data saved to {filename}")
        logger.info(f"Data saved to {filename}")
        
    except Exception as e:
        print(f"\nError writing to log file: {e}")
        logger.error(f"Error writing to log file: {e}")
    
    # Close serial device
    try:
        logger.info("Closing serial device.")
        dev.close()
    except Exception as e:
        print(f"Error closing serial device: {e}")
        logger.error(f"Error closing serial device: {e}")

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
            while elapsed_time < start_time:
                elapsed_time = (time.time() - global_start_time) / 60
                remaining = start_time - elapsed_time
                if remaining > 0:
                    print(f"\rGlobal Timer: {elapsed_time:.2f} min | Waiting for Test {test_num} (starts at {start_time} min, {remaining:.2f} min remaining)...", end="", flush=True)
                    time.sleep(1)
            
            print(f"\n\n=== Starting Test {test_num} at {elapsed_time:.2f} minutes ===")
            logger.info(f"Starting Test {test_num}: {filename} for {duration} minutes (started at {elapsed_time:.2f} min)")
            
            serialFunction(minutes=duration, filename=filename, global_timer_start=global_start_time)
            
            # Update elapsed time after test
            elapsed_time = (time.time() - global_start_time) / 60
            
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

