from pywerHelper import serialComm, excelHelper
import time, logging, os, yaml

# Check config file for settings
configFilePath = "./config.yaml"
with open(configFilePath, 'r') as file:
    config = yaml.safe_load(file)

def display_ascii_art():
    """Display ASCII art for pywerMeter."""
    art = r"""
    ____                          __  ___      __           
   / __ \__  ___      _____  ____/  |/  /__  / /____  _____
  / /_/ / / / / | /| / / _ \/ __/ /|_/ / _ \/ __/ _ \/ ___/
 / ____/ /_/ /| |/ |/ /  __/ / / /  / /  __/ /_/  __/ /    
/_/    \__, / |__/|__/\___/_/ /_/  /_/\___/\__/\___/_/     
      /____/                                                
    """
    print(art)

def display_menu(menu_options):
    """
    Display a menu with numbered options.
    
    Args:
        menu_options (dict): Dictionary with keys as option numbers and values as option descriptions
        
    Returns:
        str: The selected option key
    """
    print("\n" + "="*60)
    for key, description in menu_options.items():
        print(f"[{key}] {description}")
    print("="*60)
    
    while True:
        choice = input("\nSelect an option: ").strip()
        if choice in menu_options:
            return choice
        else:
            print(f"Invalid option. Please select from {list(menu_options.keys())}")

def run_power_tests():
    """Execute the main power measurement test sequence."""
    # Get test settings from config
    test_settings = config.get('test_settings', {})
    
    # Get default Excel filename
    default_excel_file = test_settings.get('default_excel_file', 'power_measurements.xlsx')
    
    # Check if Excel file already exists
    if os.path.exists(default_excel_file):
        print(f"\n⚠️  Warning: Excel file '{default_excel_file}' already exists!")
        logger.warning(f"Excel file already exists: {default_excel_file}")
        
        while True:
            response = input("Do you want to overwrite it? (y/n): ").strip().lower()
            if response == 'y':
                print(f"File will be overwritten.\n")
                logger.info("User chose to overwrite existing Excel file")
                # Delete the existing file
                try:
                    os.remove(default_excel_file)
                    logger.info(f"Deleted existing file: {default_excel_file}")
                except Exception as e:
                    print(f"Error deleting file: {e}")
                    logger.error(f"Error deleting existing file: {e}")
                    exit(1)
                break
            elif response == 'n':
                # Generate new filename with timestamp
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                base_name = os.path.splitext(default_excel_file)[0]
                ext = os.path.splitext(default_excel_file)[1]
                default_excel_file = f"{base_name}_{timestamp}{ext}"
                print(f"Using new filename: {default_excel_file}\n")
                logger.info(f"User chose not to overwrite. Using new filename: {default_excel_file}")
                break
            else:
                print("Invalid input. Please enter 'y' or 'n'.")
    
    # Find all test configurations by looking for test_excel_header_x keys
    test_numbers = []
    for key in test_settings.keys():
        if key.startswith('test_excel_header_'):
            test_num = key.split('_')[-1]
            test_numbers.append(test_num)
    
    test_numbers.sort()
    
    # Initialize global timer
    global_start_time = time.time()
    elapsed_time = 0  # in minutes
    
    print("\n========== Starting Test Sequence ==========")
    print("Global timer started. Tests will run based on scheduled start times.")
    print(f"Data will be written to: {default_excel_file}\n")
    logger.info("Global timer started for test sequence")
    logger.info(f"Excel output file: {default_excel_file}")
    
    # Run each test
    for test_num in test_numbers:
        test_header = test_settings.get(f'test_excel_header_{test_num}')
        start_time_raw = test_settings.get(f'test_start_time_{test_num}')
        duration_raw = test_settings.get(f'test_duration_{test_num}')
        pause_after = test_settings.get(f'after_test_pause_{test_num}')
        
        # Parse time values
        start_time = parse_time_value(start_time_raw) if start_time_raw is not None else None
        duration = parse_time_value(duration_raw) if duration_raw is not None else None
        
        if test_header and start_time is not None and duration:
            # Wait until the global timer reaches the start time
            while elapsed_time < start_time:
                elapsed_time = (time.time() - global_start_time) / 60
                remaining = start_time - elapsed_time
                if remaining > 0:
                    print(f"\rGlobal Timer: {elapsed_time:.2f} min | Waiting for Test {test_num} (starts at {start_time} min, {remaining:.2f} min remaining)...", end="", flush=True)
                    time.sleep(1)
            
            print(f"\n\n=== Starting Test {test_num}: {test_header} at {elapsed_time:.2f} minutes ===")
            logger.info(f"Starting Test {test_num}: {test_header} for {duration} minutes (started at {elapsed_time:.2f} min)")
            
            # Run test and collect samples
            samples = serialFunction(minutes=duration, global_timer_start=global_start_time, test_header=test_header)
            
            # Write samples to Excel immediately after test completes
            if samples:
                print(f"Writing test data to Excel...")
                logger.info(f"Writing {len(samples)} samples to Excel for test: {test_header}")
                excelHelper.write_test_row_to_excel(test_header, samples, default_excel_file)
            else:
                print(f"No samples collected for {test_header}")
                logger.warning(f"No samples collected for test: {test_header}")
            
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

def serialFunction(minutes=0.25, global_timer_start=None, test_header=None):
    # Initialize serial device & variables
    dev = initSerialDevice()
    if dev is None:
        print("Failed to initialize serial device. Exiting.")
        logger.error("Failed to initialize serial device. Exiting.")
        exit(1)
    
    test_start_time = time.time()
    end_time = test_start_time + minutes * 60
    sample_count = 0
    samples = []  # Store all samples for this test
    recent_samples = []  # Keep track of last 15 samples for display
    print(f"Reading serial data for {minutes:.2f} minutes. Test: {test_header}")

    # Read serial data for specified duration
    try:
        while time.time() < end_time:
            unformattedData, retHexData, retAsciiData  = readSerialData(dev, command=b'?MPOW')
            if retAsciiData:
                # Store sample
                samples.append(retAsciiData)
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
            
            # Use Windows-compatible method to update display
            # Move cursor to beginning and clear with spaces
            lines_to_clear = 16  # 1 status + 15 samples
            
            # Move cursor up if not first iteration
            if sample_count > 1:
                for _ in range(lines_to_clear):
                    print(f"\033[F", end="")  # Move cursor up one line
            
            # Print status line
            print(f"\r{status:<120}")  # Left-align and pad to 120 chars to clear previous text
            
            # Print recent samples
            for i in range(15):
                if i < len(recent_samples):
                    print(f"  [{i+1:2d}] {recent_samples[i]:<100}")  # Pad to clear previous text
                else:
                    print(f"{' ':<120}")  # Empty line padded with spaces
            
            # Flush output
            print(end="", flush=True)
        
        # Move past the display area
        print("\n")
        print(f"Test complete: {sample_count} samples collected")
        logger.info(f"Test complete: {sample_count} samples collected")
        
    except Exception as e:
        print(f"\nError during data collection: {e}")
        logger.error(f"Error during data collection: {e}")
    
    # Close serial device
    try:
        logger.info("Closing serial device.")
        dev.close()
    except Exception as e:
        print(f"Error closing serial device: {e}")
        logger.error(f"Error closing serial device: {e}")

    print("Serial data logging complete.")
    logger.info("Serial data logging complete.")
    
    return samples  # Return collected samples


# Main execution block
if __name__ == "__main__":

    # Setup logging
    logger = loggingSetup()
    logger.info("Starting pywerMeter application.")
    
    # Display ASCII art
    display_ascii_art()
    
    # Define menu options (easily expandable - just add new entries here)
    menu_options = {
        '1': 'Run Power Measurement Tests',
        '2': 'Add Power Calculations to Existing Excel File',
        '3': 'Exit'
    }
    
    # Main menu loop
    while True:
        choice = display_menu(menu_options)
        
        if choice == '1':
            # Run power measurement tests
            logger.info("User selected: Run Power Measurement Tests")
            run_power_tests()
            
        elif choice == '2':
            # Add power calculations to existing Excel file
            logger.info("User selected: Add Power Calculations")
            print("\n=== Add Power Calculations ===")
            
            # Get filename from user
            default_file = config.get('test_settings', {}).get('default_excel_file', 'power_measurements.xlsx')
            filename = input(f"Enter Excel filename (press Enter for '{default_file}'): ").strip()
            if not filename:
                filename = default_file
            
            if not os.path.exists(filename):
                print(f"Error: File '{filename}' not found.")
                logger.error(f"Excel file not found: {filename}")
                continue
            
            # Ask which calculations to perform
            print("\nSelect calculation type:")
            print("[1] Add Averages Only")
            print("[2] Add Total Annual Power Only")
            print("[3] Add Both (Averages + Total Annual Power)")
            
            calc_choice = input("Select option: ").strip()
            
            calc = excelHelper.PowerCalc(filename, "Power Data")
            
            if calc_choice == '1':
                print("\nAdding averages...")
                if calc.add_averages():
                    print("✓ Averages added successfully!")
                else:
                    print("✗ Failed to add averages.")
                    
            elif calc_choice == '2':
                print("\nAdding Total Annual Power...")
                if calc.totalAnnualPower():
                    print("✓ Total Annual Power added successfully!")
                else:
                    print("✗ Failed to add Total Annual Power.")
                    
            elif calc_choice == '3':
                print("\nAdding averages...")
                if calc.add_averages():
                    print("✓ Averages added successfully!")
                    # Reload for totalAnnualPower
                    calc2 = excelHelper.PowerCalc(filename, "Power Data")
                    print("Adding Total Annual Power...")
                    if calc2.totalAnnualPower():
                        print("✓ Total Annual Power added successfully!")
                    else:
                        print("✗ Failed to add Total Annual Power.")
                else:
                    print("✗ Failed to add averages.")
            else:
                print("Invalid option.")
            
            print()
            
        elif choice == '3':
            # Exit
            logger.info("User selected: Exit")
            print("\nExiting pywerMeter. Goodbye!")
            break
        
        else:
            # This shouldn't happen due to validation in display_menu, but just in case
            print("Invalid option.")
            logger.warning(f"Invalid menu selection: {choice}")

