from pywerHelper import serialComm, excelHelper, dataCollector, menuHelper
import time, logging, os, yaml
import sys

# Check config file for settings
configFilePath = "./config.yaml"
try:
    with open(configFilePath, 'r') as file:
        config = yaml.safe_load(file)
except FileNotFoundError:
    print(f"ERROR: Configuration file not found: {configFilePath}")
    print("Please ensure config.yaml exists in the application directory.")
    sys.exit(1)
except yaml.YAMLError as e:
    print(f"ERROR: Invalid YAML syntax in configuration file: {e}")
    sys.exit(1)
except PermissionError:
    print(f"ERROR: Permission denied reading configuration file: {configFilePath}")
    sys.exit(1)
except Exception as e:
    print(f"ERROR: Unexpected error loading configuration: {e}")
    sys.exit(1)

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
                except PermissionError:
                    print(f"ERROR: Permission denied. File may be open in another program: {default_excel_file}")
                    logger.error(f"Permission denied deleting file: {default_excel_file}", exc_info=True)
                    return
                except OSError as e:
                    print(f"ERROR: Failed to delete file: {e}")
                    logger.error(f"OS error deleting file: {default_excel_file}", exc_info=True)
                    return
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
            samples = dataCollector.serialFunction(logger, minutes=duration, global_timer_start=global_start_time, test_header=test_header)
            
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

def rerun_specific_test():
    """Allow user to select and rerun a specific test, overwriting its data in Excel."""
    test_settings = config.get('test_settings', {})
    
    # Find all available tests
    test_mapping = {}
    test_numbers = []
    
    for key in test_settings.keys():
        if key.startswith('test_excel_header_'):
            test_num = key.split('_')[-1]
            test_header = test_settings.get(f'test_excel_header_{test_num}')
            if test_header:
                test_numbers.append(test_num)
                test_mapping[test_num] = {'header': test_header}
    
    if not test_mapping:
        print("No tests found in configuration.")
        logger.warning("No tests found in configuration for rerun")
        return
    
    # Sort test numbers
    test_numbers.sort()
    
    # Display available tests
    print("\n=== Available Tests to Rerun ===")
    for idx, test_num in enumerate(test_numbers, start=1):
        test_header = test_mapping[test_num]['header']
        print(f"[{idx}] {test_header}")
    print(f"[{len(test_numbers) + 1}] Cancel")
    print("="*60)
    
    # Get user selection
    while True:
        try:
            choice = input("\nSelect test to rerun: ").strip()
            choice_idx = int(choice)
            
            if choice_idx == len(test_numbers) + 1:
                print("Cancelled.")
                return
            
            if 1 <= choice_idx <= len(test_numbers):
                selected_num = test_numbers[choice_idx - 1]
                break
            else:
                print(f"Invalid option. Please select 1-{len(test_numbers) + 1}.")
        except ValueError:
            print("Invalid input. Please enter a number.")
    
    # Get test configuration
    test_header = test_settings.get(f'test_excel_header_{selected_num}')
    duration_raw = test_settings.get(f'test_duration_{selected_num}')
    
    duration = parse_time_value(duration_raw) if duration_raw is not None else None
    
    if not duration:
        print(f"Error: No duration configured for {test_header}")
        logger.error(f"No duration configured for test {selected_num}")
        return
    
    # Get Excel filename
    default_file = test_settings.get('default_excel_file', 'power_measurements.xlsx')
    filename = input(f"\nEnter Excel filename (press Enter for '{default_file}'): ").strip()
    if not filename:
        filename = default_file
    
    # Confirm before overwriting
    if os.path.exists(filename):
        print(f"\n⚠️  This will overwrite the existing '{test_header}' column in '{filename}'")
        confirm = input("Continue? (y/n): ").strip().lower()
        if confirm != 'y':
            print("Cancelled.")
            return
    
    # Run the test
    print(f"\n=== Running Test: {test_header} ===")
    print(f"Duration: {duration:.2f} minutes")
    logger.info(f"Rerunning test: {test_header} for {duration} minutes")
    
    samples = dataCollector.serialFunction(logger, minutes=duration, global_timer_start=None, test_header=test_header)
    
    # Write to Excel (will overwrite existing column)
    if samples:
        print(f"\nWriting test data to Excel...")
        logger.info(f"Writing {len(samples)} samples to Excel for test: {test_header}")
        if excelHelper.write_test_row_to_excel(test_header, samples, filename):
            print(f"✓ Test data written successfully! Column '{test_header}' updated.")
        else:
            print(f"✗ Failed to write test data.")
    else:
        print(f"No samples collected for {test_header}")
        logger.warning(f"No samples collected for rerun of test: {test_header}")
    
    print(f"\n=== Test Complete ===\n")

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
    """Setup logging with proper error handling."""
    try:
        # Setup Logging
        logPath = config['log_settings']['log_dir']
        logName = "pywerMeter.log"
        fullLogPath = os.path.join(logPath, logName)
        logger = logging.getLogger(__name__)
        logger.setLevel(logging.DEBUG)
        
        # Create log directory if it doesn't exist & log file. Python will not do this itself.
        if not os.path.exists(logPath):
            try:
                os.makedirs(logPath)
            except PermissionError:
                print(f"ERROR: Permission denied creating log directory: {logPath}")
                sys.exit(1)
            except OSError as e:
                print(f"ERROR: Failed to create log directory: {e}")
                sys.exit(1)
        
        if not os.path.exists(fullLogPath):
            try:
                open(fullLogPath, 'a').close()
            except PermissionError:
                print(f"ERROR: Permission denied creating log file: {fullLogPath}")
                sys.exit(1)
            except OSError as e:
                print(f"ERROR: Failed to create log file: {e}")
                sys.exit(1)
        
        logging.basicConfig(filename=fullLogPath, encoding='utf-8', level=logging.DEBUG, 
                          format='%(asctime)s - %(levelname)s - %(message)s')
        logger.info("Logging system initialized successfully")
        return logger
    except KeyError as e:
        print(f"ERROR: Missing configuration key: {e}")
        print("Please check your config.yaml file.")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: Failed to setup logging: {e}")
        sys.exit(1)


# Main execution block
if __name__ == "__main__":
    try:
        # Setup logging
        logger = loggingSetup()
        logger.info("="*60)
        logger.info("Starting pywerMeter application")
        
        # Display ASCII art
        menuHelper.display_ascii_art()
        
        # Define menu options (easily expandable - just add new entries here)
        menu_options = {
            '1': 'Run Power Measurement Tests',
            '2': 'Add Power Calculations to Existing Excel File',
            '3': 'Rerun Specific Test',
            'x': 'Exit'
        }
        
        # Main menu loop
        while True:
            try:
                choice = menuHelper.display_menu(menu_options)
                
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
                    
                    try:
                        calc = excelHelper.PowerCalc(filename, "Power Data")
                        
                        if calc_choice == '1':
                            print("\nAdding averages...")
                            logger.info(f"Adding averages to {filename}")
                            if calc.add_averages():
                                print("✓ Averages added successfully!")
                            else:
                                print("✗ Failed to add averages.")
                                
                        elif calc_choice == '2':
                            print("\nAdding Total Annual Power...")
                            logger.info(f"Adding Total Annual Power to {filename}")
                            if calc.totalAnnualPower():
                                print("✓ Total Annual Power added successfully!")
                            else:
                                print("✗ Failed to add Total Annual Power.")
                                
                        elif calc_choice == '3':
                            print("\nAdding averages...")
                            logger.info(f"Adding averages and Total Annual Power to {filename}")
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
                            logger.warning(f"Invalid calculation option selected: {calc_choice}")
                    except Exception as e:
                        print(f"ERROR: Failed to perform calculations: {e}")
                        logger.error(f"Failed to perform Excel calculations: {e}", exc_info=True)
                    
                    print()
                    
                elif choice == '3':
                    # Rerun specific test
                    logger.info("User selected: Rerun Specific Test")
                    rerun_specific_test()
                    
                elif choice == 'x':
                    # Exit
                    logger.info("User selected: Exit")
                    print("\nExiting pywerMeter. Goodbye!")
                    logger.info("Application shutdown normally")
                    break
                
                else:
                    # This shouldn't happen due to validation in display_menu, but just in case
                    print("Invalid option.")
                    logger.warning(f"Invalid menu selection: {choice}")
                    
            except KeyboardInterrupt:
                print("\n\nMenu interrupted by user (Ctrl+C)")
                logger.info("Menu loop interrupted by user")
                continue
    
    except KeyboardInterrupt:
        print("\n\nApplication interrupted by user (Ctrl+C). Exiting...")
        logger.info("Application interrupted by user. Shutting down.")
        sys.exit(0)
    except Exception as e:
        print(f"\nFATAL ERROR: Unexpected error in main execution: {e}")
        logger.critical(f"Fatal error in main execution: {e}", exc_info=True)
        sys.exit(1)

