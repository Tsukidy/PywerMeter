from pywerHelper import serialComm, excelHelper, dataCollector, menuHelper
from pywerHelper.configHelper import ConfigManager
from pywerHelper.timeUtils import parse_time_value, get_formatted_start_time, format_time_minutes
import time
import logging
import os
import sys
from datetime import datetime
from tkinter import filedialog
import tkinter as tk

# Store the script's directory for config loading
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Global configuration manager
config_manager: ConfigManager = None


def select_working_folder():
    """Prompt user to select a working folder using Windows file explorer."""
    print("Please select a working folder for pywerMeter...")
    
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    folder_selected = filedialog.askdirectory(
        title="Select Working Folder for pywerMeter",
        initialdir=os.getcwd()
    )
    
    root.destroy()
    
    if not folder_selected:
        print("No folder selected. Exiting...")
        sys.exit(0)
    
    os.chdir(folder_selected)
    print(f"Working folder set to: {os.path.abspath(folder_selected)}")
    print()
    
    return folder_selected


def load_config():
    """Load configuration from config.yaml in the script's directory."""
    global config_manager
    
    config_path = os.path.join(SCRIPT_DIR, "config.yaml")
    config_manager = ConfigManager(config_path)
    # Access config to trigger loading and validation
    _ = config_manager.config


def run_power_tests():
    """Execute the main power measurement test sequence."""
    test_settings = config_manager.get_test_settings()
    
    # Get default Excel filename
    folder_name = os.path.basename(os.getcwd())
    default_excel_file = f"{folder_name}.xlsx"
    logger.info(f"Using Excel filename: {default_excel_file}")
    
    # Check if Excel file exists
    tests_to_skip = set()
    if os.path.exists(default_excel_file):
        print(f"\n⚠️  Excel file '{default_excel_file}' already exists!")
        logger.warning(f"Excel file exists: {default_excel_file}")
        
        # Check completed tests
        try:
            completed_tests = excelHelper.get_completed_tests(default_excel_file, "Power Data")
            if completed_tests:
                print(f"Found {len(completed_tests)} completed test(s): {', '.join(completed_tests)}")
                logger.info(f"Completed tests: {completed_tests}")
        except Exception as e:
            print(f"Warning: Could not read existing file: {e}")
            logger.warning(f"Could not read existing file: {e}")
            completed_tests = []
        
        print("\nOptions:")
        print("[1] Continue from where testing left off")
        print("[2] Overwrite file and start fresh")
        print("[3] Create new file with timestamp")
        print("[x] Cancel")
        
        while True:
            response = input("\nSelect option: ").strip()
            
            if response == '1':
                if completed_tests:
                    tests_to_skip = set(completed_tests)
                    print(f"\nContinuing. Will skip: {', '.join(completed_tests)}\n")
                    logger.info(f"Skipping: {tests_to_skip}")
                else:
                    print("\nNo completed tests. Running all.\n")
                break
            elif response == '2':
                print("File will be overwritten.\n")
                logger.info("Overwriting file")
                try:
                    os.remove(default_excel_file)
                except PermissionError:
                    print(f"ERROR: Permission denied: {default_excel_file}")
                    return
                except OSError as e:
                    print(f"ERROR: Failed to delete: {e}")
                    return
                break
            elif response == '3':
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                base_name = os.path.splitext(default_excel_file)[0]
                ext = os.path.splitext(default_excel_file)[1]
                default_excel_file = f"{base_name}_{timestamp}{ext}"
                print(f"Using: {default_excel_file}\n")
                logger.info(f"New filename: {default_excel_file}")
                break
            elif response == 'x':
                print("Cancelled.")
                return
            else:
                print("Invalid option.")
    
    # Find all test configurations
    test_numbers = []
    for key in test_settings.keys():
        if key.startswith('test_excel_header_'):
            test_num = key.split('_')[-1]
            test_numbers.append(test_num)
    
    test_numbers.sort()
    
    # Initialize global timer
    global_start_time = time.time()
    elapsed_time = 0
    
    print("\n========== Starting Test Sequence ==========")
    print("Global timer started.")
    print(f"Data will be written to: {default_excel_file}\n")
    logger.info("Global timer started")
    
    # Initialize Excel structure
    if not os.path.exists(default_excel_file) or not tests_to_skip:
        test_headers = []
        for test_num in test_numbers:
            test_header = test_settings.get(f'test_excel_header_{test_num}')
            if test_header and test_header not in tests_to_skip:
                test_headers.append(test_header)
        
        if test_headers:
            print("Initializing Excel file...")
            logger.info(f"Creating headers: {test_headers}")
            if not excelHelper.initialize_excel_headers(test_headers, default_excel_file, "Power Data"):
                print("ERROR: Failed to initialize Excel file.")
                return
            print()
    
    # Run each test
    for test_num in test_numbers:
        test_header = test_settings.get(f'test_excel_header_{test_num}')
        start_time_raw = test_settings.get(f'test_start_time_{test_num}')
        duration_raw = test_settings.get(f'test_duration_{test_num}')
        pause_after = test_settings.get(f'after_test_pause_{test_num}')
        
        start_time = parse_time_value(start_time_raw)
        duration = parse_time_value(duration_raw)
        
        if test_header and start_time is not None and duration:
            if test_header in tests_to_skip:
                print(f"\n=== Skipping: {test_header} (completed) ===\n")
                continue
            
            # Wait for start time
            while elapsed_time < start_time:
                elapsed_time = (time.time() - global_start_time) / 60
                remaining = start_time - elapsed_time
                if remaining > 0:
                    print(f"\rGlobal: {format_time_minutes(elapsed_time)} | Waiting for {test_header} (starts {format_time_minutes(start_time)}, {format_time_minutes(remaining)} remaining)...", end="", flush=True)
                    time.sleep(1)
            
            elapsed_time = (time.time() - global_start_time) / 60
            
            print(f"\n\n=== Starting: {test_header} at {format_time_minutes(elapsed_time)} ===")
            logger.info(f"Starting {test_header} for {duration} min")
            
            start_time_str = get_formatted_start_time(elapsed_time)
            
            samples = []
            try:
                samples = dataCollector.serialFunction(logger, minutes=duration, global_timer_start=global_start_time, test_header=test_header)
                
                if samples:
                    print("Writing to Excel...")
                    excelHelper.write_test_row_to_excel(test_header, samples, default_excel_file, start_time_str=start_time_str)
                else:
                    print(f"No samples collected")
                    logger.warning(f"No samples: {test_header}")
                    
            except KeyboardInterrupt:
                # Still write whatever samples were collected before interruption
                if samples:
                    print("\nWriting collected samples to Excel before exit...")
                    logger.info(f"Writing {len(samples)} samples after interruption")
                    excelHelper.write_test_row_to_excel(test_header, samples, default_excel_file, start_time_str=start_time_str)
                    print(f"Wrote {len(samples)} samples to Excel.")
                print("\nTest sequence interrupted by user.")
                logger.info("Test sequence interrupted by user")
                return
            
            elapsed_time = (time.time() - global_start_time) / 60
            
            if pause_after:
                print(f"\nGlobal timer paused at {format_time_minutes(elapsed_time)}.")
                input("Press Enter to continue...")
                global_start_time = time.time() - (elapsed_time * 60)
                print("Timer resumed.\n")
            
            print(f"=== Complete: {test_header} ===\n")
    
    final_elapsed = (time.time() - global_start_time) / 60
    print(f"\n========== All Tests Complete ==========")
    print(f"Total time: {format_time_minutes(final_elapsed)}")
    print(f"Data: {default_excel_file}\n")
    logger.info(f"Complete. Time: {final_elapsed:.2f} min")


def rerun_specific_test():
    """Allow user to select and rerun specific tests."""
    test_settings = config_manager.get_test_settings()
    
    # Find available tests
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
        print("No tests found.")
        return
    
    test_numbers.sort()
    
    # Show menu
    print("\n=== Select Tests to Rerun ===")
    print("[1] Off")
    print("[2] Short Idle")
    print("[3] Short Idle, Long Idle")
    print("[4] Short Idle, Long Idle, Sleep")
    print("[5] Custom - Select specific tests")
    print("[x] Cancel")
    print("="*60)
    
    selected_test_nums = []
    while True:
        choice = input("\nSelect option: ").strip()
        
        if choice == 'x':
            print("Cancelled.")
            return
        
        if choice == '1':
            selected_test_nums = ['1']
            break
        elif choice == '2':
            selected_test_nums = ['2']
            break
        elif choice == '3':
            selected_test_nums = ['2', '3']
            break
        elif choice == '4':
            selected_test_nums = ['2', '3', '4']
            break
        elif choice == '5':
            print("\n=== Available Tests ===")
            for idx, test_num in enumerate(test_numbers, start=1):
                print(f"[{idx}] {test_mapping[test_num]['header']}")
            print("="*60)
            print("\nEnter test numbers (e.g., 1,3,4) or 'x' to cancel")
            
            custom_input = input("\nSelect tests: ").strip()
            
            if custom_input.lower() == 'x':
                print("Cancelled.")
                return
            
            try:
                selected_indices = [int(x.strip()) for x in custom_input.split(',')]
                
                if any(idx < 1 or idx > len(test_numbers) for idx in selected_indices):
                    print("Invalid selection.")
                    continue
                
                selected_test_nums = [test_numbers[idx - 1] for idx in selected_indices]
                selected_names = [test_mapping[num]['header'] for num in selected_test_nums]
                print(f"\nSelected: {', '.join(selected_names)}")
                break
            except ValueError:
                print("Invalid input.")
                continue
        else:
            print("Invalid option.")
    
    # Get filename
    folder_name = os.path.basename(os.getcwd())
    default_file = f"{folder_name}.xlsx"
    
    # Find all Excel files in current directory
    excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~$')]
    
    if excel_files:
        print("\n=== Available Excel Files ===")
        for idx, file in enumerate(excel_files, start=1):
            print(f"[{idx}] {file}")
        print("[n] Create new file (default name)")
        print("[x] Cancel")
        print("="*60)
        
        while True:
            file_choice = input("\nSelect file: ").strip()
            
            if file_choice.lower() == 'x':
                print("Cancelled.")
                return
            elif file_choice.lower() == 'n':
                filename = default_file
                print(f"\nUsing new file: {filename}")
                break
            else:
                try:
                    file_idx = int(file_choice)
                    if 1 <= file_idx <= len(excel_files):
                        filename = excel_files[file_idx - 1]
                        print(f"\nSelected: {filename}")
                        break
                    else:
                        print("Invalid selection.")
                except ValueError:
                    print("Invalid input. Enter a number, 'n', or 'x'.")
    else:
        print(f"\nNo Excel files found. Using default: {default_file}")
        filename = default_file
    
    # Replace or append mode - Always ask if file exists
    replace_mode = True
    file_exists = os.path.exists(filename)
    
    if file_exists:
        print(f"\n{'='*60}")
        print(f"⚠️  File '{filename}' exists.")
        print(f"{'='*60}")
        print("[1] Replace - Overwrite column data")
        print("[2] Append - Add below existing data")
        print("[x] Cancel")
        
        while True:
            mode_choice = input("\nSelect mode: ").strip()
            if mode_choice == '1':
                replace_mode = True
                print(f"\n⚠️  Will REPLACE data in '{filename}'")
                logger.info("Mode: Replace")
                break
            elif mode_choice == '2':
                replace_mode = False
                print(f"\n⚠️  Will APPEND data to '{filename}'")
                logger.info("Mode: Append")
                break
            elif mode_choice == 'x':
                print("Cancelled.")
                logger.info("User cancelled at mode selection")
                return
            else:
                print("Invalid option. Enter 1, 2, or x.")
    else:
        print(f"\n⚠️  File '{filename}' does not exist. Will create new file.")
        logger.info(f"File does not exist: {filename}")
    
    print("⚠️  Close Excel file before continuing!")
    input("\nPress Enter to continue...")
    
    # Initialize timer
    global_start_time = time.time()
    elapsed_time = 0
    
    # Time adjustment if "Off" excluded
    time_adjustment = 0
    adjust_enabled = test_settings.get('adjust_time_without_off', 'Off')
    
    if '1' not in selected_test_nums and isinstance(adjust_enabled, str) and adjust_enabled.lower() == 'on':
        off_duration = parse_time_value(test_settings.get('test_duration_1'))
        off_start = parse_time_value(test_settings.get('test_start_time_1', 0))
        time_adjustment = off_start + off_duration
        
        print(f"\n⚠️  'Off' test excluded. Adjusting times by -{format_time_minutes(time_adjustment)}")
        logger.info(f"Time adjustment: -{time_adjustment:.2f} min")
    
    print("\n========== Starting Test Sequence ==========")
    logger.info("Starting rerun sequence")
    
    # Run tests
    for test_num in selected_test_nums:
        test_header = test_settings.get(f'test_excel_header_{test_num}')
        start_time_raw = test_settings.get(f'test_start_time_{test_num}')
        duration_raw = test_settings.get(f'test_duration_{test_num}')
        pause_after = test_settings.get(f'after_test_pause_{test_num}')
        fast_start = test_settings.get(f'test_fast_start_{test_num}', 'Off')
        
        start_time = parse_time_value(start_time_raw)
        duration = parse_time_value(duration_raw)
        
        # Apply adjustments
        start_time = max(0, start_time - time_adjustment)
        
        if isinstance(fast_start, str) and fast_start.lower() == 'on':
            start_time = 0
            logger.info(f"FastStart enabled for {test_num}")
        
        if not duration:
            print(f"Error: No duration for {test_header}")
            continue
        
        # Wait for start time
        while elapsed_time < start_time:
            elapsed_time = (time.time() - global_start_time) / 60
            remaining = start_time - elapsed_time
            if remaining > 0:
                print(f"\rGlobal: {format_time_minutes(elapsed_time)} | Waiting for '{test_header}' (starts {format_time_minutes(start_time)}, {format_time_minutes(remaining)} remaining)...", end="", flush=True)
                time.sleep(1)
        
        elapsed_time = (time.time() - global_start_time) / 60
        
        print(f"\n\n=== Running: {test_header} ===")
        print(f"Starting: {format_time_minutes(elapsed_time)}")
        print(f"Duration: {format_time_minutes(duration)}")
        logger.info(f"Rerunning {test_header} for {duration} min")
        
        start_time_str = get_formatted_start_time(elapsed_time)
        
        samples = dataCollector.serialFunction(logger, minutes=duration, global_timer_start=global_start_time, test_header=test_header)
        
        if samples:
            mode_text = "replaced" if replace_mode else "appended"
            print("\nWriting to Excel...")
            if excelHelper.write_test_row_to_excel(test_header, samples, filename, start_time_str=start_time_str, replace_mode=replace_mode):
                print(f"✓ Data {mode_text}.")
            else:
                print("✗ Failed to write.")
        else:
            print("No samples collected")
        
        elapsed_time = (time.time() - global_start_time) / 60
        
        if pause_after:
            print(f"\nTimer paused at {format_time_minutes(elapsed_time)}.")
            input("Press Enter to continue...")
            global_start_time = time.time() - (elapsed_time * 60)
            print("Timer resumed.\n")
        
        print(f"\n=== Complete: {test_header} ===\n")
    
    final_elapsed = (time.time() - global_start_time) / 60
    print(f"\n========== All Tests Complete ==========")
    print(f"Total time: {format_time_minutes(final_elapsed)}\n")


def loggingSetup():
    """Setup logging."""
    try:
        log_settings = config_manager.get_log_settings()
        logPath = log_settings['log_dir']
        logName = "pywerMeter.log"
        fullLogPath = os.path.join(logPath, logName)
        
        logger = logging.getLogger(__name__)
        logger.setLevel(logging.DEBUG)
        
        if not os.path.exists(logPath):
            os.makedirs(logPath)
        
        if not os.path.exists(fullLogPath):
            open(fullLogPath, 'a').close()
        
        if not logger.handlers:
            file_handler = logging.FileHandler(fullLogPath, encoding='utf-8')
            file_handler.setLevel(logging.DEBUG)
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        
        logger.info("Logging initialized")
        return logger
    except Exception as e:
        print(f"ERROR: Logging setup failed: {e}")
        sys.exit(1)


# Main execution
if __name__ == "__main__":
    try:
        load_config()
        working_folder = select_working_folder()
        logger = loggingSetup()
        
        logger.info("="*60)
        logger.info("pywerMeter started")
        logger.info(f"Script: {SCRIPT_DIR}")
        logger.info(f"Working: {working_folder}")
        
        menuHelper.display_ascii_art()
        
        menu_options = {
            '1': 'Run Power Measurement Tests',
            '2': 'Add Power Calculations to Excel',
            '3': 'Rerun Specific Test',
            'x': 'Exit'
        }
        
        while True:
            try:
                choice = menuHelper.display_menu(menu_options)
                
                if choice == '1':
                    logger.info("User: Run Tests")
                    run_power_tests()
                    
                elif choice == '2':
                    logger.info("User: Add Calculations")
                    print("\n=== Add Power Calculations ===")
                    print("⚠️  Close Excel file before continuing!\n")
                    
                    folder_name = os.path.basename(os.getcwd())
                    default_file = f"{folder_name}.xlsx"
                    filename = input(f"Enter filename (Enter for '{default_file}'): ").strip()
                    if not filename:
                        filename = default_file
                    
                    if not os.path.exists(filename):
                        print(f"Error: File '{filename}' not found.")
                        continue
                    
                    print("\n[1] Add Averages Only")
                    print("[2] Add Total Annual Power Only")
                    print("[3] Add Both")
                    print("[x] Cancel")
                    
                    calc_choice = input("Select: ").strip()
                    
                    if calc_choice == 'x':
                        print("Cancelled.")
                        continue
                    
                    try:
                        calc = excelHelper.PowerCalc(filename, "Power Data")
                        
                        if calc_choice == '1':
                            print("\nAdding averages...")
                            if calc.add_averages():
                                print("✓ Averages added!")
                            else:
                                print("✗ Failed.")
                        elif calc_choice == '2':
                            print("\nAdding Total Annual Power...")
                            if calc.totalAnnualPower():
                                print("✓ Total Annual Power added!")
                            else:
                                print("✗ Failed.")
                        elif calc_choice == '3':
                            print("\nAdding averages...")
                            if calc.add_averages():
                                print("✓ Averages added!")
                                calc2 = excelHelper.PowerCalc(filename, "Power Data")
                                print("Adding Total Annual Power...")
                                if calc2.totalAnnualPower():
                                    print("✓ Total Annual Power added!")
                                else:
                                    print("✗ Failed.")
                            else:
                                print("✗ Failed.")
                        else:
                            print("Invalid option.")
                    except Exception as e:
                        print(f"ERROR: {e}")
                        logger.error(f"Calculation error: {e}", exc_info=True)
                    
                    print()
                    
                elif choice == '3':
                    logger.info("User: Rerun Test")
                    rerun_specific_test()
                    
                elif choice == 'x':
                    logger.info("User: Exit")
                    print("\nExiting. Goodbye!")
                    logger.info("Shutdown normally")
                    break
                    
            except KeyboardInterrupt:
                print("\n\nInterrupted (Ctrl+C)")
                logger.info("Menu interrupted")
                continue
    
    except KeyboardInterrupt:
        print("\n\nApplication interrupted. Exiting...")
        logger.info("Application interrupted")
        sys.exit(0)
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        logger.critical(f"Fatal error: {e}", exc_info=True)
        sys.exit(1)

