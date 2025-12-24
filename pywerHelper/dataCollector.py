# This module handles serial data collection for the pywerHelper package.
# It includes functions to initialize serial devices, read data, and perform data collection tests.
# Author: Dylan Pope
# Date: 2024-12-24
# Version: 1.0.0

import time
from . import serialComm


def initSerialDevice(logger):
    """
    Initialize serial device with comprehensive error handling.
    
    Args:
        logger: Logger instance for logging operations
        
    Returns:
        SerialDevice: Initialized serial device object, or None if failed
    """
    try:
        logger.info("Initializing serial device")
        dev = serialComm.SerialDevice()
        logger.info("Serial device initialized successfully")
        return dev
    except serialComm.serial.SerialException as e:
        print(f"ERROR: Serial port error: {e}")
        print("Check that the device is connected and the COM port is correct in config.yaml")
        logger.error(f"Serial port exception during initialization: {e}", exc_info=True)
        return None
    except FileNotFoundError as e:
        print(f"ERROR: Configuration file not found: {e}")
        logger.error(f"Configuration file not found during serial init: {e}", exc_info=True)
        return None
    except Exception as e:
        print(f"ERROR: Unexpected error initializing serial device: {e}")
        logger.error(f"Unexpected error initializing serial device: {e}", exc_info=True)
        return None


def readSerialData(dev, logger, command=b'?MPOW'):
    """
    Read data from serial device with proper error handling.
    
    Args:
        dev: SerialDevice object to query
        logger: Logger instance for logging operations
        command: Byte string command to send to device
        
    Returns:
        tuple: (unformattedData, retHexData, retAsciiData) or (None, None, None) on error
    """
    try:
        logger.debug(f"Querying device with command: {command}")
        unformattedData, retHexData, retAsciiData = dev.query(command=command)
        logger.debug(f"Received data: {retAsciiData}")
        return unformattedData, retHexData, retAsciiData
    except serialComm.serial.SerialException as e:
        print(f"ERROR: Serial communication error: {e}")
        logger.error(f"Serial communication error: {e}", exc_info=True)
        return None, None, None
    except AttributeError as e:
        print(f"ERROR: Invalid device object: {e}")
        logger.error(f"Invalid device object in readSerialData: {e}", exc_info=True)
        return None, None, None
    except Exception as e:
        print(f"ERROR: Unexpected error reading serial data: {e}")
        logger.error(f"Unexpected error reading serial data: {e}", exc_info=True)
        return None, None, None


def serialFunction(logger, minutes=0.25, global_timer_start=None, test_header=None):
    """
    Execute serial data collection with comprehensive error handling.
    
    Args:
        logger: Logger instance for logging operations
        minutes: Duration of data collection in minutes
        global_timer_start: Optional global timer start time for display
        test_header: Name of the test being run
        
    Returns:
        list: List of collected sample strings
    """
    # Initialize serial device & variables
    dev = initSerialDevice(logger)
    if dev is None:
        print("ERROR: Failed to initialize serial device. Cannot proceed with test.")
        logger.error("Failed to initialize serial device. Test aborted.")
        return []
    
    test_start_time = time.time()
    end_time = test_start_time + minutes * 60
    sample_count = 0
    samples = []  # Store all samples for this test
    recent_samples = []  # Keep track of last 15 samples for display
    
    print(f"Reading serial data for {minutes:.2f} minutes. Test: {test_header}")
    logger.info(f"Starting data collection: {minutes:.2f} minutes for test '{test_header}'")

    # Read serial data for specified duration
    try:
        while time.time() < end_time:
            unformattedData, retHexData, retAsciiData = readSerialData(dev, logger, command=b'?MPOW')
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
        logger.info(f"Test complete: {sample_count} samples collected for '{test_header}'")
        
    except KeyboardInterrupt:
        print("\n\nTest interrupted by user (Ctrl+C)")
        logger.warning(f"Test '{test_header}' interrupted by user. Collected {sample_count} samples before interruption.")
        print(f"Collected {sample_count} samples before interruption.")
    except Exception as e:
        print(f"\nERROR: Unexpected error during data collection: {e}")
        logger.error(f"Unexpected error during data collection for '{test_header}': {e}", exc_info=True)
    finally:
        # Always close serial device
        try:
            if dev:
                logger.info("Closing serial device")
                dev.close()
                logger.info("Serial device closed successfully")
        except Exception as e:
            print(f"ERROR: Error closing serial device: {e}")
            logger.error(f"Error closing serial device: {e}", exc_info=True)

    logger.info(f"Serial data logging complete. Returning {len(samples)} samples")
    return samples  # Return collected samples


# Module-level exports
__all__ = ['initSerialDevice', 'readSerialData', 'serialFunction']
