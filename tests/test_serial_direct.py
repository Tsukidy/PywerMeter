"""Quick test to verify serial device is responding"""
import serial
import time
import sys
import os

# Add parent directory to path to import from pywerHelper
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pywerHelper.configHelper import ConfigManager

# Try to open the port and query
try:
    # Load settings from config.yaml
    config = ConfigManager()
    settings = config.get_serial_settings()
    
    print("Loading serial settings from config.yaml:")
    print(f"  Port: {settings.get('port')}")
    print(f"  Baudrate: {settings.get('baudrate')}")
    print(f"  Bytesize: {settings.get('bytesize')}")
    print(f"  Stopbits: {settings.get('stopbits')}")
    print(f"  Timeout: {settings.get('timeout')}")
    print()
    
    ser = serial.Serial(
        port=settings.get('port', 'COM9'),
        baudrate=settings.get('baudrate', 9600),
        bytesize=settings.get('bytesize', serial.EIGHTBITS),
        stopbits=settings.get('stopbits', serial.STOPBITS_ONE),
        timeout=settings.get('timeout', 1)
    )
    
    print(f"Opened port: {ser.port}")
    print(f"Baudrate: {ser.baudrate}")
    print(f"Is open: {ser.is_open}")
    
    # Try different command formats
    commands = [
        (b'?MPOW', "No terminator"),
        (b'?MPOW\r', "With \\r"),
        (b'?MPOW\n', "With \\n"),
        (b'?MPOW\r\n', "With \\r\\n"),
    ]
    
    for cmd, desc in commands:
        print(f"\n{'='*50}")
        print(f"Testing: {desc}")
        print(f"Command: {cmd}")
        
        # Clear any pending data
        ser.reset_input_buffer()
        
        # Send query
        ser.write(cmd)
        
        # Read response
        print("Waiting for response...")
        response = ser.read(64)
        
        print(f"Response length: {len(response)} bytes")
        print(f"Response (raw): {response}")
        print(f"Response (hex): {response.hex()}")
        
        if response:
            ascii_str = response.decode(errors='ignore').strip()
            print(f"Response (ASCII): '{ascii_str}'")
            print(f"✓ SUCCESS!")
            break
        else:
            print("✗ No response")
        
        time.sleep(0.1)
    
    ser.close()
    print("\n" + "="*50)
    print("Port closed")
    
    ser.close()
    print("\nPort closed")
    
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
