"""Quick test to verify serial device is responding"""
import serial
import time

# Try to open the port and query
try:
    ser = serial.Serial(
        port="COM9",  # Adjust if your port is different
        baudrate=9600,
        bytesize=serial.EIGHTBITS,
        stopbits=serial.STOPBITS_ONE,
        timeout=1
    )
    
    print(f"Opened port: {ser.port}")
    print(f"Is open: {ser.is_open}")
    
    # Send query
    print("\nSending command: b'?MPOW'")
    ser.write(b'?MPOW')
    
    # Read response
    print("Waiting for response...")
    response = ser.read(64)
    
    print(f"\nResponse length: {len(response)} bytes")
    print(f"Response (raw): {response}")
    print(f"Response (hex): {response.hex()}")
    
    if response:
        ascii_str = response.decode(errors='ignore').strip()
        print(f"Response (ASCII): '{ascii_str}'")
    else:
        print("No response received!")
    
    ser.close()
    print("\nPort closed")
    
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
