# Python script for handling serial communication wtih a device.
import serial

def testFunction():
    return "This is a test function from serialComm.py"

def recieve_data():
    try:
        while True:
            # Open serial connection
            with serial.Serial('/dev/ttyUSB0', 9600, timeout=1) as ser:
                # Read a line from the serial port
                line = ser.readline().decode('utf-8').rstrip()
                if line:
                    print(f"Received: {line}")
    except KeyboardInterrupt:
        Print("Serial communication terminated by user.")
    finally:
        exit(0)
