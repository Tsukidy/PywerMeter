# Python script for handling serial communication wtih a device.
import serial

def testFunction():
    return "This is a test function from serialComm.py"

def returnSerialPorts():
    import serial.tools.list_ports
    ports = serial.tools.list_ports.comports()
    return [port.device for port in ports]

def recieve_data(port, baudrate, timeout, stopbits, bytesize):
    try:
        while True:
            # Open serial connection
            with serial.Serial('COM9', 9600, timeout=1) as ser:
                # Send a ? to the device to request data
                ser.write(b'?')
                # Read a line from the serial port
                line = ser.readline()
                text = line.decode('utf-8').rstrip()
                if line:
                    print(f"Received: {text}")
                    print(f"Debug output: {line}")
                    sleep(1)  # Pause for a second before next read
    except KeyboardInterrupt:
        Print("Serial communication terminated by user.")
    finally:
        exit(0)

if __name__ == "__main__":
    print ("Returning available serial ports:")
    returned_ports = returnSerialPorts()
    print(returned_ports)
