# Python script for handling serial communication wtih a device.
import serial

def testFunction():
    return "This is a test function from serialComm.py"

def returnSerialPorts():
    import serial.tools.list_ports
    ports = serial.tools.list_ports.comports()
    for port in ports:
        print(f"{port.device}: {port.description}")

def query_device(
    command=b'?',
    port="COM9",
    baudrate=9600,
    bytesize=serial.EIGHTBITS,
    stopbits=serial.STOPBITS_ONE,
    timeout=1
):
    with serial.Serial(port=port, baudrate=baudrate, bytesize=bytesize, stopbits=stopbits, timeout=timeout) as ser:
        ser.write(command)
        response = ser.read(64)  # Adjust size as needed
        hex_str = response.hex()
        spaced_hex = ' '.join([hex_str[i:i+2] for i in range(0, len(hex_str), 2)])
        return spaced_hex

class SerialDevice:
    def __init__(
        self,
        port="COM9",
        baudrate=9600,
        bytesize=serial.EIGHTBITS,
        stopbits=serial.STOPBITS_ONE,
        timeout=1
    ):
        self.ser = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=bytesize,
            stopbits=stopbits,
            timeout=timeout
        )

    def query(
        self,
        command=b'?'
    ):
        self.ser.write(command)
        response = self.ser.read(64)
        hex_str = response.hex()
        spaced_hex = ' '.join([hex_str[i:i+2] for i in range(0, len(hex_str), 2)])
        return spaced_hex

    def close(self):
        self.ser.close()

