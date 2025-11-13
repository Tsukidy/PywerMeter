# Python script for handling serial communication wtih a device.
import serial

def returnSerialPorts():
    import serial.tools.list_ports
    ports = serial.tools.list_ports.comports()
    for port in ports:
        print(f"{port.device}: {port.description}")

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
        command=b'?MPOW'
    ):
        self.ser.write(command)
        response = self.ser.read(64)
        ascii_str = response.decode(errors='ignore')
        trimmed_ascii = ascii_str.strip()
        hex_str = response.hex()
        spaced_hex = ' '.join([hex_str[i:i+2] for i in range(0, len(hex_str), 2)])
        return response, spaced_hex, trimmed_ascii

    def close(self):
        self.ser.close()

if __name__ == "__main__":
    returnSerialPorts()
