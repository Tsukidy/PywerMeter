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
        try:
            self.ser = serial.Serial(
                port=port,
                baudrate=baudrate,
                bytesize=bytesize,
                stopbits=stopbits,
                timeout=timeout
            )
        except serial.SerialException as e:
            print(f"Error opening serial port: {e}")
            self.ser = None
            raise

    def query(
        self,
        command=b'?MPOW'
    ):
        if not self.ser or not self.ser.is_open:
            raise serial.SerialException("Serial port is not open.")
            return none, None, None
        try:
            self.ser.write(command)
            response = self.ser.read(64)
            ascii_str = response.decode(errors='ignore')
            trimmed_ascii = ascii_str.strip()
            hex_str = response.hex()
            spaced_hex = ' '.join([hex_str[i:i+2] for i in range(0, len(hex_str), 2)])
            return response, spaced_hex, trimmed_ascii
        except serial.SerialException as e:
            print(f"Error during serial communication: {e}")
            return None, None, None

    def close(self):
        if self.ser and self.ser.is_open:
            try:
                self.set.close()
            except serial.SerialException as e:
                print(f"Error closing serial port: {e}")

if __name__ == "__main__":
    returnSerialPorts()
