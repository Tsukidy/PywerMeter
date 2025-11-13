from pywerHelper import serialComm, excelHelper
import time

if __name__ == "__main__":
    dev = serialComm.SerialDevice(port="COM9", baudrate=38400, timeout=0.5)
    dataLog = ""
    minutes = 0.25
    end_time = time.time() + minutes * 60
    print("Attempting to read serial data.")
    while time.time() < end_time:
        unformattedData, retHexData, retAsciiData  = dev.query(command=b'?MPOW')
        dataLog += retAsciiData + "\n"
        print(f"Received serial data: {retAsciiData}")
    dev.close()
    with open ("serialDataLog.log", "w") as file:
        try:
            file.write(dataLog)
        except Exception as e:
            print(f"Error writing to log file: {e}")
