from pywerHelper import serialComm, excelHelper
import time

print("Testing pywerHelper package")
print(serialComm.testFunction())
print(excelHelper.testFunction())

if __name__ == "__main__":
    dev = serialComm.SerialDevice(port="COM9", baudrate=38400, timeout=0.5)
    dataLog = ""
    minutes = 1
    end_time = time.time() + minutes * 60
    while time.time() < end_time:
        print("Attempting to read serial data.")
        unformattedData, retHexData, retAsciiData  = dev.query(command=b'?MPOW')
        dataLog += retAsciiData + "\n"
        print(f"Received serial data: {retAsciiData}")
    dev.close()
    with open ("serialDataLog.log", "w") as file:
        file.write(dataLog)
