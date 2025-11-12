from pywerHelper import serialComm, excelHelper
import time

print("Testing pywerHelper package")
print(serialComm.testFunction())
print(excelHelper.testFunction())

if __name__ == "__main__":
    dev = serialComm.SerialDevice(port="COM9")
    dataLog = ""
    minutes = 1
    end_time = time.time() + minutes * 60
    #for _ in range(10):
    while time.time() < end_time:
        print("Attempting to read serial data.")
        serialData = dev.query()
        dataLog += serialData + "\n"
        print(f"Received serial data: {serialData}")
    dev.close()
    with open ("serialDataLog.log", "w") as file:
        file.write(dataLog)
