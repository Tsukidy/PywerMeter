from pywerHelper import serialComm, excelHelper

print("Testing pywerHelper package")
print(serialComm.testFunction())
print(excelHelper.testFunction())

if __name__ == "__main__":
    dev = serialComm.SerialDevice(port="COM9")
    for _ in range(10):
        print("Attempting to read serial data.")
        seialData = dev.query()
        dataLog += serialData
        print(f"Received serial data: {seialData}")
    dev.close()
    with open ("serialDataLog.log", "w") as file:
        file.write(dataLog)
