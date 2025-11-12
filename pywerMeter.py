from pywerHelper import serialComm, excelHelper

print("Testing pywerHelper package")
print(serialComm.testFunction())
print(excelHelper.testFunction())

if __name__ == "__main__":
    i = 0
    while i < 5:
        print("Attempting to read serial data.")
        seialData = serialComm.SerialDevice(port="COM9").query()
        print(f"Received serial data: {seialData}")
        i += 1
    serialComm.SerialDevice(port="COM9").close()
