from pywerHelper import serialComm, excelHelper

print("Testing pywerHelper package")
print(serialComm.testFunction())
print(excelHelper.testFunction())

if __name__ == "__main__":
    while True:
        print("Attempting to read serial data.")
        seialData = serialComm.query_device()
        print(f"Received serial data: {seialData}")
