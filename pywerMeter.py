from pywerHelper import serialComm, excelHelper

print("Testing pywerHelper package")
print(serialComm.testFunction())
print(excelHelper.testFunction())

if __name__ == "__main__":
    print("Attempting to read serial data.")
    serialComm.recieve_data()
