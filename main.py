from txtFileCreator import *
from dataExtractor import *

if __name__ == '__main__':
    resultFile = open('upload_file.txt', 'w+')
    dataObj = DataPrepper("Input_Spreadsheet.xlsx")
    koteretData = dataObj.extractAndValidateKoteretData()
    tenuotData = dataObj.extractAndValidateTenuotData()
    errorList = dataObj.getErrorList()
    if len(errorList) > 0:  # if there are errors
        print("WARNING: The text file could not be created because of the following errors in your spreadsheet:\n")
        print("------------------------------------------------------------------------------------------------\n")
        for item in errorList:
            print(item)
    #else:
    writerObj = TxtCreator(resultFile, koteretData, tenuotData)
    resultFile.write(writerObj.ReshumatKoteret_Creator())
    resultFile.write(writerObj.ReshumatTenua_Creator())
    resultFile.write(writerObj.ReshumatTotal_Creator())
    resultFile.close()
