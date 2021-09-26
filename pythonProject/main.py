import os, glob, openpyxl, pandas as pd

ServerfolderPath = "C:/serverexample/"
os.chdir(ServerfolderPath)

filexlsxarray = []
ArrayRoW = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']


for filexls in glob.glob("*.xls"):
    print(ServerfolderPath+filexls)

for filexlsx in glob.glob("*.xlsx"):
    print(ServerfolderPath+filexlsx)


try: filexlsx
except:
    print()
    try:
        filexls
    except:
        print("ERROR EXCEL FILE NOT FOUND!!!! Check excel file with DataSet from path: ", ServerfolderPath)
    else:
        print()
        if (str(ServerfolderPath + filexls)) != None:
            df = pd.read_excel(ServerfolderPath + filexls)
            df.to_excel(ServerfolderPath + filexls + "x")
            os.remove(ServerfolderPath + filexls)
else:
    for file in glob.glob("*.xlsx"):
        print()
        filexlsxarray.append(file)

print(filexlsxarray)

arrayindex = 0



while arrayindex < len(filexlsxarray):
    WorkBook = openpyxl.load_workbook(ServerfolderPath + filexlsxarray[arrayindex])
    sheet = WorkBook.active

    CellIndex = "B"
    CellNumber = 1
    CellNumberStart = 2
    print("Wait some times")
    while (sheet[CellIndex + str(CellNumber)].value) != None:
        CellNumber = CellNumber + 1

    CellIndex = "C"
    while CellNumberStart + (CellNumber) - 120 <= CellNumber:
        if CellNumberStart == 2:
            MiddleRangeTemp = sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value
        else:
            if (sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) != None:
                MiddleRangeTemp = (MiddleRangeTemp + sheet[
                    str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) / 2
                CellNumberStart = CellNumberStart + 1
                print("MIDDLE TEMPERATURE", MiddleRangeTemp)
        CellNumberStart = CellNumberStart + 1

    CellIndex = "D"
    CellNumberStart = 2
    while CellNumberStart + (CellNumber) - 120 <= CellNumber:
        if CellNumberStart == 2:
            AirHumidity = sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value
        else:
            if (sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) != None:
                v = (AirHumidity + sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) / 2
                CellNumberStart = CellNumberStart + 1
                print("MIDDLE AirHumidity", AirHumidity)
        CellNumberStart = CellNumberStart + 1

    CellIndex = "E"
    CellNumberStart = 2
    while CellNumberStart + (CellNumber) - 120 <= CellNumber:
        if CellNumberStart == 2:
            CO2 = sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value
        else:
            if (sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) != None:
                v = (CO2 + sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) / 2
                CellNumberStart = CellNumberStart + 1
                print("MIDDLE CO2", CO2)
        CellNumberStart = CellNumberStart + 1

    CellIndex = "F"
    CellNumberStart = 2
    while CellNumberStart + (CellNumber) - 120 <= CellNumber:
        if CellNumberStart == 2:
            VOC = sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value
        else:
            if (sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) != None:
                v = (VOC + sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) / 2
                CellNumberStart = CellNumberStart + 1
                print("MIDDLE VOC", VOC)
        CellNumberStart = CellNumberStart + 1

    CellIndex = "G"
    CellNumberStart = 2
    while CellNumberStart + (CellNumber) - 120 <= CellNumber:
        if CellNumberStart == 2:
            Dust1 = sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value
        else:
            if (sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) != None:
                v = (Dust1 + sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) / 2
                CellNumberStart = CellNumberStart + 1
                print("MIDDLE Dust1", Dust1)
        CellNumberStart = CellNumberStart + 1

    CellIndex = "H"
    CellNumberStart = 2
    while CellNumberStart + (CellNumber) - 120 <= CellNumber:
        if CellNumberStart == 2:
            Dust2_5 = sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value
        else:
            if (sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) != None:
                v = (Dust2_5 + sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) / 2
                CellNumberStart = CellNumberStart + 1
                print("MIDDLE Dust2_5", Dust2_5)
        CellNumberStart = CellNumberStart + 1

    CellIndex = "I"
    CellNumberStart = 2
    while CellNumberStart + (CellNumber) - 120 <= CellNumber:
        if CellNumberStart == 2:
            Dust10 = sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value
        else:
            if (sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) != None:
                v = (Dust10 + sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) / 2
                CellNumberStart = CellNumberStart + 1
                print("MIDDLE Dust10", Dust10)
        CellNumberStart = CellNumberStart + 1

    CellIndex = "J"
    CellNumberStart = 2
    while CellNumberStart + (CellNumber) - 120 <= CellNumber:
        if CellNumberStart == 2:
            Pressure = sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value
        else:
            if (sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) != None:
                v = (Pressure + sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) / 2
                CellNumberStart = CellNumberStart + 1
                print("MIDDLE Pressure", Pressure)
        CellNumberStart = CellNumberStart + 1

    CellIndex = "K"
    CellNumberStart = 2
    while CellNumberStart + (CellNumber) - 120 <= CellNumber:
        if CellNumberStart == 2:
            AQI = sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value
        else:
            if (sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) != None:
                v = (AQI + sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) / 2
                CellNumberStart = CellNumberStart + 1
                print("MIDDLE AQI", AQI)
        CellNumberStart = CellNumberStart + 1

    CellIndex = "L"
    CellNumberStart = 2
    while CellNumberStart + (CellNumber) - 120 <= CellNumber:
        if CellNumberStart == 2:
            Formaldehyde = sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value
        else:
            if (sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) != None:
                v = (Formaldehyde + sheet[str(CellIndex) + str(CellNumberStart + (CellNumber) - 120)].value) / 2
                CellNumberStart = CellNumberStart + 1
                print("MIDDLE Formaldehyde", Formaldehyde)
        CellNumberStart = CellNumberStart + 1

    WorkBook.close()
    WorkBook = openpyxl.load_workbook(str(ServerfolderPath) + "/DATASETFOLDER/" + 'DataSet.xlsx')
    sheet = WorkBook.active

    sheet["A2"] = 'MiddleRangeTemp'
    sheet["A3"] = 'AirHumidity'
    sheet["A4"] = 'MiddleRangeTemp'
    sheet["A5"] = 'CO2'
    sheet["A6"] = 'VOC'
    sheet["A7"] = 'Dust1'
    sheet["A8"] = 'Dust2_5'
    sheet["A9"] = 'Dust10'
    sheet["A10"] = 'Pressure'
    sheet["A11"] = 'AQI'
    sheet["A12"] = 'Formaldehyde'
    print()
    print()
    print("DATASET")
    print()

    Row = ArrayRoW[arrayindex + 1]

    sheet[str(Row) + "1"] = filexlsxarray[arrayindex]
    print("MidRange MiddleRangeTemp: ", MiddleRangeTemp)
    sheet[str(Row) + "2"] = MiddleRangeTemp
    print("MidRange AirHumidity: ", AirHumidity)
    sheet[str(Row) + "3"] = AirHumidity
    print("MidRange MiddleRangeTemp: ", MiddleRangeTemp)
    sheet[str(Row) + "4"] = MiddleRangeTemp
    print("MidRange CO2: ", CO2)
    sheet[str(Row) + "5"] = CO2
    print("MidRange VOC: ", VOC)
    sheet[str(Row) + "6"] = VOC
    print("MidRange Dust1: ", Dust1)
    sheet[str(Row) + "7"] = Dust1
    print("MidRange Dust2_5: ", Dust2_5)
    sheet[str(Row) + "8"] = Dust2_5
    print("MidRange Dust10: ", Dust10)
    sheet[str(Row) + "9"] = Dust10
    print("MidRange Pressure: ", Pressure)
    sheet[str(Row) + "10"] = Pressure
    print("MidRange AQI: ", AQI)
    sheet[str(Row) + "11"] = AQI
    print("MidRange Formaldehyde: ", Formaldehyde)
    sheet[str(Row) + "12"] = Formaldehyde

    WorkBook.save(str(ServerfolderPath) + '/DATASETFOLDER/DataSet.xlsx')

    arrayindex = arrayindex + 1
