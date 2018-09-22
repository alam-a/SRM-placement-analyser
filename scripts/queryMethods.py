import createNewExcelSheet
import xlrd

directory = r"C:/Users/Aftab Alam/Documents/GitHub"
directory = directory + r"/SRM-placement-analyser/data/"

def printNoOfMultipleOffers():
    oneOffer = 0
    twoOffers = 0
    threeOffers = 0
    fourOffers = 0
    wb = xlrd.open_workbook(directory+"CommonPlacementList.xlsx")
    ws = wb.sheet_by_index(0)
    noOfRecords = ws.nrows
    for row in range(noOfRecords):
        value = ws.cell_value(row,5)
        if value == 1:
            oneOffer += 1
        elif value == 2:
            twoOffers += 1
        elif value == 3:
            threeOffers += 1
        elif value == 4:
            fourOffers += 1
    print("One offer: " + str(oneOffer))
    print("Two offer: " + str(twoOffers))
    print("Three offer: " + str(threeOffers))
    print("Four offer: " + str(fourOffers))
printNoOfMultipleOffers()