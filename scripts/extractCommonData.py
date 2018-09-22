import xlrd


def isPlaced(registerNo,listOfPlaced,noOfPlaced):
    x=0
    if noOfPlaced < 1:
        return -1
    while x < noOfPlaced:
        if listOfPlaced[x][0] == registerNo:
            return x
        x+=1
    return -1



def findTheStartRowOfData(sheet):
    row = 0
    col = 0
    
    #Iterate till the row containing titles for the excel sheet is found
    while sheet.cell_value(row,col) != "S.NO":
        row+=1
    row+=1          #Real data starts from the row next to row containing the titles 
    return row


def extractCommonData(fileList):
    
    listOfPlaced = []                       #For saving the details about placed students
    noOfPlaced = 0

    for singleFile in fileList:
        
        wb = xlrd.open_workbook(singleFile)     #Opens the excel documentat 
        sheet = wb.sheet_by_index(0)            #Open the first excel sheet of the document
        rowNo=0                                 #For keeping record of the row number, used to access specific row number of the excel sheet
        colNo=1                                 #For keeping record of the row number being accessed, and 2nd colum contains register no
        numberOfRecords = sheet.nrows           #Total number of rows in the sheet

        rowNo = findTheStartRowOfData(sheet)


        #Iterate till the end of the rows or untill meaningful data is present
        while( rowNo < numberOfRecords ):

            #If the row contains meaningful data continue else break the while loop
            if not(sheet.cell_value(rowNo,colNo).startswith("RA")):
                break

            #Find if the student in this excel file's 'rowNo' row is already a placed student
            index = isPlaced(sheet.cell_value(rowNo, colNo),listOfPlaced,noOfPlaced)
            if index > -1:        #If yes then increment the noOfOffers and continue the while loop
                listOfPlaced[index][5] = listOfPlaced[index][5]+1
            else:
                listOfPlaced.append(sheet.row_values(rowNo)[1:])
                noOfPlaced+=1
                listOfPlaced[noOfPlaced-1].append(1)
            rowNo+=1

        rowNo-=1    #Decrement the rowNo as it doesn't contain meaningful data

    return listOfPlaced
