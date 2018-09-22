import extractCommonData
import xlsxwriter


def createNewExcelSheet(directory,listOfPlaced):
    newBook = xlsxwriter.Workbook(directory+"CommonPlacementList.xlsx")
    newSheet = newBook.add_worksheet()
    newSheet.set_column('A:A', 20)
    newSheet.set_column('B:B', 30)
    newSheet.set_column('C:C', 10)
    newSheet.set_column('D:D', 10)
    newSheet.set_column('E:E', 30)
    row=0
    col=0
    for registerNo,name,campus,degree,branch in listOfPlaced:
        newSheet.write(row,col,registerNo)
        newSheet.write(row,col+1,name)
        newSheet.write(row,col+2,campus)
        newSheet.write(row,col+3,degree)
        newSheet.write(row,col+4,branch)
        row+=1
    newBook.close()


def driver():
    
    """The files InfosysResult.xlsx, TCSResult.xlsx, CognizantResult.xlsx, and WiproResult.xlsx are already available in the
    data folder, just add the directory of the folder where the cloned project is kept. Feel free to change the directory string
    to achieve the perfect directory in the 'fileList' tuple"""

    directory = r"C:/Users/Aftab Alam/Documents/GitHub"
    directory = directory + r"/SRM-placement-analyser/data/"
    fileList = [directory+"InfosysResult.xlsx",directory+"TCSResult.xlsx",directory+"CognizantResult.xlsx",directory+"WiproResult.xlsx"]
    
    listOfPlaced = extractCommonData.extractCommonData(fileList)
    createNewExcelSheet(directory,listOfPlaced)
