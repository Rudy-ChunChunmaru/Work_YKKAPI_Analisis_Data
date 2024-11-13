from spire.xls import *
from spire.xls.common import *

class ReadFormulaXlsx:
    def __init__(self,pathToGet,xlsxName):
        self.workbook = Workbook()
        self.workbook.LoadFromFile(f'{pathToGet}{xlsxName}.xlsx')
        self.sheet = self.workbook.Worksheets[0]
        self.sheetName = self.sheet.Name
        self.fromulaList = []

    def findFormula(self):
        usedRange = self.sheet.AllocatedRange
        for cell in usedRange:
            if(cell.HasFormula):
                cellName = cell.RangeAddressLocal
                formula = cell.Formula
                self.fromulaList.append({cellName:formula})


    def endReadProses(self):
        self.workbook.Dispose()


class CopyWorkBook:
    def __init__(self,pathToGet,xlsxName):
        self.xlsxName = xlsxName
        # old deworkbook
        self.deWorkbook = Workbook()
        self.deWorkbook.LoadFromFile(f'{pathToGet}{self.xlsxName}.xlsx')
        self.deSheet = self.deWorkbook.Worksheets[0]
        self.deSheetName = self.deSheet.Name
        # new workbook
        self.workbook = Workbook()
        self.workbook.Worksheets.Add(desheetName)
        self.workbook.CopyFrom(deSheet)
        #close deworkbook
        self.deWorkbook.Dispose()

    def saveFile(self,pathToSave):
        self.workbook.SaveToFile(f'{pathToSave}{self.xlsxName}.xlsx')
        self.workbook.Dispose()










