from spire.xls import *
from spire.xls.common import *

class ReadFormulaXlsx:
    def __init__(self,pathXlsx):
        self.workbook = Workbook()
        self.workbook.LoadFromFile(pathXlsx)
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





