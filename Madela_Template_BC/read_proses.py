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
                self.fromulaList.append({
                    'cell' : cellName,
                    'text':formula
                })


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
        self.sheet = self.workbook.Worksheets.Add(self.__switsSheetNameFix())
        self.sheet.CopyFrom(self.deSheet)
        #close deworkbook
        self.deWorkbook.Dispose()

    def __switsSheetNameFix(self,start='FIX_',end='_FIX',hasStart=False,hasEnd=False):
        if self.deSheetName.find(start) != -1:
            hasStart=True
        
        if self.deSheetName.find(end) != -1:
            hasEnd=True
        
        if hasStart and hasEnd:
            return self.deSheetName
        else:
            if hasStart:
                return f'{self.deSheetName.replace(start,'')}{end}'
            elif hasEnd:
                return f'{start}{self.deSheetName.replace(end,'')}'
            else:
                return self.deSheetName
            
    def saveFile(self,pathToSave):
        self.workbook.SaveToFile(f'{pathToSave}{self.xlsxName}.xlsx')
        self.workbook.Dispose()

    def editText(self,cell,text):
        self.sheet.Range[f'{cell}'].Text = text










