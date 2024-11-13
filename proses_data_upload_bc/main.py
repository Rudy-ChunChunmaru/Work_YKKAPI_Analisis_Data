from read_proses import ReadFormulaXlsx as rfx

pathToGet = './Madela_Template/10001.xlsx'
pathToSave = './Proses_Madela_Template/'
xlsxName = ''


prosesFormula = rfx(pathToGet=pathToGet,xlsxName=xlsxName)
# prosesFormula.findFormula()
# listFormula = prosesFormula.fromulaList

# workBook = cwb(pathToGet=pathToGet,xlsxName=xlsxName)


# Workbook.saveFile(pathToSave)
# prosesFormula.endReadProses()