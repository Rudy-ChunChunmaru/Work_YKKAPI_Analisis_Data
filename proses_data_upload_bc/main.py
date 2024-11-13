from read_proses import ReadFormulaXlsx as rfx, CopyWorkBook as cwb

pathToGet = './Madela_Template/'
pathToSave = './Proses_Madela_Template/'
xlsxName = '10001'


prosesFormula = rfx(pathToGet=pathToGet,xlsxName=xlsxName)
prosesFormula.findFormula()
listFormula = prosesFormula.fromulaList

workBook = cwb(pathToGet=pathToGet,xlsxName=xlsxName)

CellNotForReplaceFromula = [
    'AF2','AV2',
    'E3','E4','E5','E6','E7','E8','E9','E10',
    'K9','L9',
    'K10','L10',
    'U3','U4','U5','U6','U7','U8','U9','U10',
    'AK3','AK4','AK5','AK6','AK7','AK8','AK9','AK10',
    'F42','O42','P42'
]

ReplaceCell = [
    {'cell':'J4','text':'Online','hasFormula':True},
    {'cell':'J4','text':'Online','hasFormula':True},
]

for value in listFormula:
    if not value['cell'] in CellNotForReplaceFromula:
        workBook.editText(cell=value['cell'],text=f"'{value['text']}")

workBook.saveFile(pathToSave)
prosesFormula.endReadProses()