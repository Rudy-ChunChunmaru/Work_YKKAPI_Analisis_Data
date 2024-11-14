from read_proses import ReadFormulaXlsx as rfx, CopyWorkBook as cwb
from data import listXlsx

pathToGet = './Madela_Template/'
pathToSave = './Proses_Madela_Template/'
# listXlsx = ['10001','10010']


CellNotForReplaceFromula = [
    # colom page
    'AF2','AV2','BL2','CB2',
    # colom H AND W
    'K9','L9',
    'K10','L10',
    # ORDER
    'E3','E4','E5','E6','E7','E8','E9','E10',
    'U3','U4','U5','U6','U7','U8','U9','U10',
    'AK3','AK4','AK5','AK6','AK7','AK8','AK9','AK10',
    'BA3','BA4','BA5','BA6','BA7','BA8','BA9','BA10',
    'BQ3','BQ4','BQ5','BQ6','BQ7','BQ8','BQ9','BQ10',
    # colom standar page 1
    'C14','C15','C16','C17','C18',
    'E14','E15','E16','E17','E18',
    'H14','H15','H16','H17','H18',
    'J14','J15','J16','J17','J18',
    'L14','L15','L16','L17','L18',
    'N14','N15','N16','N17','N18',
    'P14','P15','P16','P17','P18',
    # colom Footer page 1
    'F42','O42','P42'
]

ReplaceCell = [
    {'cell':'J4','text':'Online','hasFormula':False},
    {'cell':'J6','text':'Online','hasFormula':False},
    {'cell':'Z4','text':'Online','hasFormula':True},
    {'cell':'Z6','text':'Online','hasFormula':True},
    {'cell':'AP4','text':'Online','hasFormula':True},
    {'cell':'AP6','text':'Online','hasFormula':True},
    {'cell':'BF4','text':'Online','hasFormula':True},
    {'cell':'BF6','text':'Online','hasFormula':True},
    {'cell':'BV4','text':'Online','hasFormula':True},
    {'cell':'BV6','text':'Online','hasFormula':True},
]

ListHeaderPage = ['B1','R1','AH1','AX1','BN1'] 

typeSheetPage3 = {
    'fabricationRowFrom': 22,
    'fabricationRowTo': 48,
    'fabricationColom': [
        {
        'cell':'U',
        'listCell':['R','S','T','U','v','W','X','Y','Z','AA','AB','AC','AD','AE','AF']
        },
    ],
    'partRowFrom': 22,
    'partRowTo': 60,
    'partColom': [
        {
        'cell':'AK',
        'listCell':['AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV']
        },
    ],
}
  
typeSheetPage5 = {
    'fabricationRowFrom': 22,
    'fabricationRowTo': 48,
    'fabricationColom': [
        {
        'cell':'U',
        'listCell':['R','S','T','U','v','W','X','Y','Z','AA','AB','AC','AD','AE','AF']
        },
        {
        'cell':'AK',
        'listCell':['AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV']
        },
    ],
    'partRowFrom': 22,
    'partRowTo': 60,
    'partColom': [
        {
        'cell':'BA',
        'listCell':['AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL']
        },
        {
        'cell':'BQ',
        'listCell':['BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ','CA','CB']
        },
    ],
}
   

# TODO-MAIN:
def mainRun(xlsxName):
    prosesFormula = rfx(pathToGet=pathToGet,xlsxName=xlsxName)
    prosesFormula.findFormula()
    dataSheet = prosesFormula.sheet
    workBook = cwb(pathToGet=pathToGet,xlsxName=xlsxName)
    listFormula = prosesFormula.fromulaList
    listCellHasFromula = [valueListFromula['cell'] for valueListFromula in listFormula]
    

    # TODO-HEADERPAGE:
    def ColomHeaderPage():
        return [dataSheet.Range[f'{valueHeaderPage}'].Text for valueHeaderPage in ListHeaderPage]
        
    # TODO-FROMULA:
    def ColomFromula():
        for value in listFormula:
            ColomCell = value['cell']
            ValueText = value['text']
            ReplaseLogic = True
            if ColomCell in CellNotForReplaceFromula:
                ReplaseLogic = False
            if ReplaseLogic:
                workBook.editText(cell=ColomCell,text=f"'{ValueText}")


    # TODO-REPLACE:
    def ColomReplace():
        for ValueReplaceCell in ReplaceCell:
            if not ValueReplaceCell['hasFormula']:
                workBook.editText(cell=ValueReplaceCell['cell'],text=ValueReplaceCell['text'])
            else:
                if ValueReplaceCell['cell'] in listCellHasFromula:
                    workBook.editText(cell=ValueReplaceCell['cell'],text=ValueReplaceCell['text'])


    # TODO-Empty Text
    def ColomEmpty():
        dataHeaderSheet = ColomHeaderPage()
        countDataHeaderSheet  = 0
        for valueDataHeaderSheet in dataHeaderSheet:
            if valueDataHeaderSheet:
                countDataHeaderSheet+=1

        if countDataHeaderSheet == 3:
            sheetpage=typeSheetPage3
        elif countDataHeaderSheet == 5:
            sheetpage=typeSheetPage5
        else:
            sheetpage={}

        if not sheetpage == {}:
            # fabrication
            for valueColomCell in sheetpage['fabricationColom']:
                for valueRowCell in range(sheetpage['fabricationRowFrom'],sheetpage['fabricationRowTo']):
                    if not workBook.sheet.Range[f'{valueColomCell['cell']}{valueRowCell}'].Text:
                        for valueListCell in valueColomCell['listCell']:
                            workBook.editText(cell=f"{valueListCell}{valueRowCell}",text="")
            # part
            for valueColomCell in sheetpage['partColom']:
                for valueRowCell in range(sheetpage['partRowFrom'],sheetpage['partRowTo']):
                    if not workBook.sheet.Range[f'{valueColomCell['cell']}{valueRowCell}'].Text:
                        for valueListCell in valueColomCell['listCell']:
                            workBook.editText(cell=f"{valueListCell}{valueRowCell}",text="")

                



    ColomFromula()
    ColomReplace()
    ColomEmpty()

    workBook.saveFile(pathToSave)
    prosesFormula.endReadProses()


for valueXlsx in listXlsx:
    mainRun(valueXlsx)