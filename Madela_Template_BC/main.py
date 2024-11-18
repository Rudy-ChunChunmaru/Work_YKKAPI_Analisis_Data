from read_proses import ReadFormulaXlsx as rfx, CopyWorkBook as cwb
from database import Database_Shiage as shiage
from data import listXlsx,listXlsxRudy

pathToGet = './Madela_Template/'
pathToSave = './Proses_Madela_Template/'
listXlsx = listXlsxRudy


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
    # Page 2 Total
    {'cell':'AF49','text':"'=(AF48*0.986)",'hasFormula':True},
    {'cell':'AF50','text':"'=((AF48*0.974)*0.986)",'hasFormula':True},
    # Page 3 Total
    {'cell':'AV49','text':"'=(AV48*0.986)",'hasFormula':True},
    {'cell':'AV50','text':"'=((AV48*0.974)*0.986)",'hasFormula':True},
]

ListHeaderPage = ['B1','R1','AH1','AX1','BN1'] 

typeSheetPage3 = {
    'fabricationRowFrom': 22,
    'fabricationRowTo': 48,
    'fabricationColom': [
        {
        'cell':'U',
        'qty':'Z',
        'weight':'AF',
        'listCell':['R','S','T','U','v','W','X','Y','Z','AA','AB','AC','AD','AE','AF']
        },
    ],
    'partRowFrom': 22,
    'partRowTo': 61,
    'partColom': [
        {
        'cell':'AK',
        'qty':'AP',
        'weight':'AV',
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
        'qty':'Z',
        'weight':'AF',
        'listCell':['R','S','T','U','v','W','X','Y','Z','AA','AB','AC','AD','AE','AF']
        },
        {
        'cell':'AK',
        'qty':'AP',
        'weight':'AV',
        'listCell':['AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV']
        },
    ],
    'partRowFrom': 22,
    'partRowTo': 61,
    'partColom': [
        {
        'cell':'BA',
        'qty':'BF',
        'weight':'BL',
        'listCell':['AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL']
        },
        {
        'cell':'BQ',
        'qty':'BV',
        'weight':'CB',
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
    dataHeaderSheet = []

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


    # TODO-Empty Text AND Fix Fromula
    dataHeaderSheet = ColomHeaderPage()
    def ColomEmptyANDFixFromula():
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


        def ReplacementFromQtyColom(value,consStart,consEnd):
            try:
                listIfCons = value.replace(consStart,'').replace(consEnd,'').split(',')
                if(len(listIfCons) == 3):
                    return f"'{consStart}{listIfCons[0]},{listIfCons[1]},({listIfCons[2]}){consEnd}"
                else:
                    return '$FAIL$'
            except:
                return '$FAIL$'
            
        def ReplacementFromWeightColom(value,consStart,consEnd):
            try:
                listIfCons = value.replace(consStart,'').replace(consEnd,'').split(',')
                if(len(listIfCons) == 3):
                    listFromula = listIfCons[1].split('*')
                    if(len(listFromula) == 3):
                        return f"'{consStart}{listIfCons[0]},(({listFromula[0]}*{listFromula[1]})*{listFromula[2].split('/')[0]})/{listFromula[2].split('/')[1]}),{listIfCons[2]}{consEnd}"
                    else:
                        return '$FAIL$'
                else:
                    return '$FAIL$'
            except:
                return '$FAIL$'
            
        def ReplacementFromulaCellColom(BomID,FormulaCode):
            dbShiage = shiage()
            getDataFromulaCell = dbShiage.GetMadelaFormula(
                strWhere=f" BOM_ID ='{BomID}' AND  Type='P' AND FormulaCode='{FormulaCode}'"
            )
            if len(getDataFromulaCell) >= 1:
                strText = f"'={getDataFromulaCell[0].Formula}"
                return strText
            else:
                return FormulaCode

        # start
        if not sheetpage == {}:
            # fabrication
            for valueColomCell in sheetpage['fabricationColom']:
                for valueRowCell in range(sheetpage['fabricationRowFrom'],sheetpage['fabricationRowTo']):
                    if not workBook.sheet.Range[f'{valueColomCell['cell']}{valueRowCell}'].Text:
                        for valueListCell in valueColomCell['listCell']:
                            workBook.editText(cell=f"{valueListCell}{valueRowCell}",text="")
                    else:
                        # qty
                        if workBook.sheet.Range[f'{valueColomCell['qty']}{valueRowCell}'].Text:
                            workBook.editText(
                                cell=f"{valueColomCell['qty']}{valueRowCell}",
                                text=ReplacementFromQtyColom(
                                    value=workBook.sheet.Range[f'{valueColomCell['qty']}{valueRowCell}'].Text,
                                    consStart='=IF(',
                                    consEnd=')'
                                )
                            )
                        # weight
                        if workBook.sheet.Range[f'{valueColomCell['weight']}{valueRowCell}'].Text:
                            workBook.editText(
                                cell=f"{valueColomCell['weight']}{valueRowCell}",
                                text=ReplacementFromWeightColom(
                                    value=workBook.sheet.Range[f'{valueColomCell['weight']}{valueRowCell}'].Text,
                                    consStart='=IF(',
                                    consEnd=')'
                                )
                            )
                            
                        
            # part
            for valueColomCell in sheetpage['partColom']:
                for valueRowCell in range(sheetpage['partRowFrom'],sheetpage['partRowTo']):
                    if not workBook.sheet.Range[f'{valueColomCell['cell']}{valueRowCell}'].Text:
                        for valueListCell in valueColomCell['listCell']:
                            workBook.editText(cell=f"{valueListCell}{valueRowCell}",text="")
                    else:
                        # qty
                        if workBook.sheet.Range[f'{valueColomCell['qty']}{valueRowCell}'].Text:
                            workBook.editText(
                                cell=f"{valueColomCell['qty']}{valueRowCell}",
                                text=ReplacementFromQtyColom(
                                    value=workBook.sheet.Range[f'{valueColomCell['qty']}{valueRowCell}'].Text,
                                    consStart='=IF(',
                                    consEnd=')'
                                )
                            )
                        # Fromula
                        if workBook.sheet.Range[f'{valueColomCell['cell']}{valueRowCell}'].Text.find('P-') != -1:
                            workBook.editText(
                                cell=f"{valueColomCell['cell']}{valueRowCell}",
                                text=ReplacementFromulaCellColom(
                                    BomID=xlsxName,
                                    FormulaCode=workBook.sheet.Range[f'{valueColomCell['cell']}{valueRowCell}'].Text
                                )
                            )

    ColomFromula()
    ColomReplace()
    ColomEmptyANDFixFromula()

    workBook.saveFile(pathToSave)
    prosesFormula.endReadProses()


for valueXlsx in listXlsx:
    mainRun(valueXlsx)