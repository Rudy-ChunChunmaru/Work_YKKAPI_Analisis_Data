from template.read_proses import ReadFormulaXlsx as rfx, CopyWorkBook as cwb
from database.db_shiage import Database_Shiage as shiage
from data.data import data

pathToGet = './data/Madela_Template/'
pathToSave = './data/Proses_Madela_Template/'
dataTemplate=data()
listXlsx = dataTemplate.listXlsxDiki

CellNotForReplaceFromula = [
    # colom page
    'AF2','AV2','BL2','CB2',
    # colom H AND W
    # 'K9','L9',
    # 'K10','L10',
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

ReplaceByTypeProduct = [
    {   
        'type':'frontera',
        'product':'door-Nonhendel',
        'bom':[
            10190,
            10191,
            10192,
            10193,
            10194,
            10195,
            10196,
            10197,

            10206,
            10207,
            10208,
            10209,
            10210,
            10211,
            10212,
            10213,
        ],
        'replacementHead':[
            {'cell':'J14','text':'ML'},
            {'cell':'J15','text':'YK1N'},
            {'cell':'J16','text':''},
            {'cell':'J17','text':'9K-40024-1'},
            {'cell':'J18','text':''},
        ],
        # Material
        'replacementMaterialOuter':[

        ],
        'replacementMaterialInner':[
            {'row':'AJ','rowtext':'AJ','text':"'=w.7"}
        ],
        # Part
        'replacementPartOuter':[
           
        ],
        'replacementPartInner':[
            {'rowValue':'BN','value':'WOODEN PANEL','rowtext':'BQ','text':"'=w.9"},
            {'rowValue':'BN','value':'WOODEN PANEL','rowtext':'BT','text':"'=w.6"}
        ],
    },
    {   
        'type':'frontera',
        'product':'db-door-Nonhendel',
        'bom':[
            10198,
            10199,
            10200,
            10201,
            10202,
            10203,
            10204,
            10205,

            10214,
            10215,
            10216,
            10217,
            10218,
            10219,
            10220,
            10221,
        ],
        'replacementHead':[
            {'cell':'J14','text':'ML'},
            {'cell':'J15','text':'YK1N'},
            {'cell':'J16','text':''},
            {'cell':'J17','text':'9K-40024-1'},
            {'cell':'J18','text':'9K-40024-2'},
        ],
        # Material
        'replacementMaterialOuter':[

        ],
        'replacementMaterialInner':[
            {'row':'AJ','rowtext':'AJ','text':"'=w.7"}
        ],
        # Part
        'replacementPartOuter':[
           
        ],
        'replacementPartInner':[
            {'rowValue':'BN','value':'WOODEN PANEL','rowtext':'BQ','text':"'=w.9",'num':1},
            {'rowValue':'BN','value':'WOODEN PANEL','rowtext':'BQ','text':"'=w.10",'num':2},
            {'rowValue':'BN','value':'WOODEN PANEL','rowtext':'BT','text':"'=w.6"}
        ],
    }
]



ListHeaderPage = ['B1','R1','AH1','AX1','BN1'] 

typeSheetPage3 = {
    'fabricationRowFrom': 22,
    'fabricationRowTo': 48,
    'fabricationColom': [
        {
        'group':'OUTER',
        'cell':'U',
        'qtyUnit':'Y',
        'qty':'Z',
        'remark':'AB',
        'weight':'AF',
        'FS':'',
        'color':'',
        'length':'X',
        'description':'R',
        'listCell':['R','S','T','U','v','W','X','Y','Z','AA','AB','AC','AD','AE','AF']
        },
    ],
    'partRowFrom': 22,
    'partRowTo': 61,
    'partColom': [
        {
        'group':'OUTER',
        'cell':'AK',
        'qtyUnit':'AO',
        'qty':'AP',
        'remark':'AR',
        'weight':'',
        'FS':'AV',
        'color':'AN',
        'length':'',
        'description':'AH',
        'listCell':['AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV']
        },
    ],
}
  
typeSheetPage5 = {
    'fabricationRowFrom': 22,
    'fabricationRowTo': 48,
    'fabricationColom': [
        {
        'group':'OUTER',
        'cell':'U',
        'qtyUnit':'Y',
        'qty':'Z',
        'remark':'AB',
        'weight':'AF',
        'FS':'',
        'color':'',
        'length':'X',
        'description':'R',
        'listCell':['R','S','T','U','v','W','X','Y','Z','AA','AB','AC','AD','AE','AF']
        },
        {
        'group':'INNER',
        'cell':'AK',
        'qtyUnit':'AO',
        'qty':'AP',
        'remark':'AS',
        'weight':'AV',
        'FS':'',
        'color':'',
        'length':'AN',
        'description':'AH',
        'listCell':['AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV']
        },
    ],
    'partRowFrom': 22,
    'partRowTo': 61,
    'partColom': [
        {
        'group':'OUTER',
        'cell':'BA',
        'qtyUnit':'BE',
        'qty':'BF',
        'remark':'BH',
        'weight':'',
        'FS':'BL',
        'color':'BD',
        'length':'',
        'description':'AX',
        'listCell':['AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL']
        },
        {
        'group':'INNER',
        'cell':'BQ',
        'qtyUnit':'BU',
        'qty':'BV',
        'remark':'BX',
        'weight':'',
        'FS':'CB',
        'color':'BT',
        'length':'',
        'description':'BN',
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
    workBook.addNewSheet()
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
            
        def ReplacementFromLengthColom(value):
            operasiMath = ['+','-']
            if(value):
                hasOprasiMath = False
                for valOperasiMath in operasiMath:
                    if value.find(valOperasiMath) != -1:
                        hasOprasiMath = True

                if hasOprasiMath:
                    return value.replace("=","'=(") + ")"
                else:
                    return "'"+value
            else:
                return value
            
        def ReplacementFromulaCellColom(BomID,FormulaCode):
            dbShiage = shiage()
            getDataFromulaCell = dbShiage.GetMadelaFormulaBom(
                strWhere=f" BOM_ID ='{BomID}' AND  Type='P' AND FormulaCode='{FormulaCode}'"
            )
            if len(getDataFromulaCell) >= 1:
                strText = f"'={getDataFromulaCell[0].Formula}"
                return strText
            else:
                return FormulaCode
            
        def ReplacementColorCellPartPageColom(BomID,group,partno,FS,color):
            strWhere= f" BOM_ID ='{BomID}' AND Cls_IOM='{group}' AND PartNo='{partno}' "

            if FS :
                strWhere += f" AND FS='S'"
            else:
                strWhere += f" AND FS=''"
            dbShiage = shiage()
            getDataMadelaPartCell = dbShiage.GetMadelaPartBom(
                strWhere=strWhere
            )
            if len(getDataMadelaPartCell) == 1:
                strText = f"{getDataMadelaPartCell[0].Colour}"
                return strText
            else:
                return color
            
        def ReplacementHoldCellPartColom(ArrayColom):
            result = []
            for valArrayColom in ArrayColom:
                if(ArrayColom.index(valArrayColom) == 0):
                    for valueValArrayColom in valArrayColom:
                        valueWorkBook = valueValArrayColom['value']
                        if valueValArrayColom['cellColom'] == valueValArrayColom['qty']:
                            valueWorkBook = "'"+valueValArrayColom['value']
                        workBook.sheet.Range[f'{valueValArrayColom["cell"]}'].Text = valueWorkBook
                else:
                    result.append(valArrayColom)
            return result
                  
        # start
        if not sheetpage == {}:
            # fabrication
            emptyColsFabrication = []
            toFillEmptyColsFabrication = []
            for valueColomCell in sheetpage['fabricationColom']:
                for valueRowCell in range(sheetpage['fabricationRowFrom'],sheetpage['fabricationRowTo']):
                    if not workBook.sheet.Range[f'{valueColomCell['cell']}{valueRowCell}'].Text:
                        emptyColsFabrication.append(valueRowCell)
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
                        # length
                        if workBook.sheet.Range[f'{valueColomCell['length']}{valueRowCell}'].Text:
                            workBook.editText(
                                cell=f"{valueColomCell['length']}{valueRowCell}",
                                text=ReplacementFromLengthColom(
                                    value=workBook.sheet.Range[f'{valueColomCell['length']}{valueRowCell}'].Text
                                )
                            )
            # part
            emptyColsPart = []
            toFillEmptyColsPart = []
            for valueColomCell in sheetpage['partColom']:
                for valueRowCell in range(sheetpage['partRowFrom'],sheetpage['partRowTo']):
                    if not workBook.sheet.Range[f'{valueColomCell['cell']}{valueRowCell}'].Text:
                        emptyColsPart.append(valueRowCell)
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
                        # Color
                        if (workBook.sheet.Range[f'{valueColomCell['cell']}{valueRowCell}'].Text and
                            workBook.sheet.Range[f'{valueColomCell['cell']}{valueRowCell}'].Text.find('P-') == -1 and 
                            workBook.sheet.Range[f'{valueColomCell['cell']}{valueRowCell}'].Text.find('=') == -1 and
                            workBook.sheet.Range[f'{valueColomCell['color']}{valueRowCell}'].Text
                            ):
                            workBook.editText(
                                cell=f"{valueColomCell['color']}{valueRowCell}",
                                text=ReplacementColorCellPartPageColom(
                                    BomID=xlsxName,
                                    group=valueColomCell['group'],
                                    partno=workBook.sheet.Range[f'{valueColomCell['cell']}{valueRowCell}'].Text,
                                    FS=workBook.sheet.Range[f'{valueColomCell['FS']}{valueRowCell}'].Text,
                                    color=workBook.sheet.Range[f'{valueColomCell['color']}{valueRowCell}'].Text
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
                        # Hold Cap Split function
                        if workBook.sheet.Range[f'{valueColomCell['description']}{valueRowCell}'].Text.find('HOLE') != -1:
                            splitArrayCol = [] 
                            qtyUnitTxtArray = workBook.sheet.Range[f'{valueColomCell['qtyUnit']}{valueRowCell}'].Text.split('+IF')
                            
                            if(len(qtyUnitTxtArray) > 1):
                                for valueQtyUnitTxt in qtyUnitTxtArray:
                                    arrayCol = []
                                    if not valueQtyUnitTxt.find("=") : 
                                        qtyUnitTxt = "'"+valueQtyUnitTxt 
                                    else: 
                                        qtyUnitTxt = "'=IF"+valueQtyUnitTxt

                                    for valueListCell in valueColomCell['listCell']:
                                        if(valueListCell == valueColomCell['qtyUnit']):
                                            valueArray = qtyUnitTxt
                                        else:
                                            valueArray = workBook.sheet.Range[f'{valueListCell}{valueRowCell}'].Text

                                            remarkTxtArray = workBook.sheet.Range[f'{valueColomCell['remark']}{valueRowCell}'].Text.split(', ')
                                            if (len(remarkTxtArray) == len(qtyUnitTxtArray) and valueListCell == valueColomCell['remark']):
                                                valueArray=remarkTxtArray[qtyUnitTxtArray.index(valueQtyUnitTxt)]
                                       

                                        arrayCol.append({
                                            'cell':f'{valueListCell}{valueRowCell}',
                                            'cellColom':f'{valueListCell}',
                                            'qtyUnit':valueColomCell['qtyUnit'],
                                            'qty':valueColomCell['qty'],
                                            'value':valueArray
                                        })

                                    if len(arrayCol):
                                        splitArrayCol.append(arrayCol)
                                
                                toFillEmptyColsPart.extend(ReplacementHoldCellPartColom(splitArrayCol))
            # add empty row
            if len(toFillEmptyColsPart):
                for valToFillEmptyColsPart in toFillEmptyColsPart:
                    row = emptyColsPart[0]
                    emptyColsPart = emptyColsPart[1:]
                    for valValToFillEmptyColsPart in valToFillEmptyColsPart:
                        valueWorkBook = valValToFillEmptyColsPart['value']
                        if valValToFillEmptyColsPart['cellColom'] == valValToFillEmptyColsPart['qty']:
                            valueWorkBook = "'"+f'=IF({valValToFillEmptyColsPart['qtyUnit']}{row}="",""' + f',(Q*{valValToFillEmptyColsPart['qtyUnit']}{row}))'
                        workBook.sheet.Range[f'{valValToFillEmptyColsPart['cellColom']}{row}'].Text = valueWorkBook


    def ColomReplaceByTypeProduct():
        def ReplaceHead(replaceList):
            for ValueReplaceList in replaceList:
                if 'cell' in ValueReplaceList.keys():
                    workBook.editText(cell=ValueReplaceList['cell'],text=ValueReplaceList['text'])

        def ReplacementMaterialInner(replaceList):
            for ValueReplaceList in replaceList:
                if 'row' in ValueReplaceList.keys():
                    for value in range(typeSheetPage5['fabricationRowFrom'],typeSheetPage5['fabricationRowTo']):
                       if(workBook.sheet.Range[f'{typeSheetPage5['fabricationColom'][1]['description']}{value}'].Text):
                            workBook.sheet.Range[f'{ValueReplaceList['rowtext']}{value}'].Text = f'{ValueReplaceList['text']}'

        def ReplacementPartInner(replaceList):
            for ValueReplaceList in replaceList:
                if 'rowValue' in ValueReplaceList.keys():
                    countnumber = 0
                    for value in range(typeSheetPage5['partRowFrom'],typeSheetPage5['partRowTo']):
                        if(workBook.sheet.Range[f'{typeSheetPage5['partColom'][1]['description']}{value}'].Text == ValueReplaceList['value']):
                            if 'num' in ValueReplaceList.keys():
                                countnumber+=1
                                if ValueReplaceList['num'] == countnumber:
                                    workBook.sheet.Range[f'{ValueReplaceList['rowtext']}{value}'].Text = ValueReplaceList['text']
                            else:
                                workBook.sheet.Range[f'{ValueReplaceList['rowtext']}{value}'].Text = ValueReplaceList['text']

        for ValueProductType in ReplaceByTypeProduct:
            if(xlsxName in ValueProductType['bom']):
                # ReplaceHead
                ReplaceHead(ValueProductType['replacementHead'])
                ReplacementMaterialInner(ValueProductType['replacementMaterialInner'])
                ReplacementPartInner(ValueProductType['replacementPartInner'])
                break;               
                    
    ColomFromula()
    ColomReplace()
    ColomEmptyANDFixFromula()
    ColomReplaceByTypeProduct()

    workBook.saveFile(pathToSave)
    prosesFormula.endReadProses()


for valueXlsx in listXlsx:
    mainRun(valueXlsx)