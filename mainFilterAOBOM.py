import pandas as pd
from template.read_proses import CopyWorkBook as cwb
from database.db_shiage import Database_Shiage as shiage

dfBOM = pd.read_excel('data/AO_and_BOM/BOM.xlsx')
dfAO = pd.read_excel('data/AO_and_BOM/AO.xlsx')

def BomFilter(BOM_ID):
    return dfBOM[(dfBOM.BOM_ID == BOM_ID)]

def BomGetColom(dfBom,col):
    return {
        'ProductCode' : str(dfBom.iloc[col]['ProductCode']),
        'BOM_ID' : int(dfBom.iloc[col]['BOM_ID']),
        'Spec1' : str(dfBom.iloc[col]['Spec1']),
        'Spec2' : str(dfBom.iloc[col]['Spec2']),
        'Spec3' : str(dfBom.iloc[col]['Spec3']),
        'Spec4' : str(dfBom.iloc[col]['Spec4']),
        'Spec5' : str(dfBom.iloc[col]['Spec5']),
    }

def AOFilter(colomBom):
    resultFilter_dfAO = dfAO
    if colomBom['ProductCode'] != 'nan':
        resultFilter_dfAO = resultFilter_dfAO[
            (resultFilter_dfAO['Product Type Madela'] == colomBom['ProductCode'])
        ]
    if('door' in colomBom['ProductCode'].lower()):
        if colomBom['Spec1'] != 'nan':
            resultFilter_dfAO = resultFilter_dfAO[
                (resultFilter_dfAO['Spec 2'] == colomBom['Spec1'])
            ]

        if colomBom['Spec2'] != 'nan':
            resultFilter_dfAO = resultFilter_dfAO[
                (resultFilter_dfAO['Spec 3'] == colomBom['Spec2'])
            ]

        if colomBom['Spec3'] != 'nan':
            resultFilter_dfAO = resultFilter_dfAO[
                (resultFilter_dfAO['Spec 4'] == colomBom['Spec3'])
            ]
    elif('t-ml' in colomBom['ProductCode'].lower()):
        if colomBom['Spec3'] != 'nan':
            resultFilter_dfAO = resultFilter_dfAO[
                (resultFilter_dfAO['Spec 4'] == colomBom['Spec3'].upper())
            ]
        else:
            resultFilter_dfAO = resultFilter_dfAO[
                (resultFilter_dfAO['Spec 4'] == 'none'.upper())
            ]
    else:
        if colomBom['Spec1'] != 'nan':
            resultFilter_dfAO = resultFilter_dfAO[
                (resultFilter_dfAO['Spec 2'] == colomBom['Spec1'])
            ]

        if colomBom['Spec2'] != 'nan':
            resultFilter_dfAO = resultFilter_dfAO[
                (resultFilter_dfAO['Option 2'] == colomBom['Spec2'].upper())
            ]
        else:
            resultFilter_dfAO = resultFilter_dfAO[
                (resultFilter_dfAO['Option 2'] == 'none'.upper())
            ]

        if colomBom['Spec3'] != 'nan':
            resultFilter_dfAO = resultFilter_dfAO[
                (resultFilter_dfAO['Option 3'] == colomBom['Spec3'].upper())
            ]
        else:
            resultFilter_dfAO = resultFilter_dfAO[
                (resultFilter_dfAO['Option 3'] == 'none'.upper())
            ]


    resultFilter_dfAO.sort_values(by='Description', ascending = False, inplace = True) 
    return resultFilter_dfAO

def getDataFromDataBase(Order_No,Item_No):
    dbShiage = shiage(database='YKK_AP')
    strSelectManufactruing = '''
        Pageno,
        ItemUnit,
        Unit_Code,
        Cls_IOM,
        Description,
        FabricationNo,
        Color,
        FabNo,
        Length,
        QTYUnit,
        QTY,
        Remark,
        Remark2,
        Remark3,
        '' AS FS,
        Weight_m,
        Weight_kg
    '''
    strSelectPart = '''
        Pageno,
		ItemUnit,
		Unit_Code,
		Cls_IOM,
		Description,
		Colour,
		PartNo,
		'' as Length,
		QTYUnit,
		QTY,
		Remark,
		Remark2,
		Remark3,
		Remark4,
		Remark5,
		FS
    '''
    status = True
    ManufactruingList = dbShiage.GetManufactruingListDetail(
        strSelect=strSelectManufactruing,
        strWhere=f" Order_No ='{Order_No}' AND Item_No='{Item_No}'",
        strOrderBy=f" FabricationNo ASC,Unit_Code ASC,Length ASC;"
    )
    PartList = dbShiage.GetPartListDetail(
        strSelect=strSelectPart,
        strWhere=f" Order_No ='{Order_No}' AND Item_No='{Item_No}'",
        strOrderBy=f" PartNo ASC,Unit_Code ASC,QTYUnit ASC;"
    )
    if(len(ManufactruingList) < 1 and len(PartList) < 1):
        status = False
    return {
        "status":status,
        "ManufactruingList": ManufactruingList,
        "PartList": PartList
    }

excelColom = {
    "shiageManufactruing":{
        "colom":['S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI'],
        "rowStart":5,
        "rowEnd":31
    },
    "shiagePart":{
        "colom":['S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH'],
        "rowStart":36,
        "rowEnd":200
    }
}

# TODO-1 Start Proses:
workBook = cwb(pathToGet='./data/AO_and_BOM/',xlsxName='template')
while True: 
    Bom_ID = input('masukan Bom ID \n BOM_ID=')
    if Bom_ID == '':
        break
    try:
        dataBom = BomFilter(int(Bom_ID))
    except:
        print('Error Bom ID not integer number !!!\n')
        continue
    else:
        if len(dataBom.index) == 0:
            print('Bom ID not in BOM xlsx data !!!\n')
            continue
        else:
            objBom = BomGetColom(dataBom,0)
            dataAO = AOFilter(objBom)
            if len(dataAO.index) == 0:
                print('Bom ID not in AO xlsx data !!!\n')
                continue
            else:
                print(dataAO.groupby(['Line No.'])['Line No.'].count())
                while True:
                    AONumber = input('pilih nomor AO yang mau di pakai \n NumberAO=')
                    if (AONumber == ''):
                        break
                    else:
                        result_dfAO = dataAO[(dataAO['Line No.'] == AONumber)]
                        if len(result_dfAO.index) == 0:
                            print('Pilihan AO Salah !!!\n')
                            continue
                        else:
                            break
                    
                if(AONumber == ''):
                    continue
                else:
                    result_dfAO = dataAO[(dataAO['Line No.'] == AONumber)]
                    if len(result_dfAO.index) == 0:
                        print('Pilihan AO Salah !!!\n')
                        continue
                    else:
                        Order_No = result_dfAO.iloc[0]['Cost/Pcs']
                        Item_No = result_dfAO.iloc[0]['Order Seq No.']
                        resultBD = getDataFromDataBase(Order_No=Order_No,Item_No=Item_No)
                        if resultBD['status'] == False:
                            print('selected data error !!!\n')
                            continue
                        else:
                            print(f'{objBom['BOM_ID']}_{objBom['ProductCode'].replace('/','_')}')
                            workBook.addNewSheet(copySheet=workBook.deSheet,namesheet=f'{objBom['BOM_ID']}_{objBom['ProductCode'].replace('/','_')}')
                            workBook.sheet.Range[f'A1'].Text = AONumber
                            workBook.sheet.Range[f'B1'].Text = Order_No
                            workBook.sheet.Range[f'C1'].Text = Item_No
                            ManufactruingList = resultBD['ManufactruingList']
                            PartList = resultBD['PartList']
                            for row in range(excelColom["shiageManufactruing"]["rowStart"],excelColom["shiageManufactruing"]["rowEnd"]):
                                if len(ManufactruingList) > 0:
                                    Manufactruing = ManufactruingList[0]
                                    ManufactruingList = ManufactruingList[1:]
                                    ColomNumber = ['AA','AB','AC']
                                    for col in excelColom["shiageManufactruing"]["colom"]:
                                        value =  Manufactruing[0]
                                        Manufactruing = Manufactruing[1:]
                                        if col not in ColomNumber:
                                            workBook.sheet.Range[f'{col}{row}'].Text = str(value).strip()
                                        else:
                                            workBook.sheet.Range[f'{col}{row}'].Number = str(value).strip()
                            for row in range(excelColom["shiagePart"]["rowStart"],excelColom["shiagePart"]["rowEnd"]):
                                if len(PartList) > 0:
                                    Part = PartList[0]
                                    PartList = PartList[1:]
                                    ColomNumber = ['AA','AB']
                                    for col in excelColom["shiagePart"]["colom"]:
                                        value =  Part[0]
                                        Part = Part[1:]
                                        if col not in ColomNumber:
                                            workBook.sheet.Range[f'{col}{row}'].Text = str(value).strip()
                                        else:
                                            workBook.sheet.Range[f'{col}{row}'].Number = str(value).strip()
                        
                         
workBook.saveFile('./data/proses_AO_and_BOM/')