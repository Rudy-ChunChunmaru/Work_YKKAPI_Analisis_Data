import os.path
import pandas as pd

from datetime import datetime

from database.db_shiage import Database_Shiage as shiage

DB='YKK_AP_DEV1'

def printTitle():
    print(chr(27) + "[2J" + "\n")
    print(chr(27) + "[2J" + "\n")
    print(chr(27) + "[2J" + "\n")
    print("""
        ___  __            _____    ____    _____  _    _   _____  _____ 
        |  \/  |    /\    |  __ \  / __ \  / ____|| |  | | / ____||_   _|
        | \  / |   /  \   | |  | || |  | || |  __ | |  | || |       | |  
        | |\/| |  / /\ \  | |  | || |  | || | |_ || |  | || |       | |  
        | |  | | / ____ \ | |__| || |__| || |__| || |__| || |____  _| |_ 
        |_|  |_|/_/    \_\|_____/  \____/  \_____| \____/  \_____||_____|
    """)

def getMadoguchiMaster(no_project):
    dbShiage = shiage(database=DB)
    strWhere = f"PROJECT_NO = '{no_project}'"
    data = dbShiage.GetMadoguchiMaster(strSelect='*',strWhere=strWhere)
    dbShiage.cursor.close()
    return data
   
def MenuProses():
    printTitle()
    print('''
    PROSES DATABASE SINKRON SHIAGE BC
    1. GET MADOGUCHI
    2. UPLOAD MADOGUCHI
    ''')
    PDSSB = input('number proses data ?\n')
    if(PDSSB == '1'):
        print(getMadoguchiMaster(no_project='9898989'))
        return True
    elif(PDSSB == '2'):
        UploadMadoguchi()
        return True
    else:
        if(PDSSB == ''):
            return False
        else:
            return True


def GetMadoguchi():
    pass

def UploadMadoguchi():
    def UploadValidasi(no_project,file_location):
        status = True
        if getMadoguchiMaster(no_project=no_project) != []:
            status = False
            print('Error no project already use !!!')

        if os.path.exists(file_location) == False :
            status = False
            print('Error cant find file directroy !!! \n')
        
        return status
    
    def ProsesXlsx(no_project,file_location):
        dbShiage = shiage(database=DB)
        df = pd.read_csv(file_location)
        status = True
        # Data Input
        order_no = df["InternalOrderNo"].values[0]
        project_name =  df["ProjectName"].values[0]
        total_qty = df.groupby('InternalOrderNo').agg({'Qty':'sum', 'AlumuniumWeight': 'sum'}).reset_index()["Qty"].values[0]
        total_weight = df.groupby('InternalOrderNo').agg({'Qty':'sum', 'AlumuniumWeight': 'sum'}).reset_index()["AlumuniumWeight"].values[0]
        dataMadoguchiMaster = [
            f"N'{no_project}'",
            f"N'{order_no}'",
            f"'{datetime.now().date()}'",
            f"'{datetime.now().date()}'",
            "'1900-01-01 00:00:00.000'",
            f"'{datetime.now().date()}'",
            f"N'{project_name}'",
            "N'ONLINE'",
            "N'ONLINE'",
            "N'605401'",
            "N'PT ykk'",
            "N'605401'",
            "N'PT. ykk'",
            "N''",
            "N''",
            "N''",
            "N'D'",
            f"{total_qty}",
            f"{total_weight}",
            "N''",
            "N''",
            "'1900-01-01 00:00:00.000'",
            "N''",
            "N''",
            "N''",
            "N'Tangerang-MDL'",
            "N'TANGERANG'",
            "N'Tangerang'",
            "N''",
            "N''",
            "N''",
            "N''",
            "N'PT. YKK AP INDONESIA'",
            "N'Tangerang'",
            "N'ID'",
            "N''",
            "N''",
            "'1900-01-01 00:00:00.000'",
            "'1900-01-01 00:00:00.000'",
            "'1900-01-01 00:00:00.000'",
            "0",
            "N'1'",
            "'1900-01-01 00:00:00.000'",
        ]
        status = dbShiage.InsertMadoguchiMaster(dataMadoguchiMaster)
        if status == False:
            print('Error no uploading InsertMadoguchiMaster !!!')
            
        for index in range(0,len(df.index)):
            dataMadoguchiDetail=[
                f"N'{no_project}'",
                f"N'{order_no}'",
                f"{index+1}",
                "1",
                f"N'{df['ItemNo'].values[index]}'",
                "N'AK'",
                "N'MADELA'",
                "N''",
                "N'1'",
                f"N'{df['ColorCode'].values[index]}'",
                "N''",
                f"{df['Width'].values[index]}",
                f"{df['Height'].values[index]}",
                f"{df['AlumuniumWeight'].values[index]}",
                "N'SET'",
                f"{df['Qty'].values[index]}",
                f"{df['Qty'].values[index]}",
                f"{df['Qty'].values[index]}",
                "N'0'",
                "N''",
                "N'A'",
                "N'1'",
                "N'M2'",
                "N'MA'",
                f"'{datetime.now().date()}'",
                "N''",
                "0",
                "N''",
                "N''",
                "0",
                "N'interface'",
                "'1900-01-01 00:00:00.000'",
                "'1900-01-01 00:00:00.000'",
                "'1900-01-01 00:00:00.000'",
                "0",
                "0",
                "0",
                "0",
                "0",
                "0",
                "0",
                "N'2'",
                "N''"
            ]
            status = dbShiage.InsertMadoguchiDetail(dataMadoguchiDetail)
            if status == False:
                print('Error no uploading InsertMadoguchiDetail !!!')
                break
                

            Spec_1=''
            Spec_2=''
            Spec_3=''
            Spec_4=''
            Spec_5=''
            Ops_1='None'
            Ops_2='None'
            Ops_3='None'
            Ops_4='None'
            Ops_5='None'
            if 'door' in df['ProductCode'].values[index].lower() :
                Spec_1=df['Spec_1'].values[index]
                Spec_2=df['Spec_2'].values[index]
                Spec_3=df['Spec_3'].values[index]
            elif 'T-ML' in df['ProductCode'].values[index].lower():
                Spec_3=df['Spec_3'].values[index]
            else:
                Spec_1=df['Spec_1'].values[index]
                Spec_2=df['Option_1'].values[index]
                Spec_3=df['Option_2'].values[index].capitalize()
                
            
            HandlePosition=df['Width'].values[index]
            if(HandlePosition == 'nan'):
                HandlePosition=''
            
            dataMadelaOrderWindowDetail =[
                f"'{datetime.now().date()}'",
                f"N'{order_no}'",
                f"{index+1}",
                f"N'{no_project}'",
                f"N'{project_name}'",
                "N'605401'",
                "N'D01023'",
                f"'{datetime.now().date()}'",
                f"N'{df['ItemNo'].values[index]}'",
                "N'JS03'",
                f"N'{df['ColorCode'].values[index]}'",
                f"{df['Width'].values[index]}",
                f"{df['Height'].values[index]}",
                f"{df['Qty'].values[index]}",
                f"{df['AlumuniumWeight'].values[index]}",
                f"{df['PuchaseUnitPrice'].values[index]}",
                f"{df['PuchaseUnitPrice'].values[index] * df['Qty'].values[index]}",
                f"N'{df['ProductCode'].values[index]}'",
                f"N'{df['ProductType'].values[index]}'",
                f"N'{df['Pressure'].values[index]}'",
                f"N'{df['GlassThickness'].values[index]}'",
                f"{df['h_1'].values[index]}",
                f"{df['h_2'].values[index]}",
                f"{df['h_3'].values[index]}",
                f"{df['h_4'].values[index]}",
                f"{df['H_1'].values[index]}",
                f"{df['H_2'].values[index]}",
                f"{df['H_3'].values[index]}",
                f"{df['H_4'].values[index]}",
                f"N'{Spec_1}'",
                f"N'{Spec_2}'",
                f"N'{Spec_3}'",
                f"N'{Spec_4}'",
                f"N'{Spec_5}'",
                f"N'{Ops_1}'",
                f"N'{Ops_2}'",
                f"N'{Ops_3}'",
                f"N'{Ops_4}'",
                f"N'{Ops_5}'",
                f"{df['G1_W'].values[index]}",
                f"{df['G1_H'].values[index]}",
                f"{df['G2_W'].values[index]}",
                f"{df['G2_H'].values[index]}",
                f"{df['G3_W'].values[index]}",
                f"{df['G3_H'].values[index]}",
                f"{df['G4_W'].values[index]}",
                f"{df['G4_H'].values[index]}",
                f"N'{HandlePosition}'",
                f"{df['HandleHeight1'].values[index]}",
                f"{df['HandleHeight2'].values[index]}",
                f"'{datetime.now().date()}'",
            ]
            status = dbShiage.InsertMadelaOrderWindowDetail(dataMadelaOrderWindowDetail)
            if status == False:
                print('Error no uploading InsertMadelaOrderWindowDetail !!!')
                break

        if(status):
            dbShiage.cursor.commit()
        else:
            dbShiage.cursor.rollback()

        return status

    while True:
        printTitle()
        no_project = input('add new no project = ')
        file_location = input('file location xlsx file =')
        if(no_project=='' or file_location==''):
            break
        else:
            if UploadValidasi(no_project,file_location):
                if ProsesXlsx(no_project,file_location):
                    break

while True:
    if MenuProses():
        continue
    else:
        break

