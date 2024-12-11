import pyodbc 

class Database_Shiage:
    def __init__(self,database='YKK_AP'):
        self.SERVER = '10.246.182.54'
        self.DATABASE = database
        self.USERNAME = 'sa'
        self.PASSWORD = 'p@ssw0rd'

        self.connectionString = f'DRIVER={{SQL Server}};SERVER={self.SERVER};DATABASE={self.DATABASE};UID={self.USERNAME};PWD={self.PASSWORD}'
        self.conn = pyodbc.connect(self.connectionString)
        self.cursor = self.conn.cursor()

    # TODO: get BOM
    def GetMadelaManufactruingBom(self,strSelect='*',strWhere=None,strOrderBy=None):
        SQL_QUERY = f"""
        SELECT 
        {strSelect}
        from 
        MADELA_MANUFACTURINGLIST_MSTDTL_BOM 
        """

        if(strWhere):
            SQL_QUERY += f"""
                where 
                {strWhere}
            """

        if(strOrderBy):
            SQL_QUERY += f"""
                order by 
                {strOrderBy}
            """

        self.cursor.execute(SQL_QUERY)
        return self.cursor.fetchall()
    
    def GetMadelaPartBom(self,strSelect='*',strWhere=None,strOrderBy=None):
        SQL_QUERY = f"""
        SELECT 
        {strSelect}
        from 
        MADELA_PARTLIST_MSTDTL_BOM 
        """

        if(strWhere):
            SQL_QUERY += f"""
                where 
                {strWhere}
            """

        if(strOrderBy):
            SQL_QUERY += f"""
                order by 
                {strOrderBy}
            """

        self.cursor.execute(SQL_QUERY)
        return self.cursor.fetchall()
    
    def GetMadelaFormulaBom(self,strSelect='*',strWhere=None,strOrderBy=None):
        SQL_QUERY = f"""
        SELECT 
        {strSelect} 
        from 
        MADELA_MASTER_FORMULA 
        """

        if(strWhere):
            SQL_QUERY += f"""
                where 
                {strWhere}
            """

        if(strOrderBy):
            SQL_QUERY += f"""
                order by 
                {strOrderBy}
            """

        self.cursor.execute(SQL_QUERY)
        return self.cursor.fetchall()
    
    # TODO: get Order
    def GetMadoguchiMaster(self,strSelect='*',strWhere=None,strOrderBy=None):
        SQL_QUERY = f"""
        SELECT 
        {strSelect} 
        from 
        MADOGUCHI_MASTER
        """

        if(strWhere):
            SQL_QUERY += f"""
                where 
                {strWhere}
            """

        if(strOrderBy):
            SQL_QUERY += f"""
                order by 
                {strOrderBy}
            """

        self.cursor.execute(SQL_QUERY)
        return self.cursor.fetchall()
    
    def GetMadoguchiDetail(self,strSelect='*',strWhere=None,strOrderBy=None):
        SQL_QUERY = f"""
        SELECT 
        {strSelect} 
        from 
        MADOGUCHI_DETAIL
        """

        if(strWhere):
            SQL_QUERY += f"""
                where 
                {strWhere}
            """

        if(strOrderBy):
            SQL_QUERY += f"""
                order by 
                {strOrderBy}
            """

        self.cursor.execute(SQL_QUERY)
        return self.cursor.fetchall()
    
    def GetMadelaOrderWindowDetail(self,strSelect='*',strWhere=None,strOrderBy=None):
        SQL_QUERY = f"""
        SELECT 
        {strSelect} 
        from 
        MADELA_ORDER_WINDOW_DETAIL
        """

        if(strWhere):
            SQL_QUERY += f"""
                where 
                {strWhere}
            """

        if(strOrderBy):
            SQL_QUERY += f"""
                order by 
                {strOrderBy}
            """

        self.cursor.execute(SQL_QUERY)
        return self.cursor.fetchall()
    
    # TODO: get Manufactring List and Part List
    def GetManufactruingListDetail(self,strSelect='*',strWhere=None,strOrderBy=None):
        SQL_QUERY = f"""
        SELECT 
        {strSelect}
        from 
        MANUFACTURINGLIST_MSTDTL 
        """

        if(strWhere):
            SQL_QUERY += f"""
                where 
                {strWhere}
            """

        if(strOrderBy):
            SQL_QUERY += f"""
                order by 
                {strOrderBy}
            """

        self.cursor.execute(SQL_QUERY)
        return self.cursor.fetchall()
    
    def GetPartListDetail(self,strSelect='*',strWhere=None,strOrderBy=None):
        SQL_QUERY = f"""
        SELECT 
        {strSelect}
        from 
        PARTLIST_MSTDTL 
        """

        if(strWhere):
            SQL_QUERY += f"""
                where 
                {strWhere}
            """

        if(strOrderBy):
            SQL_QUERY += f"""
                order by 
                {strOrderBy}
            """

        self.cursor.execute(SQL_QUERY)
        return self.cursor.fetchall()
    
    # TODO : Insert Madoguchi Master
    def InsertMadoguchiMaster(self,data):
        status = True
        SQL_QUERY = f"""
            INSERT INTO MADOGUCHI_MASTER (
                PROJECT_NO,
                ORDER_NO,
                ORDER_DATE,
                ORDER_RECEIVED_DATE,
                EXPECTED_DELIVERY_DATE,
                ESTIMATED_DELIVERY_DATE,
                PROJECT_NAME,
                SALES_PIC_CODE,
                SALES_PIC_NAME,
                CUSTOMER_CODE,
                CUSTOMER_NAME,
                DESTINATION_CODE,
                DESTINATION_NAME,
                REMARKS_1,
                REMARKS_2,
                STOP_DELIVERY_CLS,
                OVERSEAS_CLS,
                TOTAL_ITEM_NO_QTY,
                TOTAL_ITEM_NO_WEIGHT,
                MAP_CLS,
                BREAKDOWN_FINISH_CLS,
                BREAKDOWN_UPLOAD_DATE,
                BREAKDOWN_PIC,
                SENT_TO_BEONE_CLS,
                PRJ_CLSS,ORD_PLC_BY_CD,
                PRJ_LCAT_ADM_CD,PRJ_LCAT,
                OPRT_DPT_DV,
                ACTCTR_APPL_CD,
                KNCK_DWN_SHP_DV,
                ARNG_TO_CD,
                DRCT_SHP_NM,
                DRCT_SHP_ADR,
                COUNTRY,
                BEONE_PROJECT_NO,
                USERID,
                REGISTER_DATE,
                LAST_UPDATE,
                IF_CLS_BREAKDOWN_FINISH_STATUS,
                COMPLETE_CLS,
                ORDER_TYPE,
                SENT_TO_BEONE_DATE
            ) VALUES ({
                ', '.join(data)
            });
        """
        try:
            self.cursor.execute(SQL_QUERY)
        except:
            status = False

        return status
    
    # TODO : Insert Madoguchi Master
    def InsertMadoguchiDetail(self,data):
        status = True
        SQL_QUERY = f"""
            INSERT INTO MADOGUCHI_DETAIL (
                PROJECT_NO,	
                ORDER_NO,
                ORDER_SEQ_NO,	
                ORDER_LOT_NO,	
                ITEM_NO,
                SERIES_CODE,	
                SERIES_NAME,
                WIND_TYPE_CODE,	
                WINDOW_VARIATION_CODE,	
                COLOR_CODE,
                THICKNESS_CODE,
                WIDTH,	
                HEIGHT,	
                WEIGHT_PER_UNIT,
                UNIT_QTY_CODE,	
                ITEM_NO_QTY,	
                ITEM_NO_PICKING_QTY,
                ITEM_NO_DELIVERY_QTY,	
                CANCEL_CLS,
                ELEVATION_CODE,	
                BUILDING_BLOCK,
                BUILDING_FLOOR,	
                PRODUCTION_PATTERN_CODE,
                PRODUCTION_LINE_BEFORE_BREAKDOWN_CODE,
                ESTIMATED_DELIVERY_DATE,
                PRD_INPT_DV,
                FILE_ID,
                MAP_PROJECT_NO,
                MAP_ORDER_NO,
                MAP_ORDER_SEQ_NO,
                USERID,
                REGISTER_DATE,	
                LAST_UPDATE,
                IF_CLS_ESTIMATED_DELIVERY_DATE,
                ITEM_NO_STOCK_OVERALL_SET,
                COST1,
                COST2,
                COST3,
                COST4,
                COST5,	
                COST6,	
                CHRG_NON_CHRG_DV,
                PLC
            ) VALUES ({
                ', '.join(data)
            });
        """
        try:
            self.cursor.execute(SQL_QUERY)
        except:
            print(SQL_QUERY)
            status = False

        return status
    
    # TODO : Insert Madoguchi Master
    def InsertMadelaOrderWindowDetail(self,data):
        status = True
        SQL_QUERY = f"""
            INSERT INTO MADELA_ORDER_WINDOW_DETAIL (
                OrderDate,
                OrderNo,
                OrderSeqNo,
                ProjectCode,
                ProjectName,
                CustomerCode,
                DestCode,
                ExpectedDate,
                ItemNo,
                CustProductCode,
                ColorCode,
                Width,
                height,
                Qty,
                WeightTotal,
                SalesPricePerQty,
                TotalSalesAmount,
                ProductCode,
                ProductType,
                Preasure,
                GlassThickness,
                h1,
                h2,
                h3,
                h4,
                Height_1,
                Height_2,
                Height_3,
                Height_4,
                Spec1,
                Spec2,
                Spec3,
                Spec4,
                Spec5,
                Option1,
                Option2,
                Option3,
                Option4,
                Option5,
                G1_W,
                G1_H,
                G2_W,
                G2_H,
                G3_W,
                G3_H,
                G4_W,
                G4_H,
                HandlePosition,
                HandleHeight1,
                HandleHeight2,
                FaxDate
            ) VALUES ({
                ', '.join(data)
            });
        """
        try:
            self.cursor.execute(SQL_QUERY)
        except:
            print(SQL_QUERY)
            status = False

        return status

        
            

        

    
    

    
