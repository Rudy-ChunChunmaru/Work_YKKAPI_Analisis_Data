import pyodbc 

class Database_Shiage:
    def __init__(self):
        self.SERVER = '10.246.182.54'
        self.DATABASE = 'YKK_AP'
        self.USERNAME = 'sa'
        self.PASSWORD = 'p@ssw0rd'
        self.connectionString = f'DRIVER={{SQL Server}};SERVER={self.SERVER};DATABASE={self.DATABASE};UID={self.USERNAME};PWD={self.PASSWORD}'
        self.conn = pyodbc.connect(self.connectionString)
        self.cursor = self.conn.cursor()

    def GetMadelaFormulaBom(self,strWhere=None,strOrderBy=None):
        SQL_QUERY = """
        SELECT 
        * 
        from 
        YKK_AP.dbo.MADELA_MASTER_FORMULA 
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
    
    def GetMadelaPartBom(self,strWhere=None,strOrderBy=None):
        SQL_QUERY = """
        SELECT 
        * 
        from 
        YKK_AP.dbo.MADELA_PARTLIST_MSTDTL_BOM 
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
