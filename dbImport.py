import re
import datetime

class dbImport:
    def __init__(self, settingFile):
        import MyPkg.sqlDataType
        import MyPkg.MyFile as MyFile
        import psycopg2
        from sqlalchemy import create_engine
        
        
        self.xx = MyFile.FileSystem()
        self.Types = MyPkg.sqlDataType.sqlDataType()
        self.Setting = self.parseSetting(settingFile)
        self.connection = psycopg2.connect(self.Setting["DATABASE_URI"])
        self.cur = self.connection.cursor()
        self.showSetting(self.Setting)
        self.engine = create_engine(self.Setting["DATABASE_ENGINE_URI"])
    
        if not("SOURCE" in self.Setting):
            print ("There is missing source directory in .env")
            self.Setting["Source"] = "Source"
            
                    
        
        
    def showSetting(self, Setting):
        print ("Setting parameters")
        for i in self.Setting:
            print(i," = ", self.Setting[i])
        return

    def parseSetting(self, Source):
        regex = re.compile(b"[A-Za-z_]+\=.*")
        reValue = re.compile(b'\".*\"')
        result = {}
        for i in regex.findall(Source):
            result [i.split(b'=')[0].decode()] = reValue.findall(i)[0].decode()[1:-1]
        return result


    def execute_query(self, query, *args, fetch_rows=0, **kwargs):
        """Execute SQL query, set fetch_rows=-1 to return all available rows"""
        with self.connection as connection:
            with connection.cursor() as cursor:
                cursor.execute(query, *args, **kwargs)
                if fetch_rows > 0:
                    return cursor.fetchmany(fetch_rows)
                elif fetch_rows == -1:
                    return cursor.fetchall()
                else:
                    pass

    def create_signal_table(self, Scheme, SignalName):
        """Create table signal table """

        query = f"""DROP TABLE IF EXISTS "{Scheme}".{SignalName};
            CREATE TABLE "{Scheme}".{SignalName}
            (
                time_stamp time without time zone,
                time_stamp_int bigint,
                value double precision,
                "signalId" bigint,
                PRIMARY KEY (time_stamp)
            );

            ALTER TABLE "{Scheme}".{SignalName}
                OWNER to user_sandbox;
        """
        return query

    def create_table(self, Scheme, TableName, Columns, SheetData):
        
        primaryKey = ""
        query = f"""DROP TABLE IF EXISTS "{Scheme}".{TableName};
            CREATE TABLE "{Scheme}".{TableName}
            ("Id" bigint,
            """
        for Column in Columns:
            Type = self.Types.getColumnType(Column, SheetData) 
            if Type.count(" time ")>0:
                primaryKey = f"PRIMARY KEY ({Column})"
            query += Column + " " + Type+ ",\n"
   
        query += f"""
                {primaryKey}
            );

            ALTER TABLE "{Scheme}".{TableName}
                OWNER to user_sandbox;
        """
        print (query)
        self.execute_query(query)                

    def create_scheme(self, SchemeName):
        query = f"""CREATE SCHEMA IF NOT EXISTS "{SchemeName}"
                AUTHORIZATION user_sandbox;"""
        self.execute_query(query)
        
    def exists_table(self, SchemeName, TableName):
        query = f"""SELECT EXISTS (
                    SELECT FROM 
                        pg_tables
                    WHERE 
                        schemaname = '{SchemeName}' AND 
                        tablename  = '{TableName}'
                    );"""
        return self.execute_query(query,fetch_rows = -1)[-1][0]

    def insert_data(self, Scheme, TableName, Data):
        TimeStamp = ""
        query = f"""INSERT INTO {Scheme}.{TableName}(
        "Id","""
        for Column in Data[0]:
            query += Column + ", "
        query += ")\n VALUES "
        
        for Row in Data:
            Values = ""
            for i in Row:
                Value = Row[i]
                if type(Value) == datetime.datetime:
                    Id = str(int(Value.timestamp()))
                    Values = Id + ", " + Values
                    Value = str(Value)
                    TimeStamp = i

                if type(Value) == str:
                    Value = "'" + Value + "'"
                else:
                    Value = str(Value)
                Value = Value.replace("''","NULL")

                Values += Value + ", "

            query += f"({Values}),\n"
        query = query.replace(", )",")")
        query = query[:-2]
        
        query += f"\n ON CONFLICT ({TimeStamp}) DO NOTHING;"
##        print (query)
        self.execute_query(query)
        return query
                   
            
            
        

    def exists_column(self, SchemeName, TableName, ColumnName):
        query = f"""SELECT EXISTS (
                    SELECT FROM 
                        information_schema.columns
                    WHERE 
                        table_schema =  '{SchemeName}' AND 
                        table_name  = '{TableName}' AND
			column_name = '{ColumnName}'
                    ); """
        return self.execute_query(query,fetch_rows = -1)[-1][0]
    
    def get_table_from_database_using_query(self, query, *args, **kwargs):
        """Return DataFrame from DB using DataFrame.read_sql_query()"""
        import pandas as pd
        result = pd.read_sql_query(query, self.engine, "time_stamp", *args, **kwargs)
        return result
    
    def get_request_info(self, RequestId):
        """Returns tuple with machine name, table prefix, db_schema_id, request description, and requestor name"""
        query = f"""SELECT machine_name, table_prefix, schema_id, r.description, r.requestor_name
                    FROM machines
                    JOIN requests r on machines.machine_id = r.machine_id
                    WHERE request_id = {RequestId};
                """
        data = self.execute_query(query, fetch_rows=1)
        data = data[0]
        if self.verbose:
            print(f"Machine: {data[0]}")
        return data

