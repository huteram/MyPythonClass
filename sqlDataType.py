import datetime
class sqlDataType:
    def __init__(self):
        self.Types = self.getTypes()

    def getTypes(self):
        Types = {}
        Types[int] = 'double precision'
        Types[float] = 'double precision'
        Types[datetime.datetime] = 'timestamp without time zone'
        Types[str] = 'text'
        return Types
    
    def getColumnType(self, ColumnName, SheetData):
        Type = None
        for i in SheetData:
            if ColumnName in i:
                if (i[ColumnName])!= '':
                    try:
                        Type = self.Types[type(i[ColumnName])]
                        break
                    except:
                        print ("Unknown data type :",type(i[ColumnName]))
        if Type == None:
            Type = str
        return Type
f = sqlDataType()

                    
        
        
