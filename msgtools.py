from asammdf import MDF
import cantools as ct
import os
##import dotenv
##from utils import timed

##dotenv.load_dotenv()

def mdf_to_dataframe(source_mdf_file, start=None, end=None):
    """Converts mdf file into Pandas DataFrame"""
    try:
        mdf = MDF(source_mdf_file)
        mdf = mdf.cut(start=start, stop=end)
        df = mdf.to_dataframe(time_from_zero=False, time_as_date=True)
        Columns = []
        for i in df.columns:
            Columns.append(i.split('.')[-1])
                                  
        df.columns = Columns   #['BusChannel', 'ID', 'IDE', 'DLC', 'DataLength', 'DataBytes', 'Dir', 'EDL', 'ESI', 'BRS']
        DropColumns = []
        WantedColumns = ['BusChannel', 'ID', 'DataBytes']
        for i in Columns:
            if not(i in WantedColumns):
                DropColumns.append(i)
            
        df.drop(columns=DropColumns, inplace=True)
        df["Data"] = df["DataBytes"].apply(mdf_data_to_str)
        df.drop(columns="DataBytes", inplace=True)
        df.reset_index(inplace=True)
        df.columns = ['time_stamp', 'bus_channel', 'msg_id', 'msg_data']
        try:            
            df["time_stamp"] = df["time_stamp"].dt.tz_convert(None)
            print("set time zone None")
        except:
            print("time zone is not used.")
    except:
        return None

    return df

def mdf_data_to_str(source_data):
    """
    Return string representation of MDF CAN data
    """
    data = tuple(source_data)
    data = (f"{i:02x}" for i in data)
    return "".join(data)

