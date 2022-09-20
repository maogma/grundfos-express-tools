import pandas as pd
import numpy as np
from pandas._typing import FilePath,ReadBuffer
from _types._types import *

def get_df(data:Union[FilePath,ReadBuffer[bytes],bytes],sheet_names:str|list[str])->pd.DataFrame:
    """Needing Implementation:
    1) Check if sheets are iterable or not 
    2) Check if sheets are in the excel doc
    3) Check if the data is in the correct format
    
    This function returns a dataframe from excel data. 
    
    Params:
        data: This can be a path to a file or a binary object (i.e. a read file).
        sheet_names: This is a list or a singular string representing the name of a sheet in the excel file.
    
    Returns:
        A pandas dataframe object.
        """
    #Reading in the dataframe with the given sheet name
    df1=pd.read_excel(data,sheet_name=sheet_names)
    
    
    branch_df= df1.copy()[["Max Branch Flow (gpm)","Max Branch Diameter (in.)"]] #creating the branch dataframe from a copy of df1
    branch_df.loc[:,'copy_index']=branch_df.index #creating a copy of the current index as a new column 'copy_index'
    branch_df["Max Branch Flow (gpm)"] = branch_df["Max Branch Flow (gpm)"].astype(float) #casting the 'Max Branch Flow' to type float 
    branch_df.dropna(axis='index', how='any', inplace=True,subset=["Max Branch Flow (gpm)"]) #Looking in the 'Max Branch Flow' col and removing any rows with a null value
    

    header_df = df1.copy()[["Max Header Flow (gpm)","Max Header Diameter (in.)"]] #creating the header dataframe from a copy of df1
    header_df.loc[:,'copy_index']=header_df.index #creating a copy of the current index as a new column 'copy_index'
    header_df["Max Header Flow (gpm)"] = header_df["Max Header Flow (gpm)"].astype(float) #casting the 'Max Header Flow' to type float 
    header_df.dropna(axis='index', how='any', inplace=True,subset=["Max Header Flow (gpm)"]) #Looking in the 'Max Header Flow' col and removing any rows with a null value


    reference_df = df1.copy()[["Flow (gpm)", "Pipe Diameter (in.)"]] #creating the new dataframe from a copy of df1 with the columns 
    reference_df["Flow (gpm)"] = reference_df["Flow (gpm)"].astype(float) #creating a copy of the current index as a new column 'copy_index'
    reference_df.set_index("Flow (gpm)", inplace=True) #Set the dataframe index using the column Flow
    reference_df.dropna(axis='index', how='any', inplace=True) #Removing any rows with a null value in any column

    branch_output_df = pd.merge_asof(branch_df.sort_values('Max Branch Flow (gpm)'), reference_df, left_on="Max Branch Flow (gpm)", right_on="Flow (gpm)", direction='backward')
    branch_output_df.sort_values(by=['copy_index'], inplace=True)

    header_output_df = pd.merge_asof(header_df.sort_values('Max Header Flow (gpm)'), reference_df, left_on="Max Header Flow (gpm)", right_on="Flow (gpm)", direction='backward')
    header_output_df.sort_values(by=['copy_index'], inplace=True)

    output_df = pd.merge(branch_output_df, header_output_df, on="copy_index")
    output_df.drop(['copy_index'], axis=1, inplace=True)

    output_df["Max Branch Diameter (in.)"]=np.where(output_df["Max Branch Diameter (in.)"]==output_df['Pipe Diameter (in.)_x'],output_df["Max Branch Diameter (in.)"],output_df['Pipe Diameter (in.)_x'])
    output_df["Max Header Diameter (in.)"]=np.where(output_df["Max Header Diameter (in.)"]==output_df['Pipe Diameter (in.)_y'],output_df["Max Header Diameter (in.)"],output_df['Pipe Diameter (in.)_y'])
    output_df.drop(columns=['Pipe Diameter (in.)_x','Pipe Diameter (in.)_y'],inplace=True)

    return output_df

def write_df_to_excel(df:pd.DataFrame,file_path:FilePath,sheet:str)->None:
    """This function takes in a dataframe and writes it to the path given by the file path parameter.
    
    Params:
        df: A pandas dataframe for a given sheet within an excel workbook.
        file_path: A file path to write the dataframe to.
        sheet: The name of the sheet you want to append to.
    
    Returns:
        It returns nothing.
    
    This function is a wrapper for the pandas ExcelWriter method in the sense that you can call this function instead of that one."""
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:  
        df.to_excel(writer, sheet_name=sheet, index=False)