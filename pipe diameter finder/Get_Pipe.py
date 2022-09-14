from utils.file_ops import read_files_in_dir,add_str_to_filename
from utils.Dataframe_tools import write_df_to_excel,get_df
from utils.excel_tools import write_excel_file


Input_dir=r"C:\Projects\2022\Michaels_Code\grundfos-express-tools\pipe diameter finder\input files"
Output_dir=r"C:\Projects\2022\Michaels_Code\grundfos-express-tools\pipe diameter finder\output files"
a=read_files_in_dir(Input_dir)
sheetname='Max flow to diameter'
new_sheet=sheetname+'_Completed'
while True:
    try:
        (d1,file)=next(a)
        d2=get_df(d1,sheetname)
        new_path=write_excel_file(d1,Output_dir,add_str_to_filename,file,'Revised')
        write_df_to_excel(d2,new_path,new_sheet)
    except StopIteration:
        break