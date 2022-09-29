import time
from concurrent.futures import ProcessPoolExecutor,as_completed
from utils.psd_tools import process_dir,write_new_PSD
from utils.Dataframe_tools import group_then_separate_by
from utils.mp_psd import add_function_to_task,multi_PSD
from utils.file_ops import get_files_in_dir

if __name__ == '__main__':

    ##### INPUTS #######
    MYDIR= r"C:\Projects\2022\Michaels_Code\grundfos-express-tools\bronze impeller removal\input files"
    target_dir=r'C:\Projects\2022\Michaels_Code\grundfos-express-tools\bronze impeller removal\output files'

    list_of_pns = [
        96699290, 97775274, 96699299, 97775277, 96778078,
        97780992, 96699305, 96769184, 97778033, 96769190,
        96769205, 97778039, 96769256, 96896891, 96769259,
        96769271, 97780970, 96769280, 97780973
    ]
#     list_of_pns=[98876008,
# 98876012]

    str_list = [str(numstr) for numstr in list_of_pns]

    #Only want to add an Impeller Sheet 
    sheet_list=[]
    sheet='Impeller'

    for file in get_files_in_dir(MYDIR,('.xlsx')):
        sheet_list.append([sheet])
    
    tasks=process_dir(MYDIR,sheet_list,target_dir,False)
    add_function_to_task(tasks,group_then_separate_by,["Model", "Price ID"] ,"BOM",str_list=str_list)
    
    multi_PSD(tasks,write_new_PSD)