import time
from utils import add_function_to_task,multi_PSD,get_files_in_dir,group_then_separate_by,process_dir,write_new_PSD

if __name__ == '__main__':

    ##### INPUTS #######
    MYDIR= r"C:\Projects\2022\Michaels_Code\grundfos-express-tools\bronze impeller removal\input files"
    target_dir=r'C:\Projects\2022\Michaels_Code\grundfos-express-tools\bronze impeller removal\output files'

    list_of_pns = [
        97775291,96699335,96769175,97775278,96769178,96769229,97780146
    ]
#     list_of_pns=[98876008,
# 98876012]

    str_list = [str(numstr) for numstr in list_of_pns]

    #Only want to add an Impeller Sheet 
    sheet_list=[]
    sheet='ImpellerModified'

    for file in get_files_in_dir(MYDIR,('.xlsx')):
        sheet_list.append([sheet])
    
    tasks=process_dir(MYDIR,sheet_list,target_dir,False)
    add_function_to_task(tasks,group_then_separate_by,["Model", "Price ID"] ,"BOM",str_list=str_list)
    
    multi_PSD(tasks,write_new_PSD)
