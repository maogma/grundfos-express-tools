import time
from concurrent.futures import ProcessPoolExecutor,as_completed
from utils.psd_tools import process_dir,write_new_PSD
from utils.Dataframe_tools import group_then_separate_by
from utils.mp_psd import add_function_to_task
from utils.file_ops import get_files_in_dir

if __name__ == '__main__':

    MYDIR= r"C:\Projects\2022\Michaels_Code\grundfos-express-tools\bronze impeller removal\input files"
    target_dir=r'C:\Projects\2022\Michaels_Code\grundfos-express-tools\bronze impeller removal\output files'

    list_of_pns = [
        96699290, 97775274, 96699299, 97775277, 96778078,
        97780992, 96699305, 96769184, 97778033, 96769190,
        96769205, 97778039, 96769256, 96896891, 96769259,
        96769271, 97780970, 96769280, 97780973
    ]

    #Only want to add an Impeller Sheet 
    sheet_list=[]
    sheet='Impeller'
    for file in get_files_in_dir(MYDIR,('.xlsx')):
        sheet_list.append([sheet])

    tasks=process_dir(MYDIR,sheet_list,target_dir,False)
    add_function_to_task(tasks,group_then_separate_by,["Model", "Price ID"] ,"BOM",str_list=list_of_pns)
    futures=[]
    start=time.time()
    with ProcessPoolExecutor() as executor:
        for task in tasks:
            future=executor.submit(write_new_PSD,*task)
            futures.append(future)
        for future in as_completed(futures):
            future.result()
    print(f'Process took {time.time()-start}')
