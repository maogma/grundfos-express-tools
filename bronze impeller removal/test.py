import datetime
import pandas as pd
import os
import numpy as np
import time
import multiprocessing as mp
from multiprocessing import Event
from pandas import DataFrame
from openpyxl import load_workbook,worksheet,Workbook
from typing import Callable,Any,ParamSpec,TypeVar,Concatenate
from _utils.file_ops import read_files_in_dir
from _utils.Dataframe_tools import get_df
from concurrent.futures import ProcessPoolExecutor,as_completed


MYDIR= r"C:\Projects\2022\Michaels_Code\grundfos-express-tools\bronze impeller removal\input files"
P = ParamSpec('P')
files=read_files_in_dir(MYDIR)
sheetname='Impeller'

# futures_list=[]
# results = []

def print_file(file,lock):
    lock.acquire()
    _,fname=file
    inpt=input('Test:')
    lock.release()
    return fname

if __name__=='__main__':
    futures=[]
    manager=mp.Manager()
    lock=manager.Lock()
    start=time.time()
    with ProcessPoolExecutor() as executor:
        for file in files:
            future=executor.submit(print_file,file,lock)
            futures.append(future)
            future.result()
        for future in as_completed(futures):
            print(future.result())
    print(f'Process took {time.time()-start}')



##########################################
# from time import sleep
# from random import random
# from multiprocessing import Process
# from multiprocessing import Event
 
# # target task function
# def task(event, number):
#     # wait for the event to be set
#     print(f'Process {number} waiting...', flush=True)
#     event.wait()
#     # begin processing
#     value = random()
#     sleep(value)
#     print(f'Process {number} got {value}', flush=True)
 
# # entry point
# if __name__ == '__main__':
#     # create a shared event object
#     event = Event()
#     # create a suite of processes
#     processes = [Process(target=task, args=(event, i)) for i in range(5)]
#     # start all processes
#     for process in processes:
#         process.start()
#     # block for a moment
#     print('Main process blocking...')
#     sleep(2)
#     # trigger all child processes
#     event.set()
#     # wait for all child processes to terminate
#     for process in processes:
#         process.join()
