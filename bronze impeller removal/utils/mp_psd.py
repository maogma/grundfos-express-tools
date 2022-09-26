import time
from concurrent.futures import ProcessPoolExecutor,as_completed


def add_function_to_task(tasks, function, *args, **kwargs):
    for task in tasks:
        task.insert(2, (function, list(args), dict(kwargs)))

def multi_PSD(tasks,exe_func):
    futures=[]
    start=time.time()
    with ProcessPoolExecutor() as executor:
        for task in tasks:
            future=executor.submit(exe_func,*task)
            futures.append(future)
        for future in as_completed(futures):
            future.result()
    print(f'Process took {time.time()-start}')
