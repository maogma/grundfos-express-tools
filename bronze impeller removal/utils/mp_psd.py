
def add_function_to_task(tasks, function, *args, **kwargs):
    for task in tasks:
        task.insert(2, (function, list(args), dict(kwargs)))
