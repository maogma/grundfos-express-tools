import os
import typing
from _types._types import *


def write_excel_file(file_data: bytes, target_dir: str, filename_mod_func: Callable[..., typing.Any], *func_args: typing.Any) -> str:
    """The function takes in binary data and writes that data to an Excel file.
    
    Params:
        file_data: The binary data to write to the excel file.
        target_dir: The target diretory to save the file to.
        filename_mod_func: The function that modifies the file name (Should be optional but it is not right now).
        *func_args: The arguments of the filename_mod_func (Should also be optional but it is not right now).

    Return:
        Returns the file path of the excel file that the data was written to."""
    new_file_path = os.path.join(target_dir, filename_mod_func(*func_args))
    with open(new_file_path, 'wb') as f:
        f.write(file_data)
    return new_file_path  # return this only if a new file was created

# Need to filter out old rows that have been previously removed (typically at bottom of PSDs)


def last_psd_row(psd_data):
    # Find [END] row/index
    return psd_data[psd_data['Full Data'] == "[END]"].index.to_list()[0]
