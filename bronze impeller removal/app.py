# %%
import os
from utils.file_ops import add_filename_timestamp
import shutil
import datetime
import pandas as pd

myDir = r"C:\Users\104092\OneDrive - Grundfos\Documents\git\grundfos-express-tools\bronze impeller removal\input files"

list_of_pns = [
    96699290, 97775274, 96699299, 97775277, 96778078,
    97780992, 96699305, 96769184, 97778033, 96769190,
    96769205, 97778039, 96769256, 96896891, 96769259,
    96769271, 97780970, 96769280, 97780973
]

str_list = [str(numstr) for numstr in list_of_pns]

# Reason/Justification for removal
timeStamp = datetime.datetime.now().strftime("%m-%d-%Y")
reason = input("Reason for removal: ")
authority = input("Who authorized/requested this change? ")

removal_note = timeStamp + " " + reason + " per " + authority

""" Iterates through all files within a directory and perform function on each file"""
for root, subdirectories, files in os.walk(myDir):
    for file in files:
        # fileName, fileExtension = os.path.splitext(file)
        filePath = os.path.join(myDir, file)
        print("Opening file for updates: {}".format(filePath))

        # Create working copy                    
        working_copy = add_filename_timestamp(file)
        completed_dir = r"C:\Users\104092\OneDrive - Grundfos\Documents\git\grundfos-express-tools\bronze impeller removal\output files"
        working_copy = os.path.join(completed_dir,working_copy)
        shutil.copyfile(filePath, working_copy) # creates working copy to leave original file untouched


        # Read in Excel sheet to a dataframe
        sheetname = "Impeller"
        psd_data = pd.read_excel(working_copy, sheet_name=sheetname, header=5, dtype={'BOM': str})

        # Need to filter out old rows that have been previously removed (typically at bottom of PSDs)
        last_psd_row = psd_data[psd_data['Full Data'] == "[END]"].index.to_list()[0] # Find [END] row/index 
        data = psd_data.iloc[:last_psd_row]

        # Customize based on PST template used. 
        psd_startrow = 5 # Corresponds to row number containing "second" column headers
        psd_startcol = 0

        # Find matching partnumber's from list above, and then capture "process variants"
        group_model_matlcode = data.groupby(["Model", "Price ID"])

        df_list_to_remove = []
        df_list_to_keep = []

        for key, frame in group_model_matlcode:
            frame.reset_index(inplace=True)        # Reset index for each group

            if any(frame['BOM'].isin(str_list)):   # Finding matching partnumbers to remove
                df_list_to_remove.append(frame)    # Need to add this sub-group to a "removed dataframe"
            else:
                df_list_to_keep.append(frame)      # Need to retain this sub-group, add to a "keep dataframe"


        # Concatenating list of dfs to single dfs.
        removals = pd.concat(df_list_to_remove)
        keep = pd.concat(df_list_to_keep)

        # Dataframe formatting.
        removals.at[0, 'Full Data'] = ""             # This removes the [START] if the first PSD row is flagged to be removed
        removals.drop('index', axis=1, inplace=True) # Need to drop extra "Index" Column
        keep.drop('index', axis=1, inplace=True)

        # Insert removal justification note to top of removal dataframe
        new_row = pd.DataFrame({'ID': removal_note}, index =[0])
        removals = pd.concat([new_row, removals[:]]).reset_index(drop = True) # Concatenate new_row with df

        # Reorder columns to match the 2 dataframe columns prior to concatenation
        column_list = keep.columns
        removals = removals[column_list]

        # Check for any misses before writing. Should return empty dataframe
        print(keep[keep['BOM'].isin(str_list)])


        # Create copy of PSD tab with formatting
        from openpyxl import load_workbook

        wb = load_workbook(working_copy)
        ws = wb['Impeller']

        tabName = sheetname + " No Bronze"
        wb.copy_worksheet(ws).title = tabName

        # Find where to append removals and insert reason why they were removed
        original_end_location = len(psd_data) + psd_startrow          # This is the last row of the original PSD (0-indexed) 
        append_location = original_end_location - len(removals) + 2   # This is the new location/row where we will append the removals

        # Edit New Tab to clear cells with old data after first [END] value is reached
        num_rows = len(keep)                          # Finds [END] in dataframe
        after_end_row = num_rows + psd_startrow + 1   # Calculates where [END] row + 1 will be in the PSD

        ws_modified = wb[tabName]
        ws_modified.delete_rows(after_end_row-1, len(removals)-1)   # clears the rows containing old data

        # Fix formatting
        from openpyxl.styles import PatternFill
        for rows in ws_modified.iter_rows(min_row=append_location+2, max_row=append_location+2, min_col=1, max_col=40):
            for cell in rows:
                cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        wb.save(working_copy)

        # Write resulting dataframe to a template PSD
        with pd.ExcelWriter(working_copy, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:  
            keep.to_excel(writer, sheet_name=tabName, index=False, startrow=psd_startrow, startcol=psd_startcol)

        # This writes the removed data to the bottom of the newly created tab
        with pd.ExcelWriter(working_copy, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:  
            removals.to_excel(writer, sheet_name=tabName, index=False, header=False, startrow=append_location+1, startcol=psd_startcol)

        # Remove original tab, replace with new one



