"""
- Iterate through data and find which rows need to be duplicated
- Make note of original row indeces to duplicate, as well as list of values from column N and O
- Calculate number of rows to duplicate
- Create appropriate # of rows
- Write new dataframe to excel sheet
"""

from openpyxl import load_workbook
import os
import pandas as pd
import numpy as np
import time
import logging

logging.basicConfig(format='%(asctime)s %(levelname)-8s [%(filename)s:%(lineno)d] %(message)s', datefmt='%m-%d-%Y %H:%M:%S', level=logging.INFO, filename='logs.txt')

time_logger = logging.getLogger('script_timer')

myDir = r"C:\Users\104092\OneDrive - Grundfos\Documents\git\grundfos-express-tools\GES_Dimension_Expansion"
myFile = "Mech GRP foundation for VLSE rev6.xlsx"
filePath = os.path.join(myDir, myFile)
tabName = "GES_VLSE"

# Here are the 2 columns of interest
colN = "Branch Size Range, Numerical List (inches)"
colV = "Header Size Range, Numerical List (inches)"

# Creates dataframe from excel tab of interest
data = pd.read_excel(filePath, sheet_name='GES_VLSE', header=0, index_col=False)

# This creates a new empty dataframe which will hold the new rows
column_names = data.columns
extraRowsDataframe = pd.DataFrame(columns=column_names)


def keepNumType(numString):
    """ This function will retain the cell value type as a int/float instead of a string"""
    try:
        # Convert it into integer
        val = int(numString)
        return val
    except ValueError:
        try:
            # Convert it into float
            val = float(numString)
            return val
        except ValueError:
            print("No.. input is not a number. It's a string")
    return


start = time.time()

# Iterates through the original dataframe
for index, row in data.iterrows():
    # Splits any comma separated values in column N
    try:
        branchSizes = row[colN].split(',')
        multiplier1 = len(branchSizes)
    # If error, then set multiplier to 1 because only one value exists in cell   
    except AttributeError:
        multiplier1 = 1
    
    # Splits any comma separated values in column V
    try:
        headerSizes = row[colV].split(',')
        multiplier2 = len(headerSizes)
    # If error, then set multiplier to 1 because only one value exists in cell   
    except AttributeError:
        multiplier2 = 1
        
    # Checks how many rows to add, and modifies the values in the appropriate columns
    # "currentRow" variable is used to find last row of dataframe
    # Case: only 1 value in column N and column V
    if multiplier1 == 1 and multiplier2 == 1:
        currentRow = len(extraRowsDataframe)      
        extraRowsDataframe.loc[currentRow] = row # Copies existing row from original data to new dataframe
    # Case: only 1 value in column N and multiple in column V
    elif multiplier1 == 1 and multiplier2 > 1:
        for header in headerSizes:
            currentRow = len(extraRowsDataframe)
            extraRowsDataframe.loc[currentRow] = row
            extraRowsDataframe.loc[currentRow, colN] = row[colN]
            extraRowsDataframe.loc[currentRow, colV] = keepNumType(header)
    # Case: Multiple values in column N and 1 value in column V
    elif multiplier1 > 1 and multiplier2 == 1:
        for branch in branchSizes:
            currentRow = len(extraRowsDataframe)
            extraRowsDataframe.loc[currentRow] = row
            extraRowsDataframe.loc[currentRow, colN] = keepNumType(branch)
            extraRowsDataframe.loc[currentRow, colV] = row[colV]
    # Case: Multiple values in both column N and column V
    elif multiplier1 > 1 and multiplier2 > 1:
        for branch in branchSizes:
            for header in headerSizes:
                currentRow = len(extraRowsDataframe)
                extraRowsDataframe.loc[currentRow] = row
                extraRowsDataframe.loc[currentRow, colN] = keepNumType(branch)
                extraRowsDataframe.loc[currentRow, colV] = keepNumType(header)    

time_logger.info(f"iterating through dataframe took {time.time() - start}")

# Adds an extra column to filter out invalid combinations (where header is not > branch)
extraRowsDataframe['Valid Combo'] = np.where(extraRowsDataframe[colN] >= extraRowsDataframe[colV], "INVALID","")

write_start = time.time()

# Writing new dataframe to a new tab in the excel file
with pd.ExcelWriter(filePath, engine="openpyxl", mode='a') as writer:  
    extraRowsDataframe.to_excel(writer, sheet_name="Expanded Rows", index=False)

time_logger.info(f"writing to excel sheet took {time.time() - write_start}")    

# Hiding columns to match source excel sheet formatting
wb = load_workbook(filePath)
ws = wb["Expanded Rows"]
ws.column_dimensions.group(start='P', end='R', hidden=True)
ws.column_dimensions.group(start='T', end='U', hidden=True)
ws.column_dimensions.group(start='W', end='AA', hidden=True)
ws.column_dimensions.group(start='AF', end='AS', hidden=True)

# for col in ['P':'U']:
#     ws.column_dimensions[col].hidden= True

wb.save(filePath)






