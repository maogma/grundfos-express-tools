########################################################################
# TO DOs
########################################################################

if __name__ == "__main__":
    '''Modifies "Curve Header Data" tab in Curve PSD to reflect newly separated trim tabs, prior to importing into SKB.
    - Pre-Requisite: Need to have original "Curve Header Data" populated
    - Load "Curve Header Data" tab as ws
    - Loop through all curve tabs, and create copy of relevant entry from "Curve Header Data" entry
    '''

    import os
    import openpyxl
    import pandas as pd
    import time
    
    ########################################################################
    # NEED TO MODIFY FOLDERS, FILENAMES BELOW TO REFLECT FOLDER SETUP
    ########################################################################
    fileDir = r"C:\Users\104092\OneDrive - Grundfos\Documents\10-19 Projects\12 NBS Curve PSD Separation\testing"
    file_to_change = r"GXS Curve_Conexus_V2 - std models.xlsx"
    filePath = os.path.join(fileDir, file_to_change)
    tabName = "Curve Header Data"
    newPumpFamilyName = "NBS_Fixed_Trim_Std"
    ########################################################################

    start = time.time()
    raw_data = pd.read_excel(filePath, sheet_name=tabName, header=8)                            # Creates dataframe from excel tab of interest
    wb = openpyxl.load_workbook(filePath, data_only=True)                                       

    new_curves_list = [ sheet.title for sheet in wb.worksheets if type(sheet.title)==str ]      # Create list of all new curve numbers 

    rows_list = []                                                                              # will hold list of dataframes to be removed. Will concatenate to 1 dataframe later

    for _, col_data in raw_data.iterrows():                                                     # Loop through each curve number entry in curve header data tab
        if type(col_data['Curve number']) == str:                                               # For each matching trim worksheet found, append duplicate entry to new dataframe, but rename curve number to match
            for curve_number in new_curves_list:
                if col_data['Curve number'] in curve_number:
                    mydict = col_data.to_dict()
                    mydict.update({'Curve number':curve_number})
                    rows_list.append(mydict)

    updated_data = pd.DataFrame(rows_list)                                                      # Concatenating list of dfs to single dfs.
 
    psd_startrow = 10                                                                           # Write resulting dataframes to new sheet in PSD with changes.
    with pd.ExcelWriter(filePath, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:  
        updated_data.to_excel(writer, sheet_name=tabName, index=False, header=False, startrow=psd_startrow)

    wb[tabName]['B7'] = newPumpFamilyName
    wb.save(filePath)

    print(f'Entire script took {time.time() - start} seconds')