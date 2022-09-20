###############################################################################
"""  Finds number of rows in a PSD by finding [END] in column A """  
def findLastRow(sheet):
    
    for myRow in range(1, sheet.max_row + 1):
        for column in "A":  
            cell_name = "{}{}".format(column, myRow)
            if sheet[cell_name].value == "[END]":
                return myRow
            
###############################################################################

###############################################################################
def find_specific_cell(searchValue, row_or_col, cols="ABCDEFGHIJKLMNOPQRSTUVWXYZ"):
    # print("looking for {}, type: {}".format(searchValue,type(searchValue)))
    
    last_row = findLastRow(dimSheet)
    
    for eachRow in range(1, last_row):
        for column in cols:  # Here you can add or reduce the columns
            cell_name = "{}{}".format(column, eachRow)

            # Checking when excel cell value matches searchValue            
            if dimSheet[cell_name].value == searchValue:
                # print("cell position {} has value {}".format(cell_name, dimSheet[cell_name].value))
                if row_or_col == "row":
                    return eachRow
                elif row_or_col =="col":
                    return column 

##############################################################################

