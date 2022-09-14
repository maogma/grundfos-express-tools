
import datetime
import os

def add_str_to_filename(filepath_to_copy: str,strng) -> str:
    """ Creates a working copy of the file. Useful when performing changes to an excel spreadsheet and wanting to keep original file unchanged"""
    _, filename = os.path.split(filepath_to_copy) # Splits filepath to folder, file
    working_copy = strng + " " + filename # adds timestamp to filename
    return(working_copy)


def add_filename_timestamp(filepath_to_copy: str) -> str:
    """ Creates a working copy of the file. Useful when performing changes to an excel spreadsheet and wanting to keep original file unchanged"""
    dir, filename = os.path.split(filepath_to_copy) # Splits filepath to folder, file
    timeStamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    working_copy = timeStamp + " " + filename # adds timestamp to filename
    working_copy = os.path.join(dir, working_copy) # concats to new filepath
    return(working_copy)

def batchRenameAppend(myDir, appendString):
    """ This function appends a string to all the files within the directory """
    for root, subdirectories, files in os.walk(myDir):             # Goes through all dir, subdir to find new schematic pdfs
        for file in files:
            pathname, extension = os.path.splitext(file)
            newFilename = pathname + appendString                                         # Strips file type extension 
            originalFilename = os.path.join(root, pathname + extension)
            newFilename = os.path.join(root, newFilename + extension)
            os.rename(originalFilename, newFilename)
    return

def iterateSubDirFiles(workingDirectory, *extraFunctions):
    """ Iterates through all files within a directory and perform function on each file"""
    for root, subdirectories, files in os.walk(workingDirectory):
        for file in files:
            # fileName, fileExtension = os.path.splitext(file)
            fileName = os.path.join(workingDirectory, file)
            print("Opening file for updates: {}".format(fileName))
            if len(extraFunctions) > 0:
                functionForFile = extraFunctions[0]
                functionForFile(fileName)
    return
###############################################################################