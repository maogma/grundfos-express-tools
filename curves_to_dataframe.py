import os
import pandas as pd

# This is the folder/file with the curve export csv
myDir = r"C:\Users\104092\OneDrive - Grundfos\!Projects\Add to GXS - Alpha\Archive"
myFile = "PumpCurves.csv"
filePath = os.path.join(myDir, myFile)

# This creates a dataframe of the curve export csv, and fills in the RPM(curve nominal) column
data = pd.read_csv(filePath, sep=";", index_col=False, skip_blank_lines=False)
data['RPM(Curve nominal)'] = data['RPM(Curve nominal)'].ffill()

data.drop(data.iloc[:, 0:12], inplace=True, axis=1)
