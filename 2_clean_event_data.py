import pandas as pd
import os

def cleanExcelSheet(directoryString, filename):
    pathString = directoryString + "/" + filename
    df = pd.read_excel(pathString)
    df = df[["Date", "Email", "Classification"]]
    df = df.dropna(subset=["Email"])
    df.to_excel(targetDirectory + "/" + filename, index=False)


directoryString = "EventData"
targetDirectory = directoryString + "_Cleaned"

if not os.path.exists(targetDirectory):
  os.makedirs(targetDirectory)

for filename in os.listdir(directoryString):
        cleanExcelSheet(directoryString, filename)


