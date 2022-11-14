import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import os

shortTitles =  pd.read_excel("faculty_and_staff_short_title.xlsx")

def addShortTitleToDirectory(directoryString):
    for filename in os.listdir(directoryString):
        addShortTitleToExcelSheet(directoryString, filename)

def addShortTitleToExcelSheet(directoryString, filename):
    pathString = directoryString + '/' + filename
    df = pd.read_excel(pathString)
    df['Short_title'] = [''] * len(df.index)
    df.fillna("", inplace=True)
    for ind in df.index:
        if df['Classification'][ind] != 'STUDENT':
            rowsWithMatchingEmail = shortTitles[shortTitles['Email'].str.match(df['Email'][ind])].index
            if len(rowsWithMatchingEmail) != 0:
                df['Short_title'][ind] = shortTitles['Short_title'][rowsWithMatchingEmail[0]]
            else: 
                df['Short_title'][ind] = "short title missing"
                print("Short title missing from entry " + str(ind) + " " + filename)
        else:
            df['Short_title'][ind] = "Student"
    df.to_excel("EventData_WithShortTitles/"+filename, index=False)

if not os.path.exists("EventData_WithShortTitles"):
  os.makedirs("EventData_WithShortTitles")

addShortTitleToDirectory("EventData_Cleaned")