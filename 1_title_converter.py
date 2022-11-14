import pandas as pd

def titleConverter(longTitle):
    if longTitle.lower().find("dean of") != -1:
        shortTitle = "other staff member"
    elif longTitle.lower().find("visiting") != -1:
        shortTitle = "visiting professor"
    elif longTitle.lower().find("assistant professor") != -1:
        shortTitle = "assistant professor"
    elif longTitle.lower().find("associate professor") != -1:
        shortTitle = "associate professor"
    elif longTitle.lower().find("professor") != -1:
        shortTitle = "professor"
    elif longTitle.lower().find("lecturer") != -1:
        shortTitle = "lecturer"
    elif longTitle == "":
        shortTitle = "student"
    else: 
        shortTitle = "other staff member"
    
    return shortTitle

excel_sheet_with_long_titles = input("Enter name of the excel sheet containing long titles: ")
# excel_sheet_with_long_titles = input("names_titles-13Nov21.xlsx")

usefulColumns = ["Carleton_Name", "Email", "Primary_Position_Title"]

df = pd.read_excel(excel_sheet_with_long_titles)
df = df[usefulColumns]
df.fillna("Empty String", inplace=True)
df['Short_title'] = [''] * len(df.index)

rows_with_missing_data = []

for ind in df.index:
    for column in usefulColumns:
        if df[column][ind] == 'Empty String':
            print("Entry " + str(ind+1) + " has no " + column)
            rows_with_missing_data.append(ind)
            break

    if len(rows_with_missing_data) == 0 or rows_with_missing_data[-1] != ind:
        df['Short_title'][ind] = titleConverter(df['Primary_Position_Title'][ind])

df = df.drop(rows_with_missing_data)

df.to_excel("faculty_and_staff_short_title.xlsx")


