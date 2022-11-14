"""
Date last run:  Dec 14, 2021
Description:    Program reads event data from excel files and reports number of unique people and their ranks over a user-given period.
Versions:       Python - 3.8.2, Pandas - 1.3.4
Author:         Batmend Batsaikhan '24, DataSquad
Requirements:   -The program uses emails as a unique identifier. All entries should include emails. If there is no email, the program will not account that person.
                -The excel sheets should contain columns with names "Date", "Classification", "Email".
                -The "rank" column must be at the right side of the "Classification" column.
                -The dates should be formatted like 2017-01-31
"""

import os
import pandas as pd
import datetime

#These are all function declarations until line 172.

# Object that stores Data Frame of an event, its file path, and event date.
class DataFrameInfo:
  def __init__(self, pathString, eventDate, dataFrame): 
    self.pathString = pathString
    self.eventDate = eventDate
    self.dataFrame = dataFrame

# Gets the date of a given event. Here, we are picking first date of a "Date" column, and it should be a string of form 2021-01-31. However, the function also works if the date is datetime object. Finally, the function returns only the date.

def getDate(df):
    dateString = df.loc[0]["Date"]
    if (isinstance(dateString, str)): 
        dateObj = datetime.datetime.strptime(dateString, "%Y-%m-%d")
    else: 
        dateObj = dateString
    return dateObj.date()


#Given an excel sheet, reads it using pandas. Finally, it returns a DataFrameInfo object which is a dataFrame bundled with its file path and event date.
def readAndCleanExcelSheet(pathString):
    df = pd.read_excel(pathString)
    eventDate = getDate(df)
    df.fillna("Field is empty", inplace=True)
    return DataFrameInfo(pathString, eventDate, df)

#Iterates over the excel sheets in the given directory. For each file, constructs a dataframe with its metadata using subfunction, and makes a list out of them.
def readAndCleanDataInDirectory(directoryString):
    dataFrames = []
    for filename in os.listdir(directoryString):
        pathString = directoryString + "/" + filename
        dataFrames.append(readAndCleanExcelSheet(pathString))
    return dataFrames

# Given a string, change its year and return datetime.date() object.
def changeYear(year, date):
    dateString = year + date.strftime("%Y-%m-%d")[4:]
    return datetime.datetime.strptime(dateString, "%Y-%m-%d").date()

# Reads the term start and end date from "Academic_Calendar_2014-26.xlsx", and stores them in a dictionary termData. Here, notice that the dates in the excel sheet are all 2021 dates. Thus, we are changing the year using helper function.
def readTermData():
    termData = {}
    df = pd.read_excel("academic_calendar_2014-26.xlsx")
    for ind in df.index:
        startTime = changeYear(df["Term"][ind][-4:], df["Start"][ind])
        endTime = changeYear(df["Term"][ind][-4:], df["End"][ind])
        termData[df["Term"][ind]] = [startTime, endTime]
    return termData


# Function that bundles all printing functions
def printReports(classificationCounter, rankCounter, uniquePeopleNum, voidEmailsNum):
    printClassificationCounter(classificationCounter)
    printRankingCounter(rankCounter)
    print("\n----------------\nThere are", uniquePeopleNum, "unique attendees but", voidEmailsNum, "entries with void emails\n--------------------------------")

# Increments the value of a key by one. 
def insertPositionCounter(positionCounter, element):
    # Values can be NaN and we will not add them
    if isinstance(element, str):
        element = element.lower()
        #If key is not here, juust create one
        if element not in positionCounter:
            positionCounter[element] = 0
        positionCounter[element] += 1
    return positionCounter

# Prints all keys of a rankCounter dictionary
def printRankingCounter(positionCounter):
    print("\n----------------\n Ranking Report\n")
    for key in positionCounter.keys():
        print(key, ": ", positionCounter[key])

# Prints all keys of a classificationCounter dictionary
def printClassificationCounter(positionCounter):
    print("\n\n\n--------------------------------\n Classification Report\n")
    for key in positionCounter.keys():
        print(key, ": ", positionCounter[key])        

# This is the heart of our program. Every report is produced using this function. It takes a list of DataFramesInfoList object, adds all their members to a set, while counting the ranks.
def generateReportForListOfEventsInfo(reportingDataFramesInfoList):
    # Variables to store info of people.
    voidEmails = 0
    peopleSet = set()
    classificationCounter = {}
    rankCounter = {}
    for dataFrameInfo in reportingDataFramesInfoList:        
        for ind in dataFrameInfo.dataFrame.index:
            # We use email as unique identifier.
            email = dataFrameInfo.dataFrame["Email"][ind]
            # If there is no email, we can't process this person. However, our program will use voidEmails to report how many rows are missing their emails.
            if email == "Field is empty":
                voidEmails += 1
            # Add people only if they aren't in the set. This ensures uniqueness of people.
            elif email not in peopleSet:
                peopleSet.add(email)
                # both classificationCounter and rankCounter are dictionaries. They use insertPositionCounter helper function.
                classificationCounter = insertPositionCounter(classificationCounter, dataFrameInfo.dataFrame["Classification"][ind])
                rankCounter = insertPositionCounter(rankCounter, dataFrameInfo.dataFrame["Short_title"][ind])
    printReports(classificationCounter, rankCounter, len(peopleSet), voidEmails)

# Checks if a current date is in a range. Used in generateReportForTimeRangeGeneral function.
def isBetween(date, startDate, endDate):
    return (startDate < date or startDate == date) and (date < endDate or date == endDate)



# This is a main helper function that is based on generateReportForListOfEventsInfo, but works with time ranges.
def generateReportForTimeRangeGeneral(startDate, endDate):
    reportingDataFramesInfoList = []
    for dataFrameInfo in dataFrameInfoList:
        if isBetween(dataFrameInfo.eventDate, startDate, endDate):
            reportingDataFramesInfoList.append(dataFrameInfo)
    return generateReportForListOfEventsInfo(reportingDataFramesInfoList)
    
# It only takes the file name of a given event. It runs through dataFrameInfo and finds the event Info and then urns generateReportForListOfEventsInfo Helper function.
def generateReportForEvent():
    eventName = input("\nEnter filename for the event: ")
    for dataFrameInfo in dataFrameInfoList:
        if eventName in dataFrameInfo.pathString:
            return generateReportForListOfEventsInfo([dataFrameInfo])

# Takes a single term (Fall 2017 for example). It uses pre-stored termInfoDict dictionary for term start and end info. Then it uses generateReportForTimeRangeGeneral helper function.
def generateReportForTerm():
    termString = input("\nEnter Term (Example, Fall 2017): ")
    return generateReportForTimeRangeGeneral(termInfoDict[termString][0], termInfoDict[termString][1])

# Takes a single year (2017 for example). It uses generateReportForTimeRangeGeneral helper function.
def generateReportForYear():
    year = int(input("\nEnter year: "))
    return generateReportForTimeRangeGeneral(datetime.datetime(year, 1, 1).date(), datetime.datetime(year, 12, 31).date())

# Takes an user given time range. Dates should be of a for 2021-01-21. It uses generateReportForTimeRangeGeneral helper function.
def generateReportForTimeRange():
    startDate = datetime.datetime.strptime(input("\nEnter start date (Example, 2017-01-31): "), "%Y-%m-%d").date()
    endDate = datetime.datetime.strptime(input("\nEnter end date(Example, 2017-01-31): "), "%Y-%m-%d").date()
    return generateReportForTimeRangeGeneral(startDate, endDate)

def generateReportForListOfEvents():
    eventNumber = int(input("How many events are you entering: "))
    reportingDataFramesInfoList = []
    for i in range(0, eventNumber):
        eventName = input("\nEnter filename for the event " + str(eventNumber + 1) + " :")
        for dataFrameInfo in dataFrameInfoList:
            if eventName in dataFrameInfo.pathString: 
                reportingDataFramesInfoList.append(dataFrameInfo)    
    return generateReportForListOfEventsInfo(reportingDataFramesInfoList)


# Given the user choice, this function then runs corresponding functions
def generateReport():
    userInput = input("\n\n--------\n\nChoose number of the option:\n\n 1. Choose an event\n 2. Choose a term\n 3. Choose a year\n 4. Specify a range \n 5. Choose list of events\n\n (For example, to choose term, type 2 then hit ENTER)\n\n:")
    if userInput == "1":
        return generateReportForEvent()
    elif userInput == "2":
        return generateReportForTerm()
    elif userInput == "3":
        return generateReportForYear()
    elif userInput == "4":
        return generateReportForTimeRange()
    elif userInput == "5":
        return generateReportForListOfEvents()
    else:
        print("Please insert a valid option")
        return 

# It keeps running the program until the user no longer needs another report.
def queryHoweverManyTimes():
    userInputForContinue = 1
    while userInputForContinue:
        generateReport()
        userInputForContinue = int(input("\nDo you want to generate another report?\nType 1 if yes, 0 if no: "))

#Main code that actually runs.

#Function takes the name of the folder where all the excel files are. This folder must be in the same directory as this code (main.py).
dataFrameInfoList = readAndCleanDataInDirectory("EventData_WithShortTitles")

# Reads the start and end date of all terms from 2014 to 2026. Reads from Academic_Calendar_2014-26.xlsx.
termInfoDict = readTermData()

#Takes user choices and produces an output.
queryHoweverManyTimes()






