This program was written for macOS machines by [Batmend Batsaikhan 24'](batmendbar@gmail.com).

# How to run this program: 

Enter following commands in the [mac terminal](https://www.youtube.com/watch?v=aKRYQsKR46I) from the folder containing all the scripts (titleConverter.py, clean_event_data.py, etc.).

#### Step 0. If your machine doesn't have python3 or pip3 installed, [install](https://www.youtube.com/watch?v=gNlhFUY_SWU) them.

#### Step 1. Install necessary dependencies.
```
$ pip3 install pandas openpyxl
```

#### Step 2. Create "faculty_and_staff_short_title.xlsx".
```
$ python3 1_title_converter.py
```
It contains list of short titles ("Visiting Professor", "Instructor", etc.).
It depends on names_titles-13Nov21.xlsx.

#### Step 3. Clean data in EventData folder.
```
$ python3 2_clean_event_data.py
```

#### Step 4. Add short titles to "EventData" and produce "EventDataWithShortTitles" folder.
```
$ python3 3_add_short_title_to_EventData.py
```

#### Step 5. Run the core program that answers user requests.
```
$ python3 4_main.py
```

#### Step 5 (optional). When finished, clean up the created files.
```
$ python3 5_clean_up.py
```

# Additional information

#### In step 2:
- The input excel sheet containing Carleton workers info ("names_titles-13Nov21.xlsx" for our example) should be in the folder "mdrewLTC_main". 
- This excel sheet should include columns titled "Carleton_Name", "Email", and "Primary_Position_Title". 
- If some of these values are missing for certain rows, they won't be included in the report.
- The program will then produce "faculty_and_staff_short_title.xlsx".

#### In step 3: 
- The excel sheets describing events should be in the folder "EventData". 
- These excel sheets should include columns titled "Date", "Email", "Classification".
- Date should be foramtted like 2021-01-31

#### In step 5:
- The script relies on "academic_calendar_2014-26.xlsx" for term start and end dates.