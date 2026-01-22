from icalendar import Calendar, Event, vDatetime, vRecur
import pytz  # timezone libraries
from datetime import datetime, timedelta, time
import openpyxl
from pathlib import Path

ogSched = openpyxl.load_workbook('View_My_Courses.xlsx').active
savedSched = []
tz = pytz.timezone("US/Eastern")

for row in ogSched.rows:
    savedSched.append([cell.value for cell in row])
savedSched = [x[6: len(savedSched)] for x in savedSched]
savedSched = savedSched[3:len(savedSched)]
for row in range(len(savedSched)):
        for col in range(len(savedSched[0])):
            try:savedSched[row][col] = savedSched[row][col].replace("\n", " ")
            except:pass
def timeValue(x):
    splitTime = x[:x.find(" ")].split(":")
    totalTime = int(splitTime[0])*60 + int(splitTime[1])
    if "PM" in x :
        totalTime = totalTime + 60*12
    if splitTime[0] == "12":
        totalTime =  totalTime - 60*12 
    return totalTime

Classes = Calendar()
for row in savedSched:
        e = Event()
        Name = row[0]
        sections = row[4].split(" | ")
        iniSched = sections[0].split("/")
        iniTimes = sections[1].split(" - ")
        startTime = timeValue(iniTimes[0])
        endTime  = timeValue(iniTimes[1])
        length = endTime - startTime
        startTimeDT = time(int(startTime/60), int(startTime%60))
        print(startTimeDT)
        iniDate = (row[6])
        finiDate = (row[7])
        frequency = [(i[:2]).upper() for i in iniSched]
        e.add("SUMMARY", Name)
        e.DTSTART = datetime.combine(iniDate, startTimeDT)
        e.DURATION = timedelta(minutes= length)
        try:e.description = sections[2]
        except:pass
        e.add("RRULE", vRecur({"FREQ":["WEEKLY"], "BYDAY":frequency}, until = finiDate))
        Classes.add_component(e)
        e.DTSTAMP = datetime.now(tz)
Classes.add_missing_timezones()
print((Classes.to_ical()))
Path("Exported Courses.ics").write_bytes(Classes.to_ical())
