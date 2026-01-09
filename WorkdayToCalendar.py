import openpyxl
import sys
from ics import Calendar, Event
from ics.grammar.parse import ContentLine
from datetime import datetime
ogSched = openpyxl.load_workbook('View_My_Courses.xlsx').active
savedSched = []
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
        iniDate = str(row[6])
        finiDate = str(row[7])
        startDate = iniDate[:iniDate.index(" ")].split("-")
        endDate = finiDate[:iniDate.index(" ")].split("-")
        frequency = [(i[:2]).upper() for i in iniSched]
        e.name = Name
        e.begin = datetime(year = int(startDate[0]),month = int(startDate[1]),day = int(startDate[2]), hour = int(startTime/60)+4, minute = int(startTime%60), second=0)
        e.duration = {"minutes": length}
        try:e.description = sections[2]
        except:pass
        e.extra.append(ContentLine(name="RRULE", value = f"FREQ=WEEKLY;BYDAY={",".join(frequency)};UNTIL={"".join(endDate)}T{str(round(endTime/60))}{str(endTime%60)}00"))
        Classes.events.add(e)
with open('View_My_Courses.ics', 'w') as f:

     f.writelines(Classes.serialize_iter())
