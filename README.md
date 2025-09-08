# WRKDAYTOICS
Converts the export from Workday into a viable .ics file for google calendar


Issues known and will be fixed:

Lack of timezone support(currently hardcoded to offset to the EDT)

Every class is stacked on the first day of class, even if they aren't in that day's schedule

Lack of location support(I dont even know if this is reliably possible)



Simply place the python file with the .xlsx file and run the program, the ics file will be placed in the same folder after completing
