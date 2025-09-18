# ============================ Main configutation ============================ # 

'''
    Starting day of the first sheet (first date to solve availability for).

    Date and the name of the day of the week must match.
'''
START_DAY = 17              # !!! MAKE OPTIONAL !!!
START_MONTH = 9
START_DAY_OF_WEEK = "Středa" # !!! OPTIMIZE OUT !!!

'''
    Output filename
'''
FILENAME = "Availability Table.xlsx"

# ============================ Calendar config =============================== # 
''' 
    Gives the sheet names, write out in your language 
'''
# TRY OPTIMIZING OUT !!!
month_names = ["Leden", "Únor", "Březen", "Duben", "Květen", "Červen", "Červenec", "Srpen", "Září", "Říjen", "Listopad", "Prosinec"]

'''
    Is written into the first column together with the date
''' 
# TRY OPTIMIZING OUT OR AT LEAST MAKE IT JUST A LANGUAGE CHOICE !!!
weekdays = ["Pondělí", "Úterý", "Středa", "Čtvrtek", "Pátek", "Sobota", "Neděle"]

'''
    Lengths of months in days, need to take care for leap years
    Hopefully optimized out while implementing more calendar like logic !!!
'''
# TRY OPTIMIZING OUT !!!
month_lengths = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

'''
    People who will take part in the scheduling
'''
people = ["Evča", "Mari", "Michal", "Péťa", "Štěpán", "Všichni"]

# ============================ Various config ================================ #
'''
    Types of days, used for formatting
'''
# TRY MAKE MORE GENERAL, or at the very least make REHEARSAL something like SCHEDULED DAY with option to which day it is. 
REHEARSAL = 2
WEEKEND = 1
OTHER = 0
