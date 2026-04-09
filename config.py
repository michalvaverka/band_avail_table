# ============================ Main configutation ============================ # 

'''
    Starting day of the first sheet (first date to solve availability for).

    Date and the name of the day of the week must match.
'''
START_DAY = 1              # !!! MAKE OPTIONAL !!!
START_MONTH = 1
START_DAY_OF_WEEK = "Čtvrtek" # !!! OPTIMIZE OUT !!!

'''
    Output filename
'''
FILENAME = "Availability Table.xlsx"

# ============================ Calendar config =============================== # 

'''
    Lengths of months in days, need to take care for leap years
    Hopefully optimized out while implementing more calendar like logic !!!
'''
# TRY OPTIMIZING OUT !!!
month_lengths = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

'''
    People who will take part in the scheduling
'''
people = ["1", "2", "3", "4", "5", "6", "Všichni"]

'''
    Types of days, used for formatting
'''
# TRY MAKE MORE GENERAL, or at the very least make REHEARSAL something like SCHEDULED DAY with option to which day it is. 
REHEARSAL = 2
WEEKEND = 1
OTHER = 0
