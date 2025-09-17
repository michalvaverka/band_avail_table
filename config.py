# ============================ Main configutation ============================ # 

'''
    Starting day of the first sheet (first date to solve availability for).

    Date and the name of the day of the week must match.
'''
START_DAY = 17
START_MONTH = 9
START_DAY_OF_WEEK = "Středa"

'''
    Output filename
'''
FILENAME = "Availability Table.xlsx"

# ============================ Calendar config =============================== # 
''' 
    Gives the sheet names, write out in your language
'''
month_names = ["Leden", "Únor", "Březen", "Duben", "Květen", "Červen", "Červenec", "Srpen", "Září", "Říjen", "Listopad", "Prosinec"]

'''
    Is written into the first column together with the date
''' 
weekdays = ["Pondělí", "Úterý", "Středa", "Čtvrtek", "Pátek", "Sobota", "Neděle"]

'''
    Lengths of months in days, need to take care for leap years
    Hopefully optimized out while implementing more calendar like logic !!!
'''
month_lengths = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

'''
    People who will take part in the scheduling
'''
people = ["Evča", "Mari", "Michal", "Péťa", "Štěpán", "idiot", "debil", "profet", "protifet", "vintir", "zenon", "Všichni"]

# ============================ Various config ================================ #
'''
    Types of days, used for formatting
'''
REHEARSAL = 2
WEEKEND = 1
OTHER = 0
