import xlsxwriter

weekdays = ["Pondělí", "Úterý", "Středa", "Čtvrtek", "Pátek", "Sobota", "Neděle"]

month_lengths = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
month_names = ["Leden", "Únor", "Březen", "Duben", "Květen", "Červen", "Červenec", "Srpen", "Září", "Říjen", "Listopad", "Prosinec"]

band_members = ["Evča", "Mari", "Michal", "Petr", "Štěpán", "Všichni"]

REHEARSAL = 2
WEEKEND = 1
OTHER = 0
# barevne vyznacene vikendy a uterky

# Returns day of the week using generator => Starting Monday and is in Czech
def fetch_weekday():
    while(True):
        curr_day = weekdays[fetch_weekday.fetch_idx]
        
        if fetch_weekday.fetch_idx == 1:
            yield (REHEARSAL, curr_day)
        elif fetch_weekday.fetch_idx == 5 or fetch_weekday.fetch_idx == 6:
            yield (WEEKEND, curr_day)
        else:
            yield (OTHER, curr_day) 

        fetch_weekday.fetch_idx += 1
        if fetch_weekday.fetch_idx >= 7:
            fetch_weekday.fetch_idx = 0

fetch_weekday.fetch_idx = 0
        
def generate_month(month, start_day):
    ret = []

    month_idx = month - 1

    month_len = month_lengths[month_idx]
    day = start_day

    for weekday in fetch_weekday():
        if day > month_len:
            break

        ret.append((weekday, day))

        day += 1

    return ret

def generate_days(start_day, start_month) -> list[list[str]]:
    ret = []

    for i in range(start_month, 13):
        ret.append(generate_month(i, start_day))
        start_day = 1

    return ret

# Each month will be a separate worksheet, ok?
# I will need to format it, but I can do so globally after
def generate_table_month(days, month):
    month_idx = month - 1

    month_name = month_names[month_idx]

    print(month_name)

    for day in days:
        print('Here, append to the xlsx') # How do I know the month? Hard coded I'd presume
        print(day)


# XLSX Helper functions
wb = xlsxwriter.Workbook('Availability Table.xlsx')

# for everything
classic = wb.add_format()
classic.set_align('center')
classic.set_border(2)

# for names
names = wb.add_format()
names.set_align('center')
names.set_border(2)
names.set_bg_color('#339966') # swamp green

# for weekends
weekends = wb.add_format()
weekends.set_align('center')
weekends.set_border(2)
weekends.set_bg_color('#00CCFF') # cyanish blue

# for regular rehearsals
rehearsals = wb.add_format()
rehearsals.set_align('center')
rehearsals.set_border(2)
rehearsals.set_bg_color('#339966') # swamp green

# iterating formats for rows -> easier editability
row_odd = wb.add_format()
row_odd.set_align('center')
row_odd.set_border(2)
row_odd.set_bg_color('#FFFFFF') # white

row_even = wb.add_format()
row_even.set_align('center')
row_even.set_border(2)
row_even.set_bg_color('#CCFFCC') # ligth green

no_format = wb.add_format()

def prepare_if_statement(row_num):
    return '=IF(AND(Len(B' + str(row_num) + ')=0, AND(Len(C' + str(row_num) + ')=0, AND(Len(D' + str(row_num) + ')=0, AND(Len(E' + str(row_num) + ')=0, Len(F' + str(row_num) + ')=0)))) , "MOŽME", "Nemožme :(")'

def create_worksheet(month_name, day_pack):
    ws = wb.add_worksheet(month_name)

    # filling in band member's names
    for i in range(1, 7):
        ws.write(0, i, band_members[i-1], names)

    curr_row = 1
    format_row = 1

    curr_format = no_format

    for day, number in day_pack:
        if day[0] == REHEARSAL:
            curr_format = rehearsals
        elif day[0] == WEEKEND:
            curr_format = weekends
        else:
            # alternating between styles for better orientation
            if format_row % 2 == 0:
                curr_format = row_even
            else:
                curr_format = row_even
            format_row += 1

        # writing the day info into the first cell
        ws.write(curr_row, 0, day[1] + ' ' + str(number) + '.', curr_format)
        ws.write(curr_row, 6, prepare_if_statement(curr_row + 1), curr_format)

        # filling the body with no information just to format it
        for i in range(1, 6):
            ws.write(curr_row, i, '', curr_format)
        
        curr_row += 1

    # setting the width for cells
    ws.set_column(0, 6, 15)

    return ws

if __name__ == '__main__':
    
    START_DAY = 13
    START_MONTH = 5

    # First day of generation has to be a monday
    days = generate_days(START_DAY, START_MONTH)

    # print(days)
    
    month_idx = START_MONTH - 1

    curr_row = 1
    curr_col = 0

    for day_pack in days:
        ws = create_worksheet(month_names[month_idx], day_pack)
        
        month_idx += 1

    wb.close()
