import xlsxwriter
from calendar_generator import CalendarGen
from excel_generator import ExcelGen
import definitions

'''
    Move all definitions, rename to configuration
    Use some calendar algorithm (starting day of week, leap years...)
    perhaps switch to alternative excel AND syntax
    Add comments
'''
if __name__ == '__main__':
    
    START_DAY = 17
    START_MONTH = 9

    calendar = CalendarGen(START_DAY, START_MONTH, "Středa")
    excel = ExcelGen('Availability Table.xlsx')
    
    month_idx = START_MONTH - 1

    curr_row = 1
    curr_col = 0

    # iterate over months
    for day_pack in calendar.generate_days():
        ws = excel.create_worksheet(definitions.month_names[month_idx], day_pack)
        
        month_idx += 1

    excel.close() # save and close the workbook
