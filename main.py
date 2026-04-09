import xlsxwriter

from calendar_generator import CalendarGen
from excel_generator import ExcelGen

import config
from config_perm import (month_names, weekdays)

'''
    Use some calendar algorithm (starting day of week, leap years...)
    perhaps switch to alternative excel AND syntax
'''
class AvailabilityTable:
    def __init__(self):
        self.create_availability_table()

    def create_availability_table(self):
        calendar = CalendarGen()
        excel = ExcelGen()
        month_idx = config.START_MONTH - 1

        # iterate over months
        for day_pack in calendar.generate_days():
            excel.create_worksheet(month_names[month_idx], day_pack)
            month_idx += 1
        excel.close() # save and close the workbook

if __name__ == "__main__":
    # Generates an Availability Table based on config.py
    AvailabilityTable()