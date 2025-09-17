import xlsxwriter
import config
from string import ascii_uppercase

class ExcelGen:
    def __init__(self):
        
        # Create a workbook
        self.workbook = xlsxwriter.Workbook(config.FILENAME)

        # Define all the formats
        self.fmt_classic = self.workbook.add_format()
        self.fmt_names = self.workbook.add_format()
        self.fmt_weekends = self.workbook.add_format()
        self.fmt_rehearsals = self.workbook.add_format()
        self.fmt_row_odd = self.workbook.add_format()
        self.fmt_row_even = self.workbook.add_format()
        self.fmt_no_format = self.workbook.add_format()

        # Fill them out
        self.specify_formats()
    
    def specify_formats(self):
        '''
            Specifies all used formats.
            
            Feel free to modify colors, borders, alignments ...
        '''

        self.fmt_classic.set_align('center')
        self.fmt_classic.set_border(2)

        self.fmt_names.set_align('center')
        self.fmt_names.set_border(2)
        self.fmt_names.set_bg_color('#339966') # swamp green

        self.fmt_weekends.set_align('center')
        self.fmt_weekends.set_border(2)
        self.fmt_weekends.set_bg_color('#00CCFF') # cyanish blue

        self.fmt_rehearsals.set_align('center')
        self.fmt_rehearsals.set_border(2)
        self.fmt_rehearsals.set_bg_color('#339966') # swamp green

        self.fmt_row_odd.set_align('center')
        self.fmt_row_odd.set_border(2)
        self.fmt_row_odd.set_bg_color('#FFFFFF') # white

        self.fmt_row_even.set_align('center')
        self.fmt_row_even.set_border(2)
        self.fmt_row_even.set_bg_color('#CCFFCC') # ligth green

    def prepare_if_statement(self, row_num):
        '''
            Produces a statement, which checks if all specified people are available on a given day (empty cell).

            Gives a result in the form of filling of the last cell within the row.

            Completely dynamic, works based on size of definitions.people
        '''

        # Start and end of the IF statement
        ret = '=IF('
        end = ', "MOŽME", "Nemožme :(")'
        
        for char in ascii_uppercase[1:len(config.people)]:             # A is for days, so skip it 
            if ascii_uppercase.index(char) == len(config.people) - 1:  # Last person, must be without AND
                ret += 'Len(' + char + str(row_num) + ')=0'                 # Check if the single last cell is empty
            else:
                ret += 'AND(Len(' + char + str(row_num) + ')=0, '           # Check if the cell is empty and leave room for the next one

        ret += (len(config.people) - 2) * ')' + end                    # Close all the ANDs and add the end of the statement

        return ret
        
    # create a single worksheet, the logic is to have a single worksheet per month
    def create_worksheet(self, month_name, day_pack):
        '''
            Create a worksheet and fill it out.

            !!! Try splitting into multiple functions for better readability !!!
        '''
        def add_people(worksheet):
            '''
                Fill the people names into the first row
            '''
            for i in range(1, len(config.people) + 1):
                worksheet.write(0, i, config.people[i-1], self.fmt_names)
    
        def get_format(type, format_row):
            '''
                !!!
            '''
            ret = self.fmt_no_format
        
            if type == config.REHEARSAL:
                ret = self.fmt_rehearsals
            elif type == config.WEEKEND:
                ret = self.fmt_weekends
            else:
                # alternating between styles for better orientation
                if format_row % 2 == 0:
                    ret = self.fmt_row_even
                else:
                    ret = self.fmt_row_even # to enable row format alternating, switch this to row_odd
                format_row += 1
        
            return ret
        
        '''

        '''
    
        worksheet = self.workbook.add_worksheet(month_name)

        # filling in band member's names
        add_people(worksheet)

        curr_row = 1
        format_row = 1

        for type, day, number in day_pack:
            # select format based on the day type
            curr_format = get_format(type, format_row)

            # writing the day info
            worksheet.write(curr_row, 0, day + ' ' + str(number) + '.', curr_format)    # write day and date into the first column
            worksheet.write(curr_row, len(config.people), self.prepare_if_statement(curr_row + 1), curr_format)  # write command statement into the last column

            # filling the rest of the row with empty data to apply the format
            for i in range(1, len(config.people)):
                worksheet.write(curr_row, i, '', curr_format)
        
            curr_row += 1

        # EDIT Perhaps this doesnt work as intended
        # setting the width for cells to fit all the text
        worksheet.set_column(0, 6, 15)

        return worksheet
    
    def close(self):
        self.workbook.close()