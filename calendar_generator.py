import config
from config_perm import (month_names, weekdays)

class CalendarGen:
    def __init__(self):
        self.fetch_iter_day = weekdays.index(config.START_DAY_OF_WEEK) # For iteration in fetch_weekday()

    def fetch_weekday(self): 
        '''
            Generator yielding tuples of (type, day) for each day of the week.
            Type is used for formatting, day is used to name the day.    
        '''
        def update_iter_day(self):
            '''
                Updates the day iterator, wraps around.
                It is respective to the days of week.
            '''
            self.fetch_iter_day += 1
            if self.fetch_iter_day >= 7:
                self.fetch_iter_day = 0
        
        # fetch_weekday body here
        while(True):
            match self.fetch_iter_day:
                # Commented out the hardcoded rehearsal day, so it is generally usable
                # case 1:
                #    yield (config.REHEARSAL, weekdays[self.fetch_iter_day])
                case 5 | 6:
                    yield (config.WEEKEND, weekdays[self.fetch_iter_day])
                case _:
                    yield (config.OTHER, weekdays[self.fetch_iter_day]) 
            
            update_iter_day(self)

    def generate_month(self, month, day):
        '''
            Wrapper function, produces a list of tuples (type, day, date) for each day in the given month.
            Type is used for formatting, day is used to name the day, date is the number of the day in the month.
        '''

        ret = []
        month_idx = month - 1
        month_len = config.month_lengths[month_idx]

        for type, weekday in self.fetch_weekday():
            if day > month_len:
                break

            ret.append((type, weekday, day))

            day += 1

        return ret

    def generate_days(self) -> list[list[str]]:
        '''
            Wrapper, that generates all months from START_MONTH to December, with each month being a list of tuples (type, day, date).
        '''
        ret = []
        start_day = config.START_DAY # Starting day of the first month
        
        for month_idx in range(config.START_MONTH, 13):
            ret.append(self.generate_month(month_idx, start_day))
            start_day = 1

        return ret
