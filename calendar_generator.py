import config

class CalendarGen:
    def __init__(self):
        self.fetch_iter_day = config.weekdays.index(config.START_DAY_OF_WEEK) # For iteration in fetch_weekday()

    # Returns day of the week using generator => Starting Monday and is in Czech
    def fetch_weekday(self):
        while(True):
            curr_day = config.weekdays[self.fetch_iter_day]
        
            if self.fetch_iter_day == 1:
                yield (config.REHEARSAL, curr_day)
            elif self.fetch_iter_day == 5 or self.fetch_iter_day == 6:
                yield (config.WEEKEND, curr_day)
            else:
                yield (config.OTHER, curr_day) 

            # Increase index, wrap around on the end of the week
            self.fetch_iter_day += 1
            if self.fetch_iter_day >= 7:
                self.fetch_iter_day = 0

    
    # Generate a list of tuples (type, day, date) for a single month
    def generate_month(self, month):
        ret = []

        month_idx = month - 1

        month_len = config.month_lengths[month_idx]
        day = config.START_DAY

        for type, weekday in self.fetch_weekday():
            if day > month_len:
                break

            ret.append((type, weekday, day))

            day += 1

        return ret

    # Wrapper for all months generation
    def generate_days(self) -> list[list[str]]:
        ret = []

        for month_idx in range(config.START_MONTH, 13):
            ret.append(self.generate_month(month_idx)) # First start_day differs, then it always begins on one
            self.start_day = 1

        return ret
