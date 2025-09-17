import definitions

class CalendarGen:
    def __init__(self, start_day, start_month, day_of_week):
        self.start_day = start_day
        self.fetch_iter_day = definitions.weekdays.index(day_of_week) # For iteration in fetch_weekday()
        self.start_month = start_month

    # Returns day of the week using generator => Starting Monday and is in Czech
    def fetch_weekday(self):
        while(True):
            curr_day = definitions.weekdays[self.fetch_iter_day]
        
            if self.fetch_iter_day == 1:
                yield (definitions.REHEARSAL, curr_day)
            elif self.fetch_iter_day == 5 or self.fetch_iter_day == 6:
                yield (definitions.WEEKEND, curr_day)
            else:
                yield (definitions.OTHER, curr_day) 

            # Increase index, wrap around on the end of the week
            self.fetch_iter_day += 1
            if self.fetch_iter_day >= 7:
                self.fetch_iter_day = 0

    
    # Generate a list of tuples (type, day, date) for a single month
    def generate_month(self, month):
        ret = []

        month_idx = month - 1

        month_len = definitions.month_lengths[month_idx]
        day = self.start_day

        for type, weekday in self.fetch_weekday():
            if day > month_len:
                break

            ret.append((type, weekday, day))

            day += 1

        return ret

    # Wrapper for all months generation
    def generate_days(self) -> list[list[str]]:
        ret = []

        for month_idx in range(self.start_month, 13):
            ret.append(self.generate_month(month_idx)) # First start_day differs, then it always begins on one
            self.start_day = 1

        return ret
