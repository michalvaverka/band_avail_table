A simple python code used to generate an excel sheet .xlsx to keep track when members are and aren't able to attend a band rehearsal.

Used libraries: xlsxwriter

Description:
    The day generation is simplified and doesn't account for gap years, starting on a different date than Monday, etc...
    But it's enough to provide a functional table for a year and with minimal configuration, it'll be reusable.
    Furthermore I was lazy to go and calculate it in Unix time, since that'd probably make the development longer, than just creating the table by hand. 