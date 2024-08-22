from datetime import datetime, timedelta

def add_10_days(nbr_of_days, date_string, format_string):
    # Convert the date string to a datetime object using the provided format
    date = datetime.strptime(date_string, format_string)

    # Add 10 days to the date
    new_date = date + timedelta(days=nbr_of_days)

    # Convert the new date back to a string using the provided format
    new_date_string = new_date.strftime(format_string)

    # Return the new date string
    return new_date_string
date_string = "07-march-2023"
format_string = "%d-%B-%Y"
nbr_of_days = 100
new_date_string = add_10_days(nbr_of_days,date_string, format_string)
print(new_date_string)  # Output: 2023-07-11
