from datetime import datetime, timedelta

def add_Given_days(nbr_of_days, date_string, format_string):
    # Convert the date string to a datetime object using the provided format
    #date = datetime.strptime(date_string, format_string)
    date = datetime.now()

    # Add 10 days to the date
    new_date = date + timedelta(days=nbr_of_days)

    # Convert the new date back to a string using the provided format
    new_date_string = new_date.strftime(format_string)

    # Return the new date string
    return new_date_string
#date_string = "[%poexpiry]" not needed in this case since we are adding for today
format_string = "%d.%m.%Y"  
nbr_of_days = 3
new_date_string = add_Given_days(nbr_of_days,"-", format_string)
print(new_date_string) 
