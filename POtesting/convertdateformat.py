from datetime import datetime

def convert_date_format(input_date, input_format, output_format):
    try:
        # Parse the input date string into a datetime object using the input format
        datetime_obj = datetime.strptime(input_date, input_format)

        # Convert the datetime object to the desired output format
        output_date = datetime_obj.strftime(output_format)
        return output_date
    except ValueError:
        return "Invalid input date format"


newformat = convert_date_format("2022/12/25","%Y/%m/%d","%d.%m.%Y")

print(newformat)