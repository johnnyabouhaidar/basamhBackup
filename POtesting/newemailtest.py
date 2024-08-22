import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import numpy as np  # Import numpy for handling NaN values


def get_first_value(dataframe):
    if isinstance(dataframe, pd.DataFrame) and not dataframe.empty:
        first_column = dataframe.iloc[:, 0]  # Select the first column
        first_value = first_column.iloc[0]  # Get the first value of the first column
        return first_value
    else:
        return None

def dynamic_inner_join_excel(input_file1, input_file2):
    try:
        # Read the Excel files into pandas DataFrames
        df1 = pd.read_excel(input_file1)
        df2 = pd.read_excel(input_file2)

        # Identify common columns by taking the intersection of column names
        common_columns = list(set(df1.columns) & set(df2.columns))

        # Perform an inner join based on common columns
        joined_df = pd.merge(df1, df2, on=common_columns, how='left')

        # Replace NaN values with empty strings in the joined DataFrame
        joined_df.replace(np.nan, '', inplace=True)

        return joined_df
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

# Example usage:
input_file1 = r"C:\Users\rpauser\Documents\itemstoadd.xlsx"
input_file2 = r"C:\Users\rpauser\Documents\supplimentaryitems.xlsx"
joined_df = dynamic_inner_join_excel(input_file1, input_file2)
# Now, you have the joined DataFrame in memory with empty strings instead of NaN values
print(joined_df)

def generate_html_from_dataframe(df):
    header_row = """
    <tr style="height:38.25pt">
        <td nowrap="" valign="top" style="border:solid windowtext 1.0pt;background:silver;padding:0in 5.4pt 0in 5.4pt;height:38.25pt">
<p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif;color:black">Sales Doc.No</span><span style="font-family:&quot;Arial&quot;,sans-serif"><u></u><u></u></span></p>
</td>
<td valign="top" style="border:solid windowtext 1.0pt;border-left:none;background:silver;padding:0in 5.4pt 0in 5.4pt;height:38.25pt">
<p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif;color:black">Sales Document Item</span><span style="font-family:&quot;Arial&quot;,sans-serif"><u></u><u></u></span></p>
</td>
<td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-left:none;background:silver;padding:0in 5.4pt 0in 5.4pt;height:38.25pt">
<p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif;color:black">Material</span><span style="font-family:&quot;Arial&quot;,sans-serif"><u></u><u></u></span></p>
</td>
<td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-left:none;background:silver;padding:0in 5.4pt 0in 5.4pt;height:38.25pt">
<p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif;color:black">Item Description</span><span style="font-family:&quot;Arial&quot;,sans-serif"><u></u><u></u></span></p>
</td>
<td valign="top" style="border:solid windowtext 1.0pt;border-left:none;background:silver;padding:0in 5.4pt 0in 5.4pt;height:38.25pt">
<p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif;color:black">Order Quantity</span><span style="font-family:&quot;Arial&quot;,sans-serif"><u></u><u></u></span></p>
</td>
<td valign="top" style="border:solid windowtext 1.0pt;border-left:none;background:silver;padding:0in 5.4pt 0in 5.4pt;height:38.25pt">
<p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif;color:black">Confirmed Quantity</span><span style="font-family:&quot;Arial&quot;,sans-serif"><u></u><u></u></span></p>
</td>
<td valign="top" style="border:solid windowtext 1.0pt;border-left:none;background:silver;padding:0in 5.4pt 0in 5.4pt;height:38.25pt">
<p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif;color:black">Unconfirmed Qty</span><span style="font-family:&quot;Arial&quot;,sans-serif"><u></u><u></u></span></p>
</td>
<td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-left:none;background:silver;padding:0in 5.4pt 0in 5.4pt;height:38.25pt">
<p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif;color:black">Sales unit</span><span style="font-family:&quot;Arial&quot;,sans-serif"><u></u><u></u></span></p>
</td>
<td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-left:none;background:silver;padding:0in 5.4pt 0in 5.4pt;height:38.25pt">
<p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif;color:black">Usage Indicator</span><span style="font-family:&quot;Arial&quot;,sans-serif"><u></u><u></u></span></p>
</td>
<td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-left:none;background:#ffc000;padding:0in 5.4pt 0in 5.4pt;height:38.25pt">
<p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif;color:black">Remark
</span><span style="font-family:&quot;Arial&quot;,sans-serif"><u></u><u></u></span></p>
</td>
<td valign="top" style="border:solid windowtext 1.0pt;border-left:none;background:#ffc000;padding:0in 5.4pt 0in 5.4pt;height:38.25pt">
<p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif;color:black">Customer SKU
</span><span style="font-family:&quot;Arial&quot;,sans-serif"><u></u><u></u></span></p>
</td>
    </tr>
    """

    # Initialize the HTML output with the header row
    html_output = f"""<!DOCTYPE html>
<html>
<head>
    <title>Embedded HTML</title>
</head>
<body>
    <table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:none">
        <tbody>
            {header_row}
    """

    # Iterate through the DataFrame and generate HTML rows
    for index, row in df.iterrows():
        row_html = f"<tr style=\"height:12.75pt\">"
        for col in df.columns:
            cell_value = row[col]
            row_html += f"""
                <td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt;height:12.75pt">
                    <p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif">{cell_value}<u></u><u></u></span></p>
                </td>
            """
        row_html += "</tr>"
        html_output += row_html

    # Close the HTML table and body
    html_output += """
        </tbody>
    </table>
    <p>Note that:</p>
<p>For DUMMY material quantity shows as PC the same as the PO, please change the item and update the master data.</p>
<p>For deferent price please check the prices and update the customer master data.</p>
<p>For items that not defined as TR please check the QTY and contact customer service team to update the material details.</p>
<p>For duplication entry the same item with same PO was created in other sales order please check with Customer Service Team.</p>
<p></p>
<p>Regards.</p>
<p>Kindly if you have any questions or help, please contact Customer Service Team throw https://itcare.basamh.com/app/btccustomerserviceportal/.</p>
<p>Please do not reply to this email. This mailbox is not monitored, and you will not receive a response.</p>
</body>
</html>
    """

    return html_output


html_output = generate_html_from_dataframe(joined_df)
print(html_output)
def send_email(sender_email, sender_password, recipient_email, subject, html_content):
    # Create a multipart message object
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Attach the HTML content to the email
    msg.attach(MIMEText(html_content, 'html'))

    try:
        # Create a secure connection to the email server
        server = smtplib.SMTP('smtp.office365.com', 587)

        # Start the TLS encryption
        server.starttls()

        # Login to the sender's email account
        server.login(sender_email, sender_password)

        # Send the email
        server.sendmail(sender_email, recipient_email, msg.as_string())

        print("Email sent successfully!")
    except Exception as e:
        print("Error sending email:", str(e))
    finally:
        # Close the connection to the email server if it's open
        if 'server' in locals():
            server.quit()

# Example usage
sender_email = "rpauser@basamh.com"
sender_password = "A11owMe-753"
recipient_email = "johnnyabuhaydar@gmail.com"

doc_no = get_first_value(joined_df)
subject ="[%POvendorcurrent] [%shipto] - PO# [%ponumber]– {0}".format(str(doc_no))

# Define the HTML content

send_email(sender_email, sender_password, recipient_email, subject, html_output)
