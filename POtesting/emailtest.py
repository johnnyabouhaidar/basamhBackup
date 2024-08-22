import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def parse_and_generate_html(input_string):
    rows = input_string.split("__")

    output = []
    for row in rows:
        columns = row.split("|")
        variables = [column.strip() for column in columns]

        if len(variables) == 11:
            output.append(
                f"""
                <tr style="height:12.75pt">
                    <td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt;height:12.75pt">
                        <p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif">{variables[0]}<u></u><u></u></span></p>
                    </td>
                    <td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt;height:12.75pt">
                        <p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif">{variables[1]}<u></u><u></u></span></p>
                    </td>
                    <td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt;height:12.75pt">
                        <p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif">{variables[2]}<u></u><u></u></span></p>
                    </td>
                    <td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt;height:12.75pt">
                        <p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif">{variables[3]}<u></u><u></u></span></p>
                    </td>
                    <td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt;height:12.75pt">
                        <p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif">{variables[4]}<u></u><u></u></span></p>
                    </td>
                    <td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt;height:12.75pt">
                        <p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif">{variables[5]}<u></u><u></u></span></p>
                    </td>
                    <td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt;height:12.75pt">
                        <p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif">{variables[6]}<u></u><u></u></span></p>
                    </td>
                    <td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt;height:12.75pt">
                        <p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif">{variables[7]}<u></u><u></u></span></p>
                    </td>
                    <td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt;height:12.75pt">
                        <p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif">{variables[8]}<u></u><u></u></span></p>
                    </td>
                    <td nowrap="" valign="top" style="border:solid windowtext 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt;height:12.75pt">
                        <p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif">{variables[9]}<u></u><u></u></span></p>
                    </td>
                    <td nowrap="" valign="top" style="border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0in 5.4pt 0in 5.4pt;height:12.75pt">
                        <p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif">{variables[10]}<u></u>&nbsp;<u></u></span></p>
                    </td>
                </tr>
                """
            )

    return f"""<!DOCTYPE html>
<html>
<head>
    <title>Embedded HTML</title>
</head>
<body>
    <table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:none">
        <tbody>
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
<td valign="top" style="border:solid windowtext 1.0pt;border-left:none;background:#ffc000;padding:0in 5.4pt 0in 5.4pt;height:38.25pt">
<p class="MsoNormal"><span style="font-family:&quot;Arial&quot;,sans-serif;color:black">Customer SKU
</span><span style="font-family:&quot;Arial&quot;,sans-serif"><u></u><u></u></span></p>
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
</tr>
            {"".join(output)}
        </tbody>
    </table>
</body>
</html>"""


input_string = "2100509440|10|10000549| |GDY-MILK SWT CONDENSED-12X395G|10|10|0|CAR| |__2100509440|10|10000549| |GDY-MILK SWT CONDENSED-12X395G|10|10|0|CAR| |"
html_output = parse_and_generate_html(input_string)
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
sender_password = "Password1"
recipient_email = "johnnyabuhaydar@gmail.com"
subject = "HTML Email Example"

# Define the HTML content

send_email(sender_email, sender_password, recipient_email, subject, html_output)
