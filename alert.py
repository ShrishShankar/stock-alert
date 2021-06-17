import os
import gspread
import datetime
import smtplib, ssl
from email.message import EmailMessage
from oauth2client.service_account import ServiceAccountCredentials

scope = [
    "https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"
]

error_msg = ""
try:
    creds = ServiceAccountCredentials.from_json_keyfile_name("stock-alert/creds.json", scope)
except Exception as e:
    error_msg += str(e.__class__) + " (Error) while assigning credentials."

try:
    client = gspread.authorize(creds)
except Exception as e:
    error_msg += str(e.__class__) + " (Error) while authorizing credentials."

try:
    sheet = client.open("stocks")
except Exception as e:
    error_msg += str(e.__class__) + " (Error) while opening desired file."

try:
    core_data = sheet.worksheet("Core Portfolio").get_all_records()
    satellite_data = sheet.worksheet("Satellite").get_all_records()
except Exception as e:
    error_msg += str(e.__class__) + " (Error) while opening desired sheet or reading data."

core_stocks = []
satellite_stocks = []
if error_msg == "":
    for core_row in core_data:
        if isinstance(core_row["SL"], int):
            if core_row['SL'] >= core_row['CMP']:
                core_stocks.append(core_row)
                # plain_msg += "{} has hit stoploss in the Core Portfolio. The current price is {}\n".format(
                #     row1["Company Name"], row1["CMP"])

    for satellite_row in satellite_data:
        if isinstance(satellite_row["SL"], int):
            if satellite_row['SL'] >= satellite_row['CMP']:
                satellite_stocks.append(satellite_row)
                # plain_msg += "{} has hit stoploss in the Satellite Portfolio. The current price is {}\n".format(
                #     row2["Company Name"], row2["CMP"])

plain_msg = ""
table_html = ""
if core_stocks:
    table_html += """\
    <h3>Core Portfolio</h3>
    <table style="border: 1px solid black;>
    <tr style="border: 1px solid black;">
        <th style="border: 1px solid black;" >Company Name</th>
        <th style="border: 1px solid black;" >SL</th>
        <th style="border: 1px solid black;" >CMP</th>
    </tr>"""
    for stock in core_stocks:
        table_html += """\
        <tr>
            <td style="border: 1px solid black;" >{}</td>
            <td style="border: 1px solid black;" >{}</td>
            <td style="border: 1px solid black;" >{}</td>
        </tr>""".format(stock["Company Name"], stock["SL"], stock["CMP"])

        plain_msg += "{} has hit stoploss in the Core Portfolio. The current price is {}\n".format(
            stock["Company Name"], stock["CMP"])

if satellite_stocks:
    table_html += """\
    <h3>Satellite Portfolio</h3>
    <table style="border: 1px solid black;>
    <tr style="border: 1px solid black;">
        <th style="border: 1px solid black;" >Company Name</th>
        <th style="border: 1px solid black;" >SL</th>
        <th style="border: 1px solid black;" >CMP</th>
    </tr>"""
    for stock in satellite_stocks:
        table_html += """\
        <tr>
            <td style="border: 1px solid black;" >{}</td>
            <td style="border: 1px solid black;" >{}</td>
            <td style="border: 1px solid black;" >{}</td>
        </tr>""".format(stock["Company Name"], stock["SL"], stock["CMP"])

        plain_msg += "{} has hit stoploss in the Satellite Portfolio. The current price is {}\n".format(
            stock["Company Name"], stock["CMP"])

html_msg = """\
<!DOCTYPE html>
<html>
    <body>
        <h2>This following stock(s) have hit StopLoss:</h2>
        {}
    </body>
</html>
""".format(table_html)

base = datetime.datetime.today()
base_in_words = base.strftime("%d %b")
base = base.strftime("%d-%m-%Y")
time = datetime.datetime.now().strftime("%-H:%-M:%-S")

# Send Email
port = 465  # For SSL
SENDER = os.environ.get("Script_Mail")
PASSWORD = os.environ.get("Script_Mail_Pass")
email = os.environ.get("Fathers_Mail")
error_email = os.environ.get("Mail")

msg = EmailMessage()
msg['From'] = SENDER

# Create a secure SSL context
context = ssl.create_default_context()

with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
    server.login(SENDER, PASSWORD)
    if plain_msg != "" and table_html != "" and error_msg == "":
        msg['Subject'] = "[IMPORTANT] Stock Alert [at {} on {}]".format(time, base_in_words)
        msg['To'] = error_email
        msg.set_content(plain_msg)
        msg.add_alternative(html_msg, subtype='html')
        server.send_message(msg)

    elif error_msg != "":
        msg['Subject'] = "[ERROR] stock-alert error [at {} on {}]".format(time, base_in_words)
        msg['To'] = error_email
        msg.set_content(error_msg)
        server.send_message(msg)

    else:
        msg = "No results"
        print(msg)
