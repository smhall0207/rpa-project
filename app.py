# IMPORT PACKAGES IN THIS CELL

import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# INSERT ANSWER FOR QUESTION 1 HERE
while True: 
    directory = single_input("Please give the directory where the needed files are stored.")
    if directory[-1] != "\\":
        directory = directory + "\\"
    try:
        pd.read_excel(directory + "ProductData.xlsx")
        get_pdf(directory + "MasterInvoice.pdf")
        pd.read_excel(directory + "Order1.xlsx")
        pd.read_excel(directory + "Order2.xlsx")
        pd.read_excel(directory + "Order3.xlsx")
        pd.read_excel(directory + "Order4.xlsx")
        break
    except:
        show_message("Error", "Files not found. Reenter directory.")

# INSERT ANSWER FOR QUESTION 2 HERE

# Dictionary Formating --> {Unit number: Price per unit or Description}
product_descriptions = {}
product_prices = {}

descriptions_df = pd.read_excel(directory + "ProductData.xlsx", sheet_name="Descriptions")
for index, row in descriptions_df.iterrows():
    product_descriptions[row["Unit number"]] = row["Description"]

prices_df = pd.read_excel(directory + "ProductData.xlsx", sheet_name="Prices")
for index, row in prices_df.iterrows():
    product_prices[row["Unit number"]] = row["Price per unit"]


# INSERT ANSWER FOR QUESTION 3 HERE
import datetime as dt

def get_order_data(invoice_num, filename, directory):
    update_dict = {}
    order = openpyxl.load_workbook(directory + filename, data_only=True)
    sheet = order["OrderForm"]

    update_dict["InvoiceNum"] = str(invoice_num)
    update_dict["Email"] = str(sheet["C13"].value)

    month = str(sheet["C9"].value).split(" ")[0].split("-")[1]
    day = str(sheet["C9"].value).split(" ")[0].split("-")[2]
    year = str(sheet["C9"].value).split(" ")[0].split("-")[0]
    update_dict["Date"] = f"{month}/{day}/{year}"

    update_dict["CustomerName"] = str(sheet["C14"].value)
    update_dict["CustomerCompany"] = str(sheet["C15"].value)
    update_dict["CustomerAddress1"] = str(sheet["C16"].value)
    update_dict["CustomerAddress2"] = f"{sheet["C17"].value}, {sheet["C18"].value} {sheet["C19"].value}"

    num = 24
    start = 0
    total = 0
    while True:
        if sheet["B" + str(num)].value is None:
            break
        else:
            update_dict["Description" + str(start)] = str(product_descriptions[sheet["B" + str(num)].value])
            update_dict["Quantity" + str(start)] = str(sheet["C" + str(num)].value)
            update_dict["Price" + str(start)] = str(product_prices[sheet["B" + str(num)].value])
            initial_amount = (sheet["C" + str(num)].value * product_prices[sheet["B" + str(num)].value])
            update_dict["Amount" + str(start)] = f'{initial_amount:.2f}'

            total = total + initial_amount
            start = start + 1
            num = num + 1

    update_dict["Total"] = f"{total:.2f}"
    name = update_dict["CustomerName"]
    email = update_dict["Email"]
    date = update_dict["Date"]

    return update_dict, name, email, date

# Code below for identifying pdf variables
# test = get_pdf("MasterInvoice.pdf")
# test

order_data, name, email, date = get_order_data(8515, "Order1.xlsx", directory)


# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + "MasterInvoice.pdf", directory + f"Invoice{invoice_num}.pdf", order_data)
    return

create_invoice(8515, order_data, directory)


# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):
    gmail_user = "smhall0207@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
Attached is your invoice dated {date}.

Thank you for your business.
        
PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    message = add_attachment(directory + f"Invoice{invoice_num}.pdf", message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

    return

send_email(8515, "Thomas Holt", "smhall0207@gmail.com", "11/13/2023", directory)


# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
invoice_num = 8515

for i, file in enumerate(filenames):
    if "Order" in file and "xlsx" in file:
        order_data, name, email, date = get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1