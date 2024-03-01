#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#MAYBE ADD A SORT BY CBAL DESCENDING
import os
import openpyxl
import datetime
import win32com.client
from openpyxl.utils import column_index_from_string

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
inbox = mapi.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder

# Search for the most recent email with the subject "Daily Cash Balances" and containing an Excel attachment
found_attachment = False
latest_attachment_date = None
latest_attachment_filename = None

messages = inbox.Items
messages.Sort("[ReceivedTime]", True)  # Sort by received time in descending order

for message in messages:
    if message.Subject == "Daily Cash Balances":
        attachments = message.Attachments
        for attachment in attachments:
            if attachment.FileName.lower().endswith(".xlsx"):
                found_attachment = True
                if latest_attachment_date is None or message.ReceivedTime.date() > latest_attachment_date:
                    latest_attachment_date = message.ReceivedTime.date()
                    latest_attachment_filename = attachment.FileName
        if found_attachment:
            break

if not found_attachment:
    print("No suitable attachment found.")
else:
    # Create the folder if it doesn't exist
    folder_path = r"U:\Python\Cash Balances\Limited_Balances"
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    
    # Save the latest attachment to the folder
    latest_attachment_path = os.path.join(folder_path, latest_attachment_filename)
    attachment.SaveAsFile(latest_attachment_path)
    print(f"Latest attachment '{latest_attachment_filename}' downloaded successfully.")

    # Load the Excel workbook
    extracted_workbook = openpyxl.load_workbook(latest_attachment_path)
    extracted_worksheet = extracted_workbook.active  # Assume data is in the active sheet

    filtered_data = []
    for row in extracted_worksheet.iter_rows(min_row=2, values_only=True):
        if row[5] == "L":  # Assuming "POOL" column is F (index 5)
            filtered_data.append(row)

    # Define the folder and filename of the master Excel sheet
    master_folder_path = r"U:\Python\Cash Balances"
    master_filename = "February 2024 Daily Limited Account Cash Balance.xlsx"
    master_file_path = os.path.join(master_folder_path, master_filename)

    # Get the current date and determine the dates for the previous and previous previous days, excluding weekends
    current_date = datetime.date.today()

    if current_date.weekday() == 0:  # If today is Monday, use data from Friday
        previous_date = current_date - datetime.timedelta(days=3)  # Friday
        previous_previous_date = current_date - datetime.timedelta(days=4)  # Thursday
    else:
        previous_date = current_date - datetime.timedelta(days=1)  # Yesterday
        previous_previous_date = current_date - datetime.timedelta(days=2)  # Two days ago

    def format_date_without_leading_zeros(date):
        month = date.strftime("%m").lstrip("0")
        day = date.strftime("%d").lstrip("0")
        year = date.strftime("%Y")
        return f"{month}.{day}.{year}"

    # Get the current date and determine yesterday's date
    current_date = datetime.date.today()
    yesterday_date = current_date - datetime.timedelta(days=1)
    yesterday_date_str = format_date_without_leading_zeros(yesterday_date)

    # Calculate the date for the day before yesterday
    previous_previous_date = current_date - datetime.timedelta(days=2)

    # If the previous previous day falls on a weekend, use the most recent weekday instead
    if previous_previous_date.weekday() >= 5:  # Saturday or Sunday
    # Find the most recent weekday
        previous_previous_date -= datetime.timedelta(days=previous_previous_date.weekday() - 4)

    previous_previous_day_str = format_date_without_leading_zeros(previous_previous_date)

    # Load the master Excel workbook
    master_workbook = openpyxl.load_workbook(master_file_path)

    # Check if the worksheet exists, if not, create a new one
    if yesterday_date_str not in master_workbook.sheetnames:
        master_workbook.create_sheet(yesterday_date_str)

    # Select the worksheet to work with
    current_worksheet = master_workbook[yesterday_date_str]

    # Clear the contents of columns A to G in the master sheet, while preserving the rest
    for row in current_worksheet.iter_rows(min_row=2, max_row=current_worksheet.max_row, min_col=1, max_col=7):
        for cell in row:
            cell.value = None

    # Append the filtered data to the master sheet
    for i, row_data in enumerate(filtered_data):
        for j, value in enumerate(row_data):
            current_worksheet.cell(row=i + 2, column=j + 1, value=value)

    # Define the formula to update for column J
    updated_formula_J = f"=D{{row}} - INDEX('{previous_previous_day_str}'!$D:$D, MATCH('{yesterday_date_str}'!A{{row}}, '{previous_previous_day_str}'!$A:$A, 0))"

    # Update the formula in column J starting from J2
    column_index_J = column_index_from_string('J')  # Get the column index for column J
    for row in range(2, current_worksheet.max_row + 1):
        updated_formula_row_J = updated_formula_J.format(row=row)
        current_worksheet.cell(row=row, column=column_index_J, value=updated_formula_row_J)
    # Define the formula for cell Q11
    formula_Q11 = f"=Q10-'{previous_previous_day_str}'!Q10"

    # Update cell Q11 with the formula
    current_worksheet['Q11'] = formula_Q11

    # Save the master Excel sheet
    master_workbook.save(master_file_path)

#If you open this code after me then I am sorry about how disorganized I am.
#Last accesesed 2/27/2024

