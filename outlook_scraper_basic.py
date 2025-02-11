# building it up layered

######################################################################################################
'''
Setup - DO NOT CHANGE when running this code.
'''
import pandas as pd
from datetime import datetime
import win32com.client
import os
import re

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
date_today = datetime.today().strftime('%Y-%m-%d') # to add to file names
max_items = 100 # set the max number of e-mails to look through

######################################################################################################
'''
User-Defined Controls
'''
######################################################################################################

# Enter the name of the mailbox address you want to extract from (xxxx@nhs.net)
mailbox_name = 'NONPID (NHS SOUTH, CENTRAL AND WEST COMMISSIONING SUPPORT UNIT)'

# Enter the name of the mailbox folder you want to extract from (e.g. 'Inbox')
folder = 'Inbox'

# Enter the name of the directory folder you want to export the files to (O: drive etc.)
# 'output/' will export files to a folder called 'output' within your current working directory.
# output_folder = r'C:\\Users\\edward.chick\\OneDrive - NHS\Documents\\my_dev\\sat_requests\\sat0008_outlook_scraping\\output'

file_types = ['.xlsx'] # to avoid downloading images in peoples' signatures.

# Enter the date range within which you want to search the mailbox (YYYY-MM-DD)
start_date = '2025-02-10'
end_date = '2025-02-11'

# Enter the subject line of the e-mails that you want to search for. Don't add RE: or FW: 
subject_line = 'Somerset Hub Daily Bed State'

# Enter the address of the sender (e.g.xxxx@SomersetFT.nhs.uk)
email_sender = r'@somersetft\.nhs\.uk$'

######################################################################################################
'''
Process Script - DO NOT CHANGE when running this code.
'''
######################################################################################################

# Identify mailbox folder from user inputs
mailbox_folder_to_query = outlook.Folders[mailbox_name].Folders[folder]

# Create list of items in the inbox and sort them from newest to oldest
messages = mailbox_folder_to_query.Items
messages.Sort("[ReceivedTime]", True)

# Convert user input dates to datetime for later comparison
start_date = pd.to_datetime(start_date)
end_date = pd.to_datetime(end_date)

# Create empty lists to receive message extract information for report
# Less computationally expensive than appending to a DataFrame row by row
received_date = []
subject = []
sender = []
attachments = []

##### Main loop #####

processing_count = 0

for message in messages:
    if processing_count >= max_items:
        break
    try:
        if message.Class == 43:     # Mail Items only (not calendar invitations/notifications)

            processing_count +=1    # For every Mail Item, add to the count tally

            # Retrieve date e-mail was sent, as long as within defined date range
            email_date =  pd.to_datetime(message.senton.date())  
            if not (start_date <= email_date <= end_date):
                continue

            received_date.append(message.senton.date())

            # Retrieve subject line
            if subject_line is not None:
                if subject_line.lower() not in message.Subject.lower():
                    continue

            subject.append(message.Subject)

            # Retrieve sender, handling Exchange Server addresses and distribution lists
            if message.Class == 43:
                if message.SenderEmailType == 'EX':
                    if message.Sender.GetExchangeUser() != None:
                        sender_check = message.Sender.GetExchangeUser().PrimarySmtpAddress
                    else:
                        sender_check = message.GetExchangeDistributionList().PrimarySmtpAddress
                else:
                    sender_check = message.SenderEmailAddress

            if email_sender is not None:
                if not re.search(email_sender.lower(), sender_check.lower()):
                    continue
                else:
                    sender.append(sender_check)

            # Retrieve attachments where they match the file types specified by the user
            attachment_list = []
            for attachment in message.Attachments:
                file_name = attachment.FileName
                file_extension = os.path.splitext(file_name)[1].lower() # get just the extension from the FileName              

                if file_extension in [ext.lower() for ext in file_types]:
                    file_name_updated = f'{os.path.splitext(file_name)[0]}_{email_date.strftime('%Y-%m-%d')}{file_extension}'
                    file_path = f'{os.getcwd()}\\output\\{file_name_updated}'
                    attachment.SaveAsFile(file_path)
                    attachment_list.append(attachment.FileName)

            attachments.append(attachment_list)

        else:
            continue # skip anything that isn't a Mail Item

    except Exception as e:
        print(f'Error processing e-mail: {e}')


report = pd.DataFrame(list(zip(received_date,subject,sender,attachments)),
                      columns=['Received','Subject','Sender','Attachments']
                      )

if report.empty:
    print('No relevant e-mails found')
else:
    print(report)
