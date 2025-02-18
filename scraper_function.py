# scraper function

import pandas as pd
from datetime import datetime
import win32com.client
import os
import regex as re

def scraper(mailbox_name: str,folder: str,subject_line: str,email_sender: str,file_types: list[str],
            output_location: str):
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    date_today = datetime.today().strftime('%Y-%m-%d') # for date comparison and to add to filenames.
    max_items = 200 # set the max number of e-mails to look through
    
    # Identify mailbox folder from user inputs
    mailbox_folder_to_query = outlook.Folders[mailbox_name].Folders[folder]

    # Create list of items in the inbox and sort them from newest to oldest
    messages = mailbox_folder_to_query.Items
    messages.Sort("[ReceivedTime]", True)

    # Convert user input dates to datetime for later comparison
    # start_date = pd.to_datetime(start_date)
    # end_date = pd.to_datetime(end_date)

    # Convert the user input subject line into fuzzy-matching regex
    subject_line_regex = f'(?:{subject_line}){{e<=2}}'

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
                if email_date.strftime('%Y-%m-%d') != date_today:
                    continue

                received_date.append(message.senton.date())

                # Retrieve subject line
                if subject_line is not None:
                    if not re.match(subject_line_regex.lower(), message.Subject.replace("RE: ","").replace("FW: ","").lower()):
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
                        file_path = f'{output_location}\\{file_name_updated}'
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