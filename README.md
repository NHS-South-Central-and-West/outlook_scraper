# Outlook Scraper Function

For technical assistance, please contact the [Specialist Analytics Team](mailto:scwcsu.analytics.specialist@nhs.net)

## Introduction

## Installation

1. In File Explorer, navigate to the folder where you want to install this project.
2. Right-click in the folder and select "Open in Terminal". You may first need to click on "Show more options".
3. Copy in the instruction ... and press Enter.
4. 

## Instructions

1. Open the outlook_scraper_main.py file. This is where you will use the function and specify the parameters for what 
you want to extract from Outlook.
  - You can optionally copy outlook_scraper_main.py and use the copy, leaving the original as a backup master file. 
  As long as the line `from scraper_function import scraper` stays at the top.
2.Define folder locations where you want to save the extracted files.
  - Since filepaths can be quite long, it is best to assign them to a variable first and then use the variable name 
  when you call the function.
  - The filepaths need to be enclosed with `r' '` and any slashes need to be double backslashes.
  - For example: `output_location_1 = r'O:\\BSS\\BI\\Intelligence Services\\BSSW\\Somerset\\Reports\\outlook_data'`
3. To use the function, you need to write scraper() and in the brackets, you define the parameters of what you want to 
extract from Outlook.
  - For example: 
  `scraper(mailbox_name= 'firstname.lastname@nhs.net', folder= 'Inbox', subject_line= 'Sitrep Data', email_sender= 'someone@SomersetFT.nhs.uk', file_types= ['.xlsx','.csv'], output_location= output_location_1)`
  - You can also call the function without using the parameter names, as long as you keep the parameters in the same order.
  A tool-tip should appear to remind you of the order.
  `scraper('firstname.lastname@nhs.net','Inbox','Sitrep Data','someone@SomersetFT.nhs.uk',['.xlsx','.csv'],output_location_1)`
  - You can call the function as many times as you like. You could define separate output locations for each, or have them all
  point to the same output location.
4. Once the output paths and the parameters of each function call have been defined, save the file.
5. Click on the play button in the top-right corner to run the script.
6. Once the script has finished running, a little report of which relevant e-mails were found will appear in the terminal
pane at the bottom of VS Code.
7. Check the output location(s) for the downloaded files.

## Finer Points. Please read.
- When scraping from personal inboxes, you can use your NHS.net e-mail address. However, with shared inboxes, you need to use
the full name of the inbox, e.g. 'NONPID (NHS SOUTH, CENTRAL AND WEST COMMISSIONING SUPPORT UNIT)'
- Instead of specifying a single sender e-mail, you can set `email_sender` as the domain only, e.g. '@SomersetFt.nhs.uk' and it
will look for any addresses with that domain sending e-mails with the specified subject line. **Caution:** If the subject line
is quite generic and you just specify the e-mail domain, you might pick up more than you were expecting.
- When the script checks the subject line, it uses fuzzy matching, allowing for two errors of insertion, deletion or substitution.
The number of permitted errors can be increased, if needed. This can be done by editing the '2' in the following line in the __scraper_function.py__ file.
`subject_line_regex = f'(?:{subject_line}){{e<=2}}'`
- When specifying the file types that you want to extract, you need to write the list in square brackets and each item in quotes:
`file_types= ['.xlsx','.csv']`
- The function has been set up so that it will look at the latest 200 e-mails. This is to limit how long the script will run when
an e-mail folder contains a large number of e-mails. This should be enough to cover any e-mails received in one day, but if you
do want to increase the number of e-mails to check, go into the __scraper_function.py__ file and edit the value for `max_items`