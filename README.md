# Send emails to a spreadsheet list using a Python script
This code written in Python sends emails to multiple recipients listed in a spreadsheet.
# Gmail setup
Please follow the steps outlined [here](https://support.google.com/accounts/answer/185833?visit_id=638117389350145242-1419633968&p=InvalidSecondFactor&rd=1) to generate an app password for your Gmail account. 
# How to run the program
1. Open the file spreadsheet_gmail.py in your Python editor of choice.
2. Edit the config.ini file to include your Gmail address, Gmail app password, and email subject. To do this, open the config.ini file in a text editor and replace the following lines: Replace **your_email_address** with your Gmail address, **your_app_password** with your Gmail app password, and **your_email_subject** with the subject line you want to use for your emails. Save the config.ini file when you're done.
3. Make sure that your spreadsheet file **student_emails.xlsx** is in the same directory as **spreadsheet_gmail.py**.
4. Run spreadsheet_gmail.py in your Python editor. The script will read the email addresses and names from the spreadsheet file, and send an email to each address using your Gmail account. The script will also output a file called success_emails.xlsx containing all email addresses where emails were successfully sent. The script will automatically add a delay between 20 and 60 seconds before sending each email to avoid overloading the Gmail server.

**Note**: make sure that you have installed the necessary dependencies (Pandas) before running the script. You can install them using pip by running pip install pandas in your terminal or command prompt.
